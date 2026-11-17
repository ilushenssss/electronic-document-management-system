"""Предпросмотр страниц Word: фоновые задания, очередь и масштабирование."""
from collections import OrderedDict
import copy
from pathlib import Path
import shutil
import tempfile
import weakref

from PyQt5.QtCore import QEvent, QObject, QProcess, QProcessEnvironment, QTimer, Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import (QApplication, QComboBox, QHBoxLayout, QLabel, QScrollArea,
                             QLayout, QVBoxLayout, QWidget)

from .. import docx, word_preview


class _Renderer(QObject):
    """Один экземпляр Word за раз для всех экранов приложения."""
    def __init__(self, app):
        super().__init__(app)
        self.pending = OrderedDict()
        self.cache = OrderedDict()
        self.active = None
        self.stopping = False
        self.deadline = QTimer(self)
        self.deadline.setSingleShot(True)
        self.deadline.timeout.connect(self.timeout)
        app.aboutToQuit.connect(self.shutdown)

    def request(self, view, digest, data, force=False):
        if self.stopping:
            return
        self.pending.pop(id(view), None)
        if (not force and self.active and self.active['digest'] == digest
                and self.active['view']() is view):
            return
        if not force and digest in self.cache:
            directory, pages = self.cache[digest]
            if all(path.is_file() for path in pages):
                self.cache.move_to_end(digest)
                view.completed(digest, pages, None)
                return
        self.pending[id(view)] = (weakref.ref(view), digest, data, force)
        self.start_next()

    def start_next(self):
        if self.active or self.stopping:
            return
        while self.pending:
            _, (reference, digest, data, force) = self.pending.popitem(last=False)
            view = reference()
            if view is None or not view.isVisible() or view.digest != digest:
                continue
            if not force and digest in self.cache:
                _, pages = self.cache[digest]
                if all(path.is_file() for path in pages):
                    self.cache.move_to_end(digest)
                    view.completed(digest, pages, None)
                    continue
            job = {'view': reference, 'digest': digest, 'process': None, 'directory': None}
            self.active = job
            try:
                command = word_preview.powershell_command()
                directory = Path(tempfile.mkdtemp(prefix='sz-preview-'))
                job['directory'] = directory
                (directory / 'input.docx').write_bytes(data)
                process = QProcess(self)
                job['process'] = process
                env = QProcessEnvironment.systemEnvironment()
                env.insert('SZ_PREVIEW_DIRECTORY', str(directory))
                process.setProcessEnvironment(env)
                process.setWorkingDirectory(str(directory))
                process.finished.connect(lambda code, status, j=job: self.finish(j))
                process.errorOccurred.connect(lambda error, j=job: self.finish(j, 'Не удалось запустить предпросмотр Word.'))
                self.deadline.start(90000)
                process.start(command[0], command[1:])
            except Exception as error:
                self.finish(job, str(error))
            return

    def finish(self, job, error=None):
        if self.active is not job:
            return
        self.active = None
        self.deadline.stop()
        process, directory = job['process'], job['directory']
        if process and process.state() != QProcess.NotRunning:
            process.kill()
            process.waitForFinished(1000)
        pages = []
        if not error:
            try:
                pages = word_preview.read_pages(directory)
            except Exception as failure:
                error = str(failure)
        if error and directory:
            word_preview.stop_owned_word(directory)
        if not error and not self.stopping:
            previous = self.cache.pop(job['digest'], None)
            if previous:
                shutil.rmtree(str(previous[0]), ignore_errors=True)
            self.cache[job['digest']] = (directory, pages)
            while len(self.cache) > 3:
                _, (old, _) = self.cache.popitem(last=False)
                shutil.rmtree(str(old), ignore_errors=True)
        elif directory:
            shutil.rmtree(str(directory), ignore_errors=True)
        if process:
            process.deleteLater()
        view = job['view']()
        if view is not None and not self.stopping:
            view.completed(job['digest'], pages, error)
        QTimer.singleShot(0, self.start_next)

    def timeout(self):
        if self.active:
            self.finish(self.active, 'Word не ответил за 90 секунд. Откройте Word вручную, '
                        'завершите активацию/диалоги и нажмите «Обновить».')

    def shutdown(self):
        self.stopping = True
        self.pending.clear()
        if self.active:
            job = self.active
            if job['directory']:
                (job['directory'] / 'cancel').touch()
            self.finish(job, 'Закрытие приложения')
        for directory, _ in self.cache.values():
            shutil.rmtree(str(directory), ignore_errors=True)
        self.cache.clear()


def renderer():
    app = QApplication.instance()
    if not hasattr(app, '_sz_word_renderer'):
        app._sz_word_renderer = _Renderer(app)
    return app._sz_word_renderer


class WordPreview(QWidget):
    """Отображает изображения страниц, а не повторяет вёрстку Word средствами Qt."""
    def __init__(self):
        super().__init__()
        from .common import button, label  # общий стиль, без циклического импорта
        self.setMinimumWidth(380)
        self.digest = None
        self.data = None
        self.source = None
        self.context = None
        self.images = []
        self.force = False
        self.service = renderer()
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.request)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        bar = QHBoxLayout()
        bar.addWidget(label('Предпросмотр Word', 'panelTitle'))
        bar.addStretch(1)
        self.zoom = QComboBox()
        self.zoom.setAccessibleName('Масштаб предпросмотра')
        for caption, scale in [('По ширине', None), ('100%', 1.0), ('150%', 1.5)]:
            self.zoom.addItem(caption, scale)
        self.zoom.currentIndexChanged.connect(self.scale_pages)
        bar.addWidget(self.zoom)
        self.refresh_btn = button('Обновить', on_click=lambda: self.refresh(True))
        bar.addWidget(self.refresh_btn)
        layout.addLayout(bar)
        self.status = label('Выберите СЗ или заполните форму.', 'muted', wrap=True)
        self.status.setAccessibleName('Состояние предпросмотра')
        layout.addWidget(self.status)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.scroll.setAcceptDrops(False)
        self.scroll.viewport().setAcceptDrops(False)
        self.scroll_viewport = self.scroll.viewport()
        self.scroll_viewport.installEventFilter(self)
        self.paper = QWidget()
        self.paper_layout = QVBoxLayout(self.paper)
        self.paper_layout.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        self.paper_layout.setContentsMargins(12, 12, 12, 12)
        self.paper_layout.setSpacing(16)
        self.paper_layout.setSizeConstraint(QLayout.SetMinAndMaxSize)
        self.scroll.setWidget(self.paper)
        layout.addWidget(self.scroll, 1)

    def viewport(self):
        return self.scroll.viewport()

    def clear_pages(self):
        self.images = []
        while self.paper_layout.count():
            widget = self.paper_layout.takeAt(0).widget()
            if widget:
                widget.deleteLater()
        self.paper_layout.activate()

    def clear(self):
        self.timer.stop()
        self.service.pending.pop(id(self), None)
        self.digest = self.data = self.source = self.context = None
        self.clear_pages()
        self.status.setText('СЗ не выбрана.')

    def show_memo(self, memo, addressees, preview=False, document_path=None):
        if self.context != memo.get('reg_number'):
            self.clear_pages()
        self.context = memo.get('reg_number')
        self.source = (copy.deepcopy(memo), copy.deepcopy(addressees), document_path)
        self.refresh(False)

    def refresh(self, force=False):
        if self.source is None:
            return
        memo, addressees, path = self.source
        try:
            data = Path(path).read_bytes() if path else docx.build_docx(memo, addressees)
            digest = word_preview.document_digest(data)
        except Exception as error:
            self.timer.stop()
            self.digest = self.data = None
            self.clear_pages()
            self.status.setText('Не удалось прочитать DOCX: %s' % error)
            return
        if digest == self.digest and self.images and not force:
            return
        self.data, self.digest, self.force = data, digest, force
        self.status.setText('Показана предыдущая версия; обновление через Word…' if self.images else
                            'Формирование страниц через Microsoft Word…')
        if self.isVisible():
            self.timer.start(0 if force else 1000)

    def request(self):
        if self.data is not None and self.isVisible():
            self.service.request(self, self.digest, self.data, self.force)
            self.force = False

    def completed(self, digest, pages, error):
        if digest != self.digest:
            return  # результат предыдущей редакции не заменяет актуальный
        if error:
            self.clear_pages()
            self.status.setText(error + ' Сохранение СЗ доступно независимо от предпросмотра.')
            return
        if not self.isVisible():
            return
        pixmaps = [QPixmap(str(path)) for path in pages]
        if any(image.isNull() for image in pixmaps):
            self.clear_pages()
            self.status.setText('Не удалось прочитать страницы. Нажмите «Обновить».')
            return
        self.clear_pages()
        for number, pixmap in enumerate(pixmaps, 1):
            image = QLabel()
            image.setAlignment(Qt.AlignCenter)
            image.setAccessibleName('Страница %d из %d' % (number, len(pixmaps)))
            self.paper_layout.addWidget(image, 0, Qt.AlignHCenter)
            self.images.append((image, pixmap))
        self.scale_pages()
        self.status.setText('Страниц: %d. Вёрстка Microsoft Word на этом компьютере.' % len(pages))

    def scale_pages(self):
        scale = self.zoom.currentData()
        for label, pixmap in self.images:
            # PNG от Word имеет 144 dpi; 100% на экране — 96 логических dpi.
            width = (max(180, self.scroll.viewport().width() - 32) if scale is None else
                     max(1, round(pixmap.width() * 2 / 3 * scale)))
            rendered = pixmap.scaledToWidth(width, Qt.SmoothTransformation)
            label.setPixmap(rendered)
            label.setFixedSize(rendered.size())
        self.paper_layout.activate()

    def eventFilter(self, obj, event):
        if obj is self.scroll_viewport and event.type() == QEvent.Resize:
            QTimer.singleShot(0, self.scale_pages)
        return super().eventFilter(obj, event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        QTimer.singleShot(0, self.scale_pages)

    def showEvent(self, event):
        super().showEvent(event)
        if self.source is not None:
            self.refresh(False)
            self.timer.start(0)

    def hideEvent(self, event):
        self.timer.stop()
        super().hideEvent(event)
