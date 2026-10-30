"""Главное окно: боковое меню, сообщения, переключение страниц, выбор общей папки."""
import os
import sqlite3
import traceback

from PyQt5.QtCore import QEvent, QSettings, Qt, QTimer
from PyQt5.QtWidgets import (QApplication, QButtonGroup, QFileDialog, QHBoxLayout, QMainWindow,
                             QMessageBox, QPushButton, QStackedWidget, QVBoxLayout, QWidget)

from ..store import Store, StoreError
from . import pages
from .common import FLASH_STYLES, button, label, open_file

REFRESH_MS = 4000  # как часто проверять изменения других пользователей
FLASH_MS = 8000

NAV = [('new', 'Новая СЗ'), ('list', 'Реестр'), ('stats', 'Статистика'), ('dirs', 'Справочники')]


class ConnectPage(QWidget):
    """Показывается, когда общая папка недоступна."""

    def __init__(self, win):
        super().__init__()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(48, 48, 48, 48)
        lay.setSpacing(12)
        lay.addWidget(label('Служебные записки', 'title'))
        self.message = label('', 'flashError', wrap=True)
        lay.addWidget(self.message)
        lay.addWidget(label('Реестр и документы Word хранятся в общей папке:', 'muted'))
        self.path = label('', wrap=True)
        self.path.setTextInteractionFlags(Qt.TextSelectableByMouse)
        lay.addWidget(self.path)
        row = QHBoxLayout()
        row.addWidget(button('Повторить', 'primary', win.connect))
        row.addWidget(button('Выбрать другую папку…', on_click=win.choose_folder))
        row.addStretch(1)
        lay.addLayout(row)
        lay.addStretch(1)

    def show_error(self, folder, message):
        self.path.setText(folder)
        self.message.setText(message)


class MainWindow(QMainWindow):
    def __init__(self, folder_resolver):
        super().__init__()
        self.resolve_folder = folder_resolver
        self.store = None
        self.setWindowTitle('Служебные записки')
        self.resize(1280, 820)
        self.setMinimumSize(980, 620)

        root = QWidget()
        root_lay = QHBoxLayout(root)
        root_lay.setContentsMargins(0, 0, 0, 0)
        root_lay.setSpacing(0)
        self.setCentralWidget(root)

        # Боковое меню
        self.side = QWidget()
        self.side.setObjectName('side')
        self.side.setAttribute(Qt.WA_StyledBackground, True)
        self.side.setFixedWidth(230)
        side = QVBoxLayout(self.side)
        side.setContentsMargins(14, 18, 14, 14)
        side.setSpacing(4)
        side.addWidget(label('Служебные записки', 'appName'))
        self.side_org = label('', wrap=True)
        side.addWidget(self.side_org)
        side.addSpacing(14)
        self.nav = QButtonGroup(self)
        self.nav_buttons = {}
        for key, title in NAV:
            b = QPushButton(title)
            b.setObjectName('nav')
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda checked=False, k=key: self.navigate(k))
            self.nav.addButton(b)
            self.nav_buttons[key] = b
            side.addWidget(b)
        side.addStretch(1)
        self.side_folder = label('', 'sideFolder', wrap=True)
        self.side_folder.setToolTip('Общая папка')
        side.addWidget(self.side_folder)
        change = QPushButton('Сменить папку')
        change.setObjectName('sideLink')
        change.setCursor(Qt.PointingHandCursor)
        change.clicked.connect(lambda: self.choose_folder())
        side.addWidget(change)
        root_lay.addWidget(self.side)

        # Содержимое: сообщение и страницы
        content = QWidget()
        content.setObjectName('content')
        content_lay = QVBoxLayout(content)
        content_lay.setContentsMargins(0, 0, 0, 0)
        content_lay.setSpacing(0)
        self.flash_label = label('', wrap=True)
        self.flash_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.flash_label.hide()
        flash_wrap = QWidget()
        flash_lay = QVBoxLayout(flash_wrap)
        flash_lay.setContentsMargins(24, 14, 24, 0)
        flash_lay.addWidget(self.flash_label)
        content_lay.addWidget(flash_wrap)
        self.stack = QStackedWidget()
        content_lay.addWidget(self.stack, 1)
        root_lay.addWidget(content, 1)

        self.connect_page = ConnectPage(self)
        self.stack.addWidget(self.connect_page)
        self.pages = {
            'new': pages.NewMemoPage(self),
            'edit': pages.EditMemoPage(self),
            'list': pages.RegistryPage(self),
            'memo': pages.MemoPage(self),
            'stats': pages.StatsPage(self),
            'dirs': pages.DirsPage(self),
        }
        for page in self.pages.values():
            self.stack.addWidget(page)

        self.flash_timer = QTimer(self)
        self.flash_timer.setSingleShot(True)
        self.flash_timer.timeout.connect(self.flash_label.hide)
        self.refresh_timer = QTimer(self)
        self.refresh_timer.setInterval(REFRESH_MS)
        self.refresh_timer.timeout.connect(self.check_changes)

    # ---------- Подключение к папке ----------

    def connect(self):
        folder = self.resolve_folder()
        if self.store:
            self.store.close()
            self.store = None
        self.refresh_timer.stop()
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            store = Store(folder)
            store.open()
        except (StoreError, OSError, sqlite3.Error) as e:
            QApplication.restoreOverrideCursor()
            self.side.setVisible(False)
            self.flash_label.hide()
            self.connect_page.show_error(folder, 'Не удалось открыть базу: %s' % e)
            self.stack.setCurrentWidget(self.connect_page)
            return False
        QApplication.restoreOverrideCursor()
        self.store = store
        self.side.setVisible(True)
        self.update_side()
        self.navigate('list')
        self.refresh_timer.start()
        return True

    def choose_folder(self):
        start = self.store.root if self.store else self.resolve_folder()
        folder = QFileDialog.getExistingDirectory(self, 'Общая папка служебных записок', start)
        if not folder:
            return
        folder = os.path.normpath(folder)
        has_data = os.path.exists(os.path.join(folder, 'data', 'memos.db'))
        if not has_data and not self.confirm(
                'В папке %s ещё нет базы служебных записок. Создать новую базу в этой папке?' % folder):
            return
        current = self.stack.currentWidget()
        if isinstance(current, pages.EditMemoPage) and current.has_unsaved_changes():
            if not self.confirm('Сменить общую папку? Несохранённые изменения формы будут потеряны. '
                                'Если Word не обновлён, его потребуется обновить позже.'):
                return
            current.edit_snapshot = None
        QSettings().setValue('folder', folder)
        self.connect()

    def update_side(self):
        if not self.store:
            return
        self.side_org.setText(' '.join(self.store.org_name.split('\n')))
        self.side_org.setVisible(bool(self.store.org_name))
        self.side_folder.setText(self.store.root)

    # ---------- Страницы ----------

    def navigate(self, page, arg=None):
        if not self.store:
            return
        current = self.stack.currentWidget()
        if isinstance(current, pages.EditMemoPage) and current.has_unsaved_changes():
            if not self.confirm('Уйти из редактирования? Несохранённые изменения формы будут потеряны. '
                                'Если Word не обновлён, его потребуется обновить позже.'):
                return
        self.flash_label.hide()
        widget = self.pages[page]
        active = 'list' if page in ('memo', 'edit') else page
        for key, b in self.nav_buttons.items():
            b.setChecked(key == active)
        try:
            if page in ('memo', 'edit'):
                widget.show_memo(arg)
            elif page == 'dirs':
                widget.show_tab(arg or 'employees')
                widget.reload()
            else:
                widget.reload()
        except Exception as e:
            self.report(e)
            return
        self.stack.setCurrentWidget(widget)

    def open_link(self, href):
        """Ссылки в тексте страниц: dirs:<справочник>."""
        page, _, arg = href.partition(':')
        self.navigate(page, arg or None)

    def current_page(self):
        w = self.stack.currentWidget()
        return w if isinstance(w, pages.Page) else None

    def check_changes(self):
        """Подтягивает изменения других пользователей."""
        if not self.store:
            return
        try:
            if self.store.changed():
                self.update_side()
                page = self.current_page()
                if page:
                    page.on_external_change()
        except Exception:
            traceback.print_exc()

    def changeEvent(self, event):
        if event.type() == QEvent.ActivationChange and self.isActiveWindow():
            self.check_changes()
        super().changeEvent(event)

    # ---------- Сообщения и действия ----------

    def flash(self, message, kind='success'):
        self.flash_label.setObjectName('flashError' if kind == 'error' else 'flashSuccess')
        self.flash_label.setStyleSheet(FLASH_STYLES[kind == 'error'])
        self.flash_label.setText(message)
        self.flash_label.show()
        if kind == 'error':
            self.flash_timer.stop()
        else:
            self.flash_timer.start(FLASH_MS)

    def report(self, e):
        if isinstance(e, StoreError):
            self.flash(str(e), 'error')
        else:
            traceback.print_exc()
            self.flash('Ошибка: %s' % e, 'error')

    def act(self, fn):
        """Выполняет сохранение с часами вместо курсора. Возвращает (успех, результат)."""
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            result = fn()
        except Exception as e:
            QApplication.restoreOverrideCursor()
            self.report(e)
            return False, None
        QApplication.restoreOverrideCursor()
        return True, result

    def confirm(self, text):
        return QMessageBox.question(self, 'Служебные записки', text,
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.Yes

    def open_memo_document(self, memo_id):
        memo = self.store.one('SELECT * FROM memos WHERE id = ?', (memo_id,))
        path = memo and self.store.document_path(memo['file_path'])
        if not path:
            self.flash('Документ Word не найден в папке. Откройте записку и сформируйте его заново.', 'error')
            return
        open_file(path)

    def closeEvent(self, event):
        current = self.stack.currentWidget()
        if isinstance(current, pages.EditMemoPage) and current.has_unsaved_changes():
            if not self.confirm('Закрыть приложение? Несохранённые изменения формы будут потеряны. '
                                'Если Word не обновлён, его потребуется обновить позже.'):
                event.ignore()
                return
        if self.store:
            self.store.close()
        super().closeEvent(event)

