"""Страницы: новая СЗ, реестр, просмотр СЗ, статистика, справочники."""
import datetime

from PyQt5.QtCore import QDate, Qt, QTimer
from PyQt5.QtWidgets import (QAbstractItemView, QComboBox, QDateEdit, QFormLayout,
                             QGridLayout, QHBoxLayout, QHeaderView, QLineEdit, QListWidget,
                             QListWidgetItem, QPlainTextEdit, QScrollArea, QSplitter,
                             QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget)

from .. import actions, docx
from ..store import DEFAULT_NUMBER_PREFIX, DEFAULT_OUTPUT_DIR
from .common import (BarList, ColumnChart, SheetView, button, esc, hbox, label, panel,
                     sheet_frame, show_in_folder)
from .editor import FormattedTextEdit

MONTHS = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек']
MONTHS_FULL = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август',
               'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
MONTHS_IN = ['январе', 'феврале', 'марте', 'апреле', 'мае', 'июне', 'июле', 'августе',
             'сентябре', 'октябре', 'ноябре', 'декабре']


class Page(QWidget):
    def __init__(self, win):
        super().__init__()
        self.win = win

    @property
    def store(self):
        return self.win.store

    def reload(self):
        """Перечитывает данные из базы и обновляет страницу."""

    def on_external_change(self):
        """Базу изменил другой пользователь."""
        self.reload()


def page_layout(widget, title=None):
    lay = QVBoxLayout(widget)
    lay.setContentsMargins(24, 18, 24, 18)
    lay.setSpacing(12)
    if title:
        lay.addWidget(label(title, 'title'))
    return lay


def field_label(text, optional=False):
    return label(text + (' <span style="color:#6b7385;font-weight:normal">необязательно</span>'
                         if optional else ''), 'fieldLabel')


def fill_combo(combo, rows, text, placeholder):
    """Заполняет список, сохраняя выбранное значение."""
    current = combo.currentData()
    combo.blockSignals(True)
    combo.clear()
    combo.addItem(placeholder, None)
    for r in rows:
        combo.addItem(text(r), r['id'])
    index = combo.findData(current) if current is not None else 0
    combo.setCurrentIndex(max(0, index))
    combo.blockSignals(False)


def read_only_table(headers):
    t = QTableWidget(0, len(headers))
    t.setHorizontalHeaderLabels(headers)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    t.setSelectionMode(QAbstractItemView.SingleSelection)
    t.verticalHeader().setVisible(False)
    t.setWordWrap(True)
    t.setAlternatingRowColors(False)
    t.setShowGrid(False)
    t.verticalHeader().setDefaultSectionSize(32)
    return t


def cell(text, data=None):
    item = QTableWidgetItem(str(text if text is not None else ''))
    if data is not None:
        item.setData(Qt.UserRole, data)
    return item


def scrolled(widget):
    area = QScrollArea()
    area.setWidgetResizable(True)
    area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    widget.setObjectName('scrollBody')
    area.setWidget(widget)
    return area


# ---------- Новая СЗ ----------

class NewMemoPage(Page):
    def __init__(self, win):
        super().__init__(win)
        self.topics = []
        self.addressees = []
        self.employees = []
        self.signer = None
        self.chosen = []  # id адресатов в порядке выбора
        self.next_numbers = {}  # год -> следующий номер, чтобы не читать базу на каждое нажатие

        form = QWidget()
        lay = page_layout(form, 'Новая служебная записка')
        self.warn = label('', 'warn', wrap=True)
        self.warn.linkActivated.connect(self.win.open_link)
        lay.addWidget(self.warn)

        lay.addWidget(field_label('Тематика работ'))
        self.topic = QComboBox()
        lay.addWidget(self.topic)

        lay.addWidget(field_label('Кому'))
        self.addr_filter = QLineEdit()
        self.addr_filter.setPlaceholderText('Найти отдел или фамилию')
        self.addr_filter.setClearButtonEnabled(True)
        self.addr_filter.textChanged.connect(self.filter_addressees)
        lay.addWidget(self.addr_filter)
        self.addr_list = QListWidget()
        self.addr_list.setMinimumHeight(140)
        self.addr_list.itemChanged.connect(self.addressee_toggled)
        lay.addWidget(self.addr_list)
        lay.addWidget(label('Можно выбрать несколько. В записке адресаты идут в порядке выбора.', 'hint', wrap=True))

        lay.addWidget(field_label('Дата'))
        self.date = QDateEdit(QDate.currentDate())
        self.date.setCalendarPopup(True)
        self.date.setDisplayFormat('dd.MM.yyyy')
        self.date.setMaximumWidth(160)
        lay.addWidget(self.date)

        lay.addWidget(field_label('Текст'))
        self.body = FormattedTextEdit('Каждый абзац с новой строки')
        self.body.setMinimumHeight(220)
        lay.addWidget(self.body)

        lay.addWidget(field_label('Приложение', optional=True))
        self.attachment = QPlainTextEdit()
        self.attachment.setPlaceholderText('Например: смета на 2 л. в 1 экз.')
        self.attachment.setFixedHeight(60)
        lay.addWidget(self.attachment)

        lay.addWidget(field_label('Подписант'))
        self.signer_label = label('', wrap=True)
        lay.addWidget(hbox(self.signer_label,
                           button('Изменить', 'link', lambda: self.win.navigate('dirs', 'employees'))))

        lay.addWidget(field_label('Исполнитель'))
        self.executor = QComboBox()
        lay.addWidget(self.executor)

        self.submit = button('Создать и сформировать Word', 'primary', self.create)
        self.number_note = label('', 'hint', wrap=True)
        lay.addSpacing(6)
        lay.addWidget(hbox(self.submit, self.number_note))
        lay.addStretch(1)

        self.sheet = SheetView()
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(scrolled(form))
        splitter.addWidget(sheet_frame(self.sheet))
        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 6)
        splitter.setSizes([520, 620])
        splitter.setChildrenCollapsible(False)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(splitter)

        for signal in (self.topic.currentIndexChanged, self.executor.currentIndexChanged,
                       self.date.dateChanged, self.body.changed, self.attachment.textChanged):
            signal.connect(self.update_preview)

    def reload(self):
        s = self.store
        self.topics = s.all('SELECT * FROM topics ORDER BY name')
        self.addressees = s.all('SELECT * FROM addressees ORDER BY department, head_name')
        self.employees = s.all('SELECT * FROM employees ORDER BY full_name')
        self.signer = s.signer
        self.next_numbers = {}
        self.prefix, self.org_name, self.city = s.number_prefix, s.org_name, s.city
        fill_combo(self.topic, self.topics, lambda t: t['name'], 'Выберите тематику')
        fill_combo(self.executor, self.employees,
                   lambda e: e['full_name'] + (', ' + e['phone'] if e['phone'] else ''),
                   'Выберите сотрудника')
        ids = {a['id'] for a in self.addressees}
        self.chosen = [i for i in self.chosen if i in ids]
        self.addr_list.blockSignals(True)
        self.addr_list.clear()
        for a in self.addressees:
            item = QListWidgetItem()
            item.setData(Qt.UserRole, a['id'])
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if a['id'] in self.chosen else Qt.Unchecked)
            self.addr_list.addItem(item)
        self.addr_list.blockSignals(False)
        self.addr_filter.setVisible(len(self.addressees) > 6)
        self.update_picker()
        self.filter_addressees()

        missing = []
        if not self.topics:
            missing.append('<a href="dirs:topics">тематики работ</a>')
        if not self.addressees:
            missing.append('<a href="dirs:addressees">адресатов</a>')
        if not self.employees:
            missing.append('<a href="dirs:employees">сотрудников</a>')
        warnings = []
        if missing:
            warnings.append('Чтобы создать записку, добавьте в справочники %s.' % ', '.join(missing))
        if self.employees and not self.signer:
            warnings.append('Подписант не назначен. Отметьте его галочкой в справочнике '
                            '<a href="dirs:employees">«Сотрудники»</a>.')
        self.warn.setText('<br>'.join(warnings))
        self.warn.setVisible(bool(warnings))

        if self.signer:
            self.signer_label.setText('%s<br><span style="color:#6b7385">%s</span>' % (
                esc(docx.short_name(self.signer['full_name'])),
                esc(self.signer['position'] or 'должность не указана').replace('\n', '<br>')))
        else:
            self.signer_label.setText('<span style="color:#6b7385">не назначен</span>')
        self.update_preview()

    def on_external_change(self):
        # Введённый текст не трогаем, обновляем только списки и номер
        self.reload()

    def addressee_text(self, a):
        pos = self.chosen.index(a['id']) + 1 if a['id'] in self.chosen else 0
        return '%s%s\n%s %s' % ('%d. ' % pos if pos else '', a['department'], a['head_position'],
                                docx.addressee_name(a))

    def update_picker(self):
        by_id = {a['id']: a for a in self.addressees}
        self.addr_list.blockSignals(True)
        for i in range(self.addr_list.count()):
            item = self.addr_list.item(i)
            item.setText(self.addressee_text(by_id[item.data(Qt.UserRole)]))
        self.addr_list.blockSignals(False)

    def addressee_toggled(self, item):
        aid = item.data(Qt.UserRole)
        if item.checkState() == Qt.Checked and aid not in self.chosen:
            self.chosen.append(aid)
        elif item.checkState() != Qt.Checked and aid in self.chosen:
            self.chosen.remove(aid)
        self.update_picker()
        self.update_preview()

    def filter_addressees(self):
        q = self.addr_filter.text().strip().lower()
        for i in range(self.addr_list.count()):
            item = self.addr_list.item(i)
            item.setHidden(bool(q) and q not in item.text().lower())

    def form_data(self):
        return {
            'topic_id': self.topic.currentData(),
            'executor_id': self.executor.currentData(),
            'memo_date': self.date.date().toString('yyyy-MM-dd'),
            'body_rich': self.body.paragraphs(),
            'attachment': self.attachment.toPlainText(),
        }

    def update_preview(self):
        if not self.store or not hasattr(self, 'prefix'):
            return
        d = self.form_data()
        year = int(d['memo_date'][:4])
        if year not in self.next_numbers:
            self.next_numbers[year] = actions.next_number(self.store, year)
        reg = docx.format_reg_number(self.prefix, self.next_numbers[year], year)
        topic = next((t for t in self.topics if t['id'] == d['topic_id']), {})
        executor = next((e for e in self.employees if e['id'] == d['executor_id']), {})
        by_id = {a['id']: a for a in self.addressees}
        signer = self.signer or {}
        self.sheet.show_memo({
            'reg_number': reg, 'memo_date': d['memo_date'], 'topic_name': topic.get('name'),
            'body': docx.rich_plain_text(d['body_rich']), 'body_rich': docx.dump_rich(d['body_rich']), 'attachment': d['attachment'], 'org_name': self.org_name,
            'city': self.city, 'signer_name': signer.get('full_name'),
            'signer_position': signer.get('position'), 'executor_name': executor.get('full_name'),
            'executor_phone': executor.get('phone'),
        }, [by_id[i] for i in self.chosen], preview=True)
        self.number_note.setText('Номер %s закрепится при сохранении.' % reg)

    def create(self):
        ok, result = self.win.act(lambda: actions.create_memo(self.store, self.form_data(), list(self.chosen)))
        if not ok:
            return
        memo_id, doc_error = result
        reg = self.store.one('SELECT reg_number FROM memos WHERE id = ?', (memo_id,))['reg_number']
        self.chosen = []
        self.body.reset()
        self.attachment.clear()
        self.date.setDate(QDate.currentDate())
        self.win.navigate('memo', memo_id)
        if doc_error:
            self.win.flash('Записка %s сохранена в реестре, но документ Word создать не удалось: %s'
                           % (reg, doc_error), 'error')
        else:
            self.win.flash('Записка %s создана, документ Word сохранён в папку.' % reg)


# ---------- Реестр ----------

class RegistryPage(Page):
    BODY_PREVIEW_LIMIT = 100  # символов текста до многоточия

    def __init__(self, win):
        super().__init__(win)
        lay = page_layout(self)
        lay.addWidget(hbox(label('Реестр', 'title'), None,
                           button('Новая СЗ', 'primary', lambda: self.win.navigate('new')),
                           stretch_end=False))

        self.q = QLineEdit()
        self.q.setPlaceholderText('Номер, тема, адресат, исполнитель или слово из текста')
        self.q.setClearButtonEnabled(True)
        self.topic = QComboBox()
        self.year = QComboBox()
        self.reset = button('Сбросить', on_click=self.reset_filters)
        bar = QHBoxLayout()
        bar.setSpacing(8)
        bar.addWidget(self.q, 3)
        bar.addWidget(self.topic, 2)
        bar.addWidget(self.year, 1)
        bar.addWidget(self.reset)
        lay.addLayout(bar)

        # Поиск по мере ввода, с небольшой задержкой
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(250)
        self.timer.timeout.connect(self.reload)
        self.q.textChanged.connect(lambda: self.timer.start())
        self.topic.currentIndexChanged.connect(lambda: self.reload())
        self.year.currentIndexChanged.connect(lambda: self.reload())

        self.table = read_only_table(['Номер', 'Дата', 'Тематика', 'Кому', 'Исполнитель', 'Текст СЗ'])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.cellDoubleClicked.connect(lambda row, col: self.open_selected())
        self.table.itemSelectionChanged.connect(self.update_buttons)
        lay.addWidget(self.table, 1)

        self.count = label('', 'muted')
        self.open_btn = button('Открыть', on_click=self.open_selected)
        self.word_btn = button('Открыть Word', on_click=self.open_word)
        lay.addWidget(hbox(self.open_btn, self.word_btn, None, self.count, stretch_end=False))

    def reload(self):
        s = self.store
        fill_combo(self.topic, s.all('SELECT * FROM topics ORDER BY name'), lambda t: t['name'], 'Все тематики')
        years = [r['year'] for r in s.all('SELECT DISTINCT year FROM memos ORDER BY year DESC')]
        current = self.year.currentData()
        self.year.blockSignals(True)
        self.year.clear()
        self.year.addItem('Все годы', None)
        for y in years:
            self.year.addItem(str(y), y)
        self.year.setCurrentIndex(max(0, self.year.findData(current)) if current else 0)
        self.year.blockSignals(False)

        q = self.q.text().strip()
        memos = actions.search_memos(s, q, self.topic.currentData(), self.year.currentData())
        addr = {}
        for a in s.all('SELECT * FROM memo_addressees ORDER BY memo_id, ord'):
            addr.setdefault(a['memo_id'], []).append(a)
        total = s.one('SELECT COUNT(*) AS n FROM memos')['n']

        selected = self.selected_id()
        self.table.setRowCount(len(memos))
        for row, m in enumerate(memos):
            lst = addr.get(m['id'], [])
            to = '\n'.join(a['department'] for a in lst[:2]) + \
                ('\nи ещё %d' % (len(lst) - 2) if len(lst) > 2 else '')
            values = [m['reg_number'], docx.format_date(m['memo_date']), m['topic_name'], to,
                      m['executor_name']]
            body = m.get('body') or ''
            compact = ' '.join(body.split())
            preview = (compact[:self.BODY_PREVIEW_LIMIT].rstrip() + '…'
                       if len(compact) > self.BODY_PREVIEW_LIMIT else compact)
            values.append(preview)
            # Подсказка доступна над любой ячейкой записки; теги из текста
            # показываем буквально, исходные переносы сохраняем.
            tooltip = '<qt>%s</qt>' % esc(body).replace('\n', '<br>') if body else ''
            for col, value in enumerate(values):
                item = cell(value, m['id'])
                item.setToolTip(tooltip)
                self.table.setItem(row, col, item)
            if m['id'] == selected:
                self.table.selectRow(row)
        self.table.resizeRowsToContents()

        filtered = bool(q or self.topic.currentData() or self.year.currentData())
        self.reset.setVisible(filtered)
        if not total:
            self.count.setText('Здесь появятся созданные служебные записки.')
        elif not memos:
            self.count.setText('По этим условиям ничего не нашлось.')
        else:
            self.count.setText('Найдено %d из %d' % (len(memos), total) if filtered else 'Всего %d' % total)
        self.update_buttons()

    def reset_filters(self):
        for w in (self.q, self.topic, self.year):
            w.blockSignals(True)
        self.q.clear()
        self.topic.setCurrentIndex(0)
        self.year.setCurrentIndex(0)
        for w in (self.q, self.topic, self.year):
            w.blockSignals(False)
        self.reload()

    def selected_id(self):
        items = self.table.selectedItems()
        return items[0].data(Qt.UserRole) if items else None

    def update_buttons(self):
        has = self.selected_id() is not None
        self.open_btn.setEnabled(has)
        self.word_btn.setEnabled(has)

    def open_selected(self):
        memo_id = self.selected_id()
        if memo_id is not None:
            self.win.navigate('memo', memo_id)

    def open_word(self):
        memo_id = self.selected_id()
        if memo_id is not None:
            self.win.open_memo_document(memo_id)


# ---------- Просмотр СЗ ----------

class MemoPage(Page):
    def __init__(self, win):
        super().__init__(win)
        self.memo_id = None
        self.memo = None

        side = QWidget()
        lay = page_layout(side)
        lay.addWidget(hbox(button('← Реестр', 'link', lambda: self.win.navigate('list'))))
        self.reg = label('', 'reg')
        self.reg.setTextInteractionFlags(Qt.TextSelectableByMouse)
        lay.addWidget(self.reg)
        box, box_lay = panel()
        self.details = label('', wrap=True)
        self.details.setTextInteractionFlags(Qt.TextSelectableByMouse)
        box_lay.addWidget(self.details)
        lay.addWidget(box)

        self.word_btn = button('Открыть Word', 'primary', self.open_word)
        self.folder_btn = button('Показать в папке', on_click=self.show_folder)
        self.regen_btn = button('Сформировать Word заново', on_click=self.regenerate)
        self.delete_btn = button('Удалить из реестра', 'danger', self.delete)
        for b in (self.word_btn, self.folder_btn, self.regen_btn, self.delete_btn):
            lay.addWidget(b)
        lay.addStretch(1)

        self.sheet = SheetView()
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(scrolled(side))
        splitter.addWidget(sheet_frame(self.sheet))
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 7)
        splitter.setSizes([380, 660])
        splitter.setChildrenCollapsible(False)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(splitter)

    def show_memo(self, memo_id):
        self.memo_id = memo_id
        self.reload()

    def reload(self):
        s = self.store
        self.memo = s.one('SELECT * FROM memos WHERE id = ?', (self.memo_id,))
        m = self.memo
        for b in (self.word_btn, self.folder_btn, self.regen_btn, self.delete_btn):
            b.setEnabled(bool(m))
        if not m:
            self.reg.setText('Записки нет в реестре')
            self.details.setText('Её удалили. Вернитесь в реестр.')
            self.sheet.clear()
            return
        addressees = s.memo_addressees(m['id'])
        path = s.document_path(m['file_path'])
        rows = [
            ('Дата', esc(docx.format_date(m['memo_date']))),
            ('Тематика', esc(m['topic_name'])),
            ('Кому', '<br>'.join(esc(a['department']) + (
                '<br><span style="color:#6b7385">%s %s</span>' % (
                    esc(a['head_position']).replace('\n', '<br>'), esc(docx.addressee_name(a)))
                if (a['head_position'], docx.addressee_name(a)) != (a['department'], '') else '')
                for a in addressees)),
        ]
        if m['signer_name']:
            rows.append(('Подписант', '%s<br><span style="color:#6b7385">%s</span>' % (
                esc(docx.short_name(m['signer_name'])), esc(m['signer_position']).replace('\n', '<br>'))))
        rows.append(('Исполнитель', esc(m['executor_name']) + (
            '<br><span style="color:#6b7385">%s</span>' % esc(m['executor_phone']) if m['executor_phone'] else '')))
        rows.append(('Создана', '%s в %s' % (docx.format_date(m['created_at'][:10]), esc(m['created_at'][11:16]))))
        rows.append(('Документ', '<span style="color:#6b7385">%s</span>' % esc(m['file_path']) if path else
                     '<span style="color:#b42318">Не найден в папке.</span> '
                     'Нажмите «Сформировать Word заново».'))
        self.details.setText('<table cellspacing="0" cellpadding="3">%s</table>' % ''.join(
            '<tr><td style="color:#6b7385; padding-right:12px" valign="top">%s</td><td>%s</td></tr>' % r
            for r in rows))
        self.reg.setText(m['reg_number'])
        self.word_btn.setEnabled(bool(path))
        self.folder_btn.setEnabled(bool(path))
        # Записки прошлых версий не хранят шапку - показываем текущую из настроек
        memo = dict(m)
        if memo['org_name'] is None:
            memo['org_name'] = s.org_name
        self.sheet.show_memo(memo, addressees)

    def open_word(self):
        self.win.open_memo_document(self.memo_id)

    def show_folder(self):
        path = self.store.document_path(self.memo and self.memo['file_path'])
        if path:
            show_in_folder(path)

    def regenerate(self):
        ok, _ = self.win.act(lambda: actions.regenerate_document(self.store, self.memo_id))
        self.reload()
        if ok:
            self.win.flash('Документ Word сформирован заново.')

    def delete(self):
        reg = self.memo['reg_number']
        if not self.win.confirm('Удалить %s из реестра? Документ Word в папке останется.' % reg):
            return
        ok, _ = self.win.act(lambda: actions.delete_memo(self.store, self.memo_id))
        if ok:
            self.win.navigate('list')
            self.win.flash('%s удалена из реестра. Документ Word остался в папке.' % reg)


# ---------- Статистика ----------

class StatsPage(Page):
    def __init__(self, win):
        super().__init__(win)
        self.as_table = False
        content = QWidget()
        lay = page_layout(content)
        self.year = QComboBox()
        self.year.currentIndexChanged.connect(lambda: self.reload())
        lay.addWidget(hbox(label('Статистика', 'title'), None, self.year, stretch_end=False))

        self.empty = label('', 'muted')
        lay.addWidget(self.empty)

        self.body = QWidget()
        body = QVBoxLayout(self.body)
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(12)
        lay.addWidget(self.body)

        head, head_lay = panel(QHBoxLayout)
        hero = QVBoxLayout()
        self.hero_value = label('', 'heroValue')
        self.hero_label = label('', 'muted')
        hero.addWidget(self.hero_value)
        hero.addWidget(self.hero_label)
        head_lay.addLayout(hero)
        head_lay.addStretch(1)
        self.tiles = []
        for _ in range(4):
            tile = QVBoxLayout()
            caption, value = label('', 'muted'), label('', 'tileValue')
            tile.addWidget(caption)
            tile.addWidget(value)
            head_lay.addSpacing(24)
            head_lay.addLayout(tile)
            self.tiles.append((caption, value))
        body.addWidget(head)

        chart_box, chart_lay = panel()
        self.toggle = button('Показать таблицей', on_click=self.toggle_table)
        chart_lay.addWidget(hbox(label('По месяцам', 'panelTitle'), None, self.toggle, stretch_end=False))
        self.chart = ColumnChart()
        chart_lay.addWidget(self.chart)
        self.month_table = read_only_table(['Месяц', 'СЗ'])
        self.month_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.month_table.setMinimumHeight(390)
        chart_lay.addWidget(self.month_table)
        body.addWidget(chart_box)

        grid = QGridLayout()
        grid.setSpacing(12)
        self.bars = {}
        for i, (kind, title) in enumerate([('topic', 'По тематикам'), ('addressee', 'По адресатам'),
                                           ('executor', 'По исполнителям')]):
            box, box_lay = panel()
            box_lay.addWidget(label(title, 'panelTitle'))
            bars = BarList()
            box_lay.addWidget(bars)
            grid.addWidget(box, 0, i, Qt.AlignTop)
            self.bars[kind] = bars
        body.addLayout(grid)
        lay.addStretch(1)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scrolled(content))

    def toggle_table(self):
        self.as_table = not self.as_table
        self.reload()

    def reload(self):
        s = self.store
        now = datetime.date.today()
        years = [r['year'] for r in s.all('SELECT DISTINCT year FROM memos ORDER BY year DESC')]
        if now.year not in years:
            years.insert(0, now.year)
        year = self.year.currentData()
        if year not in years:
            year = now.year
        self.year.blockSignals(True)
        self.year.clear()
        for y in years:
            self.year.addItem(str(y), y)
        self.year.setCurrentIndex(years.index(year))
        self.year.blockSignals(False)

        is_current = year == now.year
        total = s.one('SELECT COUNT(*) AS n FROM memos WHERE year = ?', (year,))['n']
        self.body.setVisible(bool(total))
        self.empty.setVisible(not total)
        self.empty.setText('За %d год служебных записок нет.' % year)
        if not total:
            return
        all_time = s.one('SELECT COUNT(*) AS n FROM memos')['n']
        by_month = actions.month_counts(s, year)

        def distinct(col):
            return s.one('SELECT COUNT(DISTINCT %s) AS n FROM memos WHERE year = ?' % col, (year,))['n']

        if is_current:
            month_tile = ('В %s' % MONTHS_IN[now.month - 1], str(by_month[now.month - 1]))
        else:
            month_tile = ('В среднем за месяц', ('%.1f' % (total / 12)).replace('.', ',').replace(',0', ''))
        self.hero_value.setText(str(total))
        self.hero_label.setText('служебных записок за %d год' % year)
        tiles = [month_tile, ('Тематик', str(distinct('topic_name'))),
                 ('Исполнителей', str(distinct('executor_name'))), ('Всего в реестре', str(all_time))]
        for (caption, value), (c, v) in zip(self.tiles, tiles):
            caption.setText(c)
            value.setText(v)

        self.chart.set_data(by_month, MONTHS, ['%s %d: %d СЗ' % (MONTHS_FULL[i], year, n)
                                               for i, n in enumerate(by_month)],
                            now.month - 1 if is_current else -1)
        self.month_table.setRowCount(12)
        for i, n in enumerate(by_month):
            self.month_table.setItem(i, 0, cell(MONTHS_FULL[i]))
            self.month_table.setItem(i, 1, cell(n))
        self.chart.setVisible(not self.as_table)
        self.month_table.setVisible(self.as_table)
        self.toggle.setText('Показать график' if self.as_table else 'Показать таблицей')
        for kind, bars in self.bars.items():
            bars.set_data(actions.group_counts(s, kind, year), total)


# ---------- Справочники ----------

class DirectoryTab(QWidget):
    def __init__(self, page, kind):
        super().__init__()
        self.page = page
        self.win = page.win
        self.kind = kind
        self.spec = actions.DIRECTORIES[kind]
        self.edit_id = None
        self.rows = []

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 12, 0, 0)
        lay.setSpacing(12)
        box, box_lay = panel()
        self.form_title = label('Добавить запись', 'panelTitle')
        box_lay.addWidget(self.form_title)
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignTop)
        self.inputs = {}
        for name, title, required, multiline in self.spec['fields']:
            if multiline:
                w = QPlainTextEdit()
                w.setFixedHeight(54)
                w.setPlaceholderText('Enter - перенос строки')
            else:
                w = QLineEdit()
                w.returnPressed.connect(self.save)
            self.inputs[name] = w
            form.addRow(field_label(title, optional=not required), w)
        box_lay.addLayout(form)
        self.save_btn = button('Добавить', 'primary', self.save)
        self.cancel_btn = button('Отмена', on_click=self.cancel)
        box_lay.addWidget(hbox(self.save_btn, self.cancel_btn))
        lay.addWidget(box)

        headers = [f[1] for f in self.spec['fields']] + (['Подписант'] if self.spec.get('signer') else [])
        self.table = read_only_table(headers)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        if self.spec.get('signer'):
            header.setSectionResizeMode(len(headers) - 1, QHeaderView.ResizeToContents)
        self.table.cellDoubleClicked.connect(lambda row, col: self.edit_selected())
        self.table.itemSelectionChanged.connect(self.update_buttons)
        self.table.itemChanged.connect(self.item_changed)
        lay.addWidget(self.table, 1)
        self.edit_btn = button('Изменить', on_click=self.edit_selected)
        self.delete_btn = button('Удалить', 'danger', self.delete_selected)
        self.count = label('', 'muted')
        lay.addWidget(hbox(self.edit_btn, self.delete_btn, None, self.count, stretch_end=False))
        self.set_editing(None)

    @property
    def store(self):
        return self.win.store

    def reload(self):
        self.rows = actions.directory_rows(self.store, self.kind)
        if self.edit_id is not None and not any(r['id'] == self.edit_id for r in self.rows):
            self.set_editing(None)
        fields = self.spec['fields']
        signer = next((r for r in self.rows if r.get('is_signer')), None)
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.rows))
        for row, r in enumerate(self.rows):
            for col, f in enumerate(fields):
                self.table.setItem(row, col, cell(r[f[0]] or '—', r['id']))
            if self.spec.get('signer'):
                self.table.setItem(row, len(fields), self.signer_item(r, signer))
            if any('\n' in str(r[f[0]] or '') for f in fields):
                self.table.resizeRowToContents(row)
        self.table.blockSignals(False)
        self.count.setText('Записей: %d' % len(self.rows) if self.rows else
                           'Справочник «%s» пока пуст. Добавьте первую запись в форме выше.' % self.spec['title'])
        self.update_buttons()

    def signer_item(self, row, current):
        """Подписант один: пока галочка стоит у одного сотрудника, у остальных она заблокирована."""
        item = cell('', row['id'])
        item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
        item.setCheckState(Qt.Checked if row['is_signer'] else Qt.Unchecked)
        if current and current['id'] != row['id']:
            item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsSelectable)
            item.setToolTip('Подписант уже назначен: %s. Сначала снимите галочку у него.'
                            % docx.short_name(current['full_name']))
        return item

    def item_changed(self, item):
        if self.spec.get('signer') and item.column() == len(self.spec['fields']):
            self.set_signer(item.data(Qt.UserRole), item.checkState() == Qt.Checked)

    def set_signer(self, employee_id, on):
        ok, _ = self.win.act(lambda: actions.set_signer(self.store, employee_id, on))
        if ok:
            self.win.flash('Подписант назначен. Он будет подставляться в новые записки.' if on
                           else 'Подписант снят.')
        # Перерисовываем после выхода из обработчика галочки
        QTimer.singleShot(0, self.reload)

    def selected(self):
        items = self.table.selectedItems()
        if not items:
            return None
        rid = items[0].data(Qt.UserRole)
        return next((r for r in self.rows if r['id'] == rid), None)

    def update_buttons(self):
        has = self.selected() is not None
        self.edit_btn.setEnabled(has)
        self.delete_btn.setEnabled(has)

    def set_editing(self, row):
        self.edit_id = row['id'] if row else None
        for name, w in self.inputs.items():
            value = row[name] if row else ''
            if isinstance(w, QPlainTextEdit):
                w.setPlainText(value or '')
            else:
                w.setText(value or '')
        self.form_title.setText('Изменить запись' if row else 'Добавить запись')
        self.save_btn.setText('Сохранить' if row else 'Добавить')
        self.cancel_btn.setVisible(bool(row))
        if row:
            next(iter(self.inputs.values())).setFocus()

    def edit_selected(self):
        row = self.selected()
        if row:
            self.set_editing(row)

    def cancel(self):
        self.set_editing(None)

    def values(self):
        return {name: (w.toPlainText() if isinstance(w, QPlainTextEdit) else w.text())
                for name, w in self.inputs.items()}

    def save(self):
        editing = self.edit_id is not None
        ok, _ = self.win.act(lambda: actions.save_directory_row(self.store, self.kind, self.values(), self.edit_id))
        if ok:
            self.win.flash('Изменения сохранены.' if editing else 'Запись добавлена.')
            self.set_editing(None)
            self.reload()

    def delete_selected(self):
        row = self.selected()
        if not row or not self.win.confirm('Удалить запись? В уже созданных записках данные сохранятся.'):
            return
        ok, _ = self.win.act(lambda: actions.delete_directory_row(self.store, self.kind, row['id']))
        if ok:
            self.win.flash('Запись удалена.')
            self.reload()


class SettingsTab(QWidget):
    def __init__(self, page):
        super().__init__()
        self.win = page.win
        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(0, 12, 0, 0)
        lay.setSpacing(12)

        box, f = panel()
        f.addWidget(field_label('Наименование организации'))
        self.org_name = QPlainTextEdit()
        self.org_name.setPlaceholderText('Каждая строка шапки с новой строки')
        self.org_name.setFixedHeight(80)
        f.addWidget(self.org_name)
        f.addWidget(label('Печатается в шапке по центру полужирным. Переносы строк сохраняются как введены.', 'hint'))
        f.addWidget(field_label('Город'))
        self.city = QLineEdit()
        self.city.setPlaceholderText('Например, Москва')
        self.city.setMaximumWidth(320)
        f.addWidget(self.city)
        f.addWidget(label('Печатается под заголовком «СЛУЖЕБНАЯ ЗАПИСКА».', 'hint'))
        f.addWidget(field_label('Префикс номера'))
        self.prefix = QLineEdit()
        self.prefix.setMaxLength(20)
        self.prefix.setMaximumWidth(320)
        self.prefix.textChanged.connect(self.update_example)
        f.addWidget(self.prefix)
        self.example = label('', 'hint')
        f.addWidget(self.example)
        f.addWidget(field_label('Папка для документов Word'))
        self.output_dir = QLineEdit()
        f.addWidget(self.output_dir)
        f.addWidget(label('Внутри общей папки, например «%s» или «Документы/СЗ». '
                          'Подпапки по годам создаются сами.' % DEFAULT_OUTPUT_DIR, 'hint', wrap=True))
        f.addSpacing(4)
        f.addWidget(hbox(button('Сохранить настройки', 'primary', self.save)))
        lay.addWidget(box)

        box, f = panel()
        f.addWidget(label('Общая папка на этом компьютере', 'panelTitle'))
        self.folder = label('', wrap=True)
        self.folder.setTextInteractionFlags(Qt.TextSelectableByMouse)
        f.addWidget(self.folder)
        f.addWidget(label('Здесь лежат база (data/memos.db) и документы Word. '
                          'Настройка запоминается только для этого пользователя Windows.', 'hint', wrap=True))
        f.addWidget(hbox(button('Выбрать другую папку…', on_click=self.win.choose_folder)))
        lay.addWidget(box)
        lay.addStretch(1)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scrolled(content))

    def reload(self):
        s = self.win.store
        self.org_name.setPlainText(s.org_name)
        self.city.setText(s.city)
        self.prefix.setText(s.number_prefix)
        self.output_dir.setText(s.output_dir)
        self.folder.setText(s.root)

    def on_external_change(self):
        pass  # не затираем то, что пользователь сейчас вводит

    def update_example(self):
        prefix = self.prefix.text().strip() or DEFAULT_NUMBER_PREFIX
        self.example.setText('Номер выглядит так: %s. Уже созданные записки сохраняют свои номера.'
                             % docx.format_reg_number(prefix, 1, datetime.date.today().year))

    def save(self):
        ok, _ = self.win.act(lambda: actions.save_settings(
            self.win.store, self.org_name.toPlainText(), self.city.text(), self.prefix.text(),
            self.output_dir.text()))
        if ok:
            self.win.flash('Настройки сохранены.')
            self.reload()
            self.win.update_side()


class DirsPage(Page):
    KINDS = ['employees', 'addressees', 'topics', 'settings']

    def __init__(self, win):
        super().__init__(win)
        lay = page_layout(self, 'Справочники')
        self.tabs = QTabWidget()
        self.tab_widgets = {}
        for kind in self.KINDS:
            w = SettingsTab(self) if kind == 'settings' else DirectoryTab(self, kind)
            self.tab_widgets[kind] = w
            self.tabs.addTab(w, 'Настройки' if kind == 'settings' else actions.DIRECTORIES[kind]['title'])
        self.tabs.currentChanged.connect(lambda i: self.current().reload())
        lay.addWidget(self.tabs, 1)

    def current(self):
        return self.tabs.currentWidget()

    def show_tab(self, kind):
        if kind in self.tab_widgets:
            self.tabs.blockSignals(True)
            self.tabs.setCurrentWidget(self.tab_widgets[kind])
            self.tabs.blockSignals(False)

    def reload(self):
        self.current().reload()

    def on_external_change(self):
        tab = self.current()
        if isinstance(tab, SettingsTab):
            tab.on_external_change()
        else:
            tab.reload()

