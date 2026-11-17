"""Справочник шаблонов и ввод полей перед вставкой в СЗ."""
import json

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFormLayout,
                             QHeaderView, QLineEdit, QSplitter, QVBoxLayout, QWidget)

from .. import templates
from .common import button, hbox, label, panel
from .editor import FormattedTextEdit


class TemplateFieldsDialog(QDialog):
    def __init__(self, names, defaults, parent=None):
        super().__init__(parent)
        from .pages import scrolled
        self.setWindowTitle('Заполнить поля шаблона')
        self.resize(500, min(560, 180 + len(names) * 45))
        lay = QVBoxLayout(self)
        lay.addWidget(label('Значения будут вставлены в текст СЗ. Все поля обязательны.', 'hint', wrap=True))
        form_widget = QWidget()
        form = QFormLayout(form_widget)
        self.inputs = {}
        for key, name in names.items():
            edit = QLineEdit(defaults.get(key, ''))
            edit.setAccessibleName(name)
            form.addRow(name + ':', edit)
            self.inputs[key] = edit
        lay.addWidget(scrolled(form_widget), 1)
        self.error = label('', 'warn', wrap=True)
        self.error.hide()
        lay.addWidget(self.error)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText('Вставить')
        buttons.button(QDialogButtonBox.Cancel).setText('Отмена')
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        lay.addWidget(buttons)
        if self.inputs:
            next(iter(self.inputs.values())).setFocus()

    def values(self):
        return {key: edit.text().strip() for key, edit in self.inputs.items()}

    def accept(self):
        for edit in self.inputs.values():
            if not edit.text().strip():
                self.error.setText('Заполните все поля. Для отмены вставки нажмите «Отмена».')
                self.error.show()
                edit.setFocus()
                return
        super().accept()


class TemplatesTab(QWidget):
    def __init__(self, page):
        super().__init__()
        from .pages import read_only_table, scrolled
        self.page, self.win = page, page.win
        self.edit_id, self.version = None, None
        self.rows = []
        self.baseline = None
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 12, 0, 0)
        box, form_lay = panel()
        self.title = label('Добавить шаблон', 'panelTitle')
        form_lay.addWidget(self.title)
        form = QFormLayout()
        self.name = QLineEdit()
        self.name.setAccessibleName('Название шаблона')
        self.kind = QComboBox()
        for key, name in templates.KINDS.items():
            self.kind.addItem(name, key)
        self.kind.setAccessibleName('Вид шаблона')
        self.topic = QComboBox()
        self.topic.setAccessibleName('Тематика шаблона')
        self.active = QCheckBox('Доступен для вставки')
        form.addRow('Название:', self.name)
        form.addRow('Вид:', self.kind)
        form.addRow('Тематика:', self.topic)
        form.addRow('', self.active)
        form_lay.addLayout(form)
        form_lay.addWidget(label('Поля в фигурных скобках заполняются при вставке: '
                                '{наименование}, {количество}, {цель}. '
                                '{дата}, {исполнитель}, {тематика} берутся из формы СЗ; их можно изменить.',
                                'hint', wrap=True))
        self.body = FormattedTextEdit('Введите текст шаблона или отдельную формулировку')
        self.body.setMinimumHeight(200)
        self.body.edit.setAccessibleName('Текст шаблона')
        form_lay.addWidget(self.body)
        self.save_btn = button('Добавить', 'primary', self.save)
        self.cancel_btn = button('Отмена', on_click=self.cancel)
        top = QWidget()
        top_lay = QVBoxLayout(top)
        top_lay.setContentsMargins(0, 0, 0, 0)
        top_lay.addWidget(scrolled(box), 1)
        top_lay.addWidget(hbox(self.save_btn, self.cancel_btn))
        self.table = read_only_table(['Название', 'Вид', 'Тематика', 'Доступен'])
        self.table.setAccessibleName('Шаблоны СЗ')
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.cellDoubleClicked.connect(lambda row, col: self.edit_selected())
        self.table.itemSelectionChanged.connect(self.update_buttons)
        bottom = QWidget()
        bottom_lay = QVBoxLayout(bottom)
        bottom_lay.setContentsMargins(0, 0, 0, 0)
        bottom_lay.addWidget(self.table, 1)
        self.edit_btn = button('Изменить', on_click=self.edit_selected)
        self.delete_btn = button('Удалить', 'danger', self.delete_selected)
        self.examples_btn = button('Добавить примеры', on_click=self.add_examples)
        bottom_lay.addWidget(hbox(self.edit_btn, self.delete_btn, None, self.examples_btn, stretch_end=False))
        self.count = label('', 'hint', wrap=True)
        bottom_lay.addWidget(self.count)
        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(top)
        splitter.addWidget(bottom)
        splitter.setChildrenCollapsible(False)
        splitter.setSizes([430, 230])
        lay.addWidget(splitter)
        self.reset_form()

    @property
    def store(self):
        return self.win.store

    def data(self):
        return dict(name=self.name.text(), kind=self.kind.currentData(),
                    topic_id=self.topic.currentData(), active=self.active.isChecked(),
                    paragraphs=self.body.paragraphs())

    def state(self):
        return json.dumps(self.data(), ensure_ascii=False, sort_keys=True)

    def has_unsaved_changes(self):
        return self.baseline is not None and self.state() != self.baseline

    def may_discard(self):
        return not self.has_unsaved_changes() or self.win.confirm(
            'Отменить несохранённые изменения шаблона? Введённый текст будет потерян.')

    def reset_form(self):
        self.edit_id, self.version = None, None
        self.name.clear()
        self.kind.setCurrentIndex(0)
        self.topic.setCurrentIndex(0)
        self.active.setChecked(True)
        self.body.reset()
        self.title.setText('Добавить шаблон')
        self.save_btn.setText('Добавить')
        self.baseline = self.state()

    def cancel(self):
        if self.may_discard():
            self.reset_form()

    def reload(self):
        from .pages import cell, fill_combo
        was_clean = not self.has_unsaved_changes()
        # Сохраняем выбранную удалённую тематику, чтобы обновление извне не меняло черновик.
        topic_id, topic_name = self.topic.currentData(), self.topic.currentText()
        topics = self.store.all('SELECT * FROM topics ORDER BY name')
        if topic_id is not None and not any(row['id'] == topic_id for row in topics):
            topics.append({'id': topic_id, 'name': topic_name})
        fill_combo(self.topic, topics, lambda row: row['name'], 'Любая тематика')
        if was_clean:
            self.baseline = self.state()
        self.rows = templates.list_templates(self.store)
        selected = self.selected()
        selected_id = selected['id'] if selected else None
        self.table.setRowCount(len(self.rows))
        for i, row in enumerate(self.rows):
            values = [row['name'], templates.KINDS[row['kind']], row['topic_name'] or 'Любая',
                      'Да' if row['active'] else 'Нет']
            for j, value in enumerate(values):
                item = cell(value, row['id'])
                item.setToolTip(row['body'])
                self.table.setItem(i, j, item)
            if row['id'] == selected_id:
                self.table.selectRow(i)
        self.count.setText('Шаблонов: %d' % len(self.rows) if self.rows else
                           'Шаблонов пока нет. Добавьте свой шаблон или загрузите примеры.')
        self.update_buttons()

    def selected(self):
        item = self.table.item(self.table.currentRow(), 0)
        row_id = item.data(Qt.UserRole) if item else None
        return next((row for row in self.rows if row['id'] == row_id), None)

    def update_buttons(self):
        selected = self.selected() is not None
        self.edit_btn.setEnabled(selected)
        self.delete_btn.setEnabled(selected)

    def edit_selected(self):
        row = self.selected()
        if not row or not self.may_discard():
            return
        ok, fresh = self.win.act(lambda: templates.get_template(self.store, row['id']))
        if not ok:
            return
        self.reload()
        self.edit_id, self.version = fresh['id'], fresh['version']
        self.name.setText(fresh['name'])
        self.kind.setCurrentIndex(self.kind.findData(fresh['kind']))
        self.topic.setCurrentIndex(max(0, self.topic.findData(fresh['topic_id'])))
        self.active.setChecked(bool(fresh['active']))
        self.body.set_paragraphs(templates.paragraphs(fresh))
        self.title.setText('Изменить шаблон')
        self.save_btn.setText('Сохранить')
        self.baseline = self.state()
        self.name.setFocus()

    def save(self):
        if not self.name.text().strip():
            self.name.setFocus()
        elif not self.body.plain_text().strip():
            self.body.edit.setFocus()
        ok, _ = self.win.act(lambda: templates.save_template(
            self.store, self.data(), self.edit_id, self.version))
        if ok:
            self.reset_form()
            self.reload()
            self.win.flash('Шаблон сохранён.')

    def delete_selected(self):
        row = self.selected()
        if not row or not self.may_discard():
            return
        if not self.win.confirm('Удалить шаблон «%s»? Тексты созданных СЗ сохранятся.' % row['name']):
            return
        ok, _ = self.win.act(lambda: templates.delete_template(self.store, row['id'], row['version']))
        if ok:
            self.reset_form()
            self.reload()
            self.win.flash('Шаблон удалён.')

    def add_examples(self):
        ok, count = self.win.act(lambda: templates.add_examples(self.store))
        if ok:
            self.reload()
            self.win.flash('Добавлено примеров: %d.' % count)
