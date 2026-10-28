"""Поле «Текст» с панелью оформления: жирный, курсив, подчёркивание, размер шрифта,
выравнивание, отступ первой строки и перенос строки внутри абзаца.

Оформление переводится в список абзацев (docx.normalize_rich) - из него строятся и лист
предпросмотра, и документ Word.
"""
from PyQt5.QtCore import QPointF, QRectF, Qt, pyqtSignal
from PyQt5.QtGui import (QColor, QFont, QIcon, QKeySequence, QPainter, QPainterPath, QPen, QPixmap,
                         QTextBlockFormat, QTextCharFormat, QTextDocument)
from PyQt5.QtWidgets import (QAction, QActionGroup, QComboBox, QFrame, QHBoxLayout, QTextEdit, QToolButton,
                             QVBoxLayout, QWidget)

from .. import docx
from .common import SHEET_SCALE

LINE_SEPARATOR = ' '  # так Qt хранит перенос строки внутри абзаца (Shift+Enter)
OBJECT_CHAR = '￼'  # вставленные картинки - в записку не попадают
INDENT_PX = 1.25 * 37.8 * SHEET_SCALE  # отступ первой строки 1,25 см, как на листе
ALIGNS = [('left', Qt.AlignLeft, 'По левому краю', 'Ctrl+L'),
          ('center', Qt.AlignHCenter, 'По центру', 'Ctrl+E'),
          ('right', Qt.AlignRight, 'По правому краю', 'Ctrl+R'),
          ('both', Qt.AlignJustify, 'По ширине', 'Ctrl+J')]


def align_name(flags):
    h = int(flags) & int(Qt.AlignHorizontal_Mask)
    if h & Qt.AlignJustify:
        return 'both'
    if h & Qt.AlignHCenter:
        return 'center'
    if h & Qt.AlignRight:
        return 'right'
    return 'left'


def screen_size(size):
    """Размер шрифта в поле: тот же масштаб, что у листа предпросмотра."""
    return (size or docx.FONT_SIZE) * SHEET_SCALE


def char_format(run, size=None):
    fmt = QTextCharFormat()
    fmt.setFontWeight(QFont.Bold if run.get('bold') else QFont.Normal)
    fmt.setFontItalic(bool(run.get('italic')))
    fmt.setFontUnderline(bool(run.get('underline')))
    fmt.setFontPointSize(screen_size(run.get('size') or size))
    return fmt


def block_format(align='both', indent=True):
    fmt = QTextBlockFormat()
    fmt.setAlignment(dict((a[0], a[1]) for a in ALIGNS)[align])
    fmt.setTextIndent(INDENT_PX if indent else 0)
    return fmt


def document_paragraphs(doc, keep_layout=True):
    """Абзацы документа Qt в виде списка для docx. Без keep_layout берутся только
    жирный, курсив и подчёркивание (для текста, вставленного из других программ)."""
    paragraphs = []
    block = doc.begin()
    while block.isValid():
        runs = []
        it = block.begin()
        while not it.atEnd():
            frag = it.fragment()
            if frag.isValid():
                cf = frag.charFormat()
                run = {'text': frag.text().replace(LINE_SEPARATOR, '\n').replace(OBJECT_CHAR, '')}
                if cf.fontWeight() >= QFont.DemiBold:
                    run['bold'] = True
                if cf.fontItalic():
                    run['italic'] = True
                if cf.fontUnderline():
                    run['underline'] = True
                if keep_layout and cf.fontPointSize():
                    run['size'] = int(cf.fontPointSize() / SHEET_SCALE + 0.5)
                runs.append(run)
            it += 1
        fmt = block.blockFormat()
        paragraphs.append({'align': align_name(fmt.alignment()), 'indent': fmt.textIndent() > 0,
                           'runs': runs})
        block = block.next()
    return paragraphs


class RichEdit(QTextEdit):
    def insertFromMimeData(self, source):
        # Из Word, браузера и почты берём только текст, жирный, курсив и подчёркивание:
        # чужие шрифты, цвета и размеры сделали бы записки разномастными
        cursor = self.textCursor()
        size = cursor.charFormat().fontPointSize() / SHEET_SCALE or None
        if source.hasHtml():
            doc = QTextDocument()
            doc.setHtml(source.html())
            # Скопированное из этого же поля вставляется с размером шрифта
            own = source.hasFormat('application/x-qrichtext')
            paragraphs = document_paragraphs(doc, keep_layout=own)
        elif source.hasText():
            paragraphs = [{'runs': [{'text': line}]} for line in source.text().splitlines()]
        else:
            return
        cursor.beginEditBlock()
        for i, p in enumerate(paragraphs):
            if i:
                cursor.insertBlock()  # новый абзац с оформлением текущего
            for r in p['runs']:
                if r['text']:
                    cursor.insertText(r['text'].replace('\n', LINE_SEPARATOR), char_format(r, size))
        cursor.endEditBlock()
        self.ensureCursorVisible()


def _icon(draw):
    pix = QPixmap(36, 36)
    pix.setDevicePixelRatio(2)
    pix.fill(Qt.transparent)
    p = QPainter(pix)
    p.setRenderHint(QPainter.Antialiasing)
    p.setPen(QPen(QColor('#3c4457'), 1.6, Qt.SolidLine, Qt.RoundCap))
    draw(p)
    p.end()
    return QIcon(pix)


def align_icon(align):
    def draw(p):
        for i, width in enumerate((12, 8, 12, 6)):
            y = 4 + i * 3.4
            if align == 'both':
                width = 12
            x = {'left': 3, 'center': 3 + (12 - width) / 2, 'right': 15 - width, 'both': 3}[align]
            p.drawLine(QPointF(x, y), QPointF(x + width, y))
    return _icon(draw)


def indent_icon():
    def draw(p):
        p.drawLine(QPointF(8, 4), QPointF(15, 4))
        for y in (7.4, 10.8, 14.2):
            p.drawLine(QPointF(3, y), QPointF(15, y))
        p.drawLine(QPointF(3, 4), QPointF(5.5, 4))
    return _icon(draw)


def break_icon():
    def draw(p):
        path = QPainterPath(QPointF(14, 3.5))
        path.lineTo(14, 10)
        path.lineTo(4.5, 10)
        p.drawPath(path)
        p.drawLine(QPointF(4.5, 10), QPointF(7.5, 7))
        p.drawLine(QPointF(4.5, 10), QPointF(7.5, 13))
    return _icon(draw)


def clear_icon():
    def draw(p):
        font = QFont(p.font())
        font.setPixelSize(11)
        font.setBold(True)
        p.setFont(font)
        p.drawText(QRectF(1, 1, 12, 13), Qt.AlignCenter, 'A')
        p.setPen(QPen(QColor('#b42318'), 1.6, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(QPointF(10.5, 10.5), QPointF(16, 16))
        p.drawLine(QPointF(16, 10.5), QPointF(10.5, 16))
    return _icon(draw)


class FormattedTextEdit(QWidget):
    """Поле ввода с панелью оформления над ним."""
    changed = pyqtSignal()

    def __init__(self, placeholder=''):
        super().__init__()
        self.edit = RichEdit()
        self.edit.setObjectName('richEdit')
        self.edit.setPlaceholderText(placeholder)
        self.edit.setTabChangesFocus(True)
        font = QFont(docx.FONT)
        font.setPointSizeF(screen_size(None))
        self.edit.document().setDefaultFont(font)
        self.edit.textChanged.connect(self.changed)
        self.edit.currentCharFormatChanged.connect(lambda fmt: self.sync())
        self.edit.cursorPositionChanged.connect(self.sync)

        bar = QFrame()
        bar.setObjectName('formatBar')
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(4, 3, 4, 3)
        lay.setSpacing(2)

        def add(action):
            self.edit.addAction(action)  # сочетания клавиш работают, пока курсор в поле
            action.setShortcutContext(Qt.WidgetShortcut)
            b = QToolButton()
            b.setDefaultAction(action)
            b.setAutoRaise(True)
            b.setFocusPolicy(Qt.NoFocus)
            lay.addWidget(b)
            return b

        def action(text, tip, shortcut=None, checkable=False, icon=None, on=None):
            a = QAction(text, self)
            if icon:
                a.setIcon(icon)
            a.setCheckable(checkable)
            if shortcut:
                a.setShortcut(QKeySequence(shortcut))
                tip = '%s (%s)' % (tip, shortcut)
            a.setToolTip(tip)
            if on:
                a.triggered.connect(on)
            return a

        self.bold = action('Ж', 'Полужирный', 'Ctrl+B', True, on=self.set_bold)
        self.italic = action('К', 'Курсив', 'Ctrl+I', True, on=self.set_italic)
        self.underline = action('Ч', 'Подчёркнутый', 'Ctrl+U', True, on=self.set_underline)
        for a, style in ((self.bold, 'font-weight: bold;'), (self.italic, 'font-style: italic;'),
                         (self.underline, 'text-decoration: underline;')):
            add(a).setStyleSheet('QToolButton { %s font-family: "Times New Roman", serif; font-size: 11pt; '
                                 'min-width: 22px; }' % style)

        self.size = QComboBox()
        self.size.setToolTip('Размер шрифта')
        self.size.setFocusPolicy(Qt.NoFocus)
        for s in docx.RICH_SIZES:
            self.size.addItem(str(s), s)
        self.size.activated.connect(self.set_size)
        lay.addSpacing(6)
        lay.addWidget(self.size)
        lay.addSpacing(6)

        self.align_group = QActionGroup(self)
        self.align_actions = {}
        for name, _, title, shortcut in ALIGNS:
            a = action(title, title, shortcut, True, align_icon(name),
                       on=lambda checked=False, n=name: self.set_align(n))
            self.align_group.addAction(a)
            self.align_actions[name] = a
            add(a)
        lay.addSpacing(6)
        self.indent = action('Отступ', 'Отступ первой строки', checkable=True, icon=indent_icon(),
                             on=self.set_indent)
        add(self.indent)
        # Shift+Enter поле обрабатывает само, кнопка - для тех, кто не знает сочетания
        add(action('Перенос', 'Перенос строки без нового абзаца (Shift+Enter)', icon=break_icon(),
                   on=self.insert_break))
        lay.addStretch(1)
        add(action('Сбросить', 'Убрать оформление: обычный шрифт 14, по ширине, с отступом',
                   icon=clear_icon(),
                   on=self.clear_format))

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(4)
        outer.addWidget(bar)
        outer.addWidget(self.edit)
        self.reset()

    # ---------- Содержимое ----------

    def paragraphs(self):
        return docx.normalize_rich(document_paragraphs(self.edit.document()))

    def plain_text(self):
        return docx.rich_plain_text(self.paragraphs())

    def reset(self):
        """Пустое поле: по ширине, с отступом первой строки, шрифт 14."""
        self.set_paragraphs([])

    def set_paragraphs(self, paragraphs):
        self.edit.blockSignals(True)
        self.edit.clear()
        cursor = self.edit.textCursor()
        cursor.setBlockFormat(block_format())
        cursor.setCharFormat(char_format({}))
        for i, p in enumerate(paragraphs):
            if i:
                cursor.insertBlock()
            cursor.setBlockFormat(block_format(p.get('align', 'both'), p.get('indent', True)))
            for r in p['runs']:
                cursor.insertText(r['text'].replace('\n', LINE_SEPARATOR), char_format(r))
        self.edit.setTextCursor(cursor)
        self.edit.setCurrentCharFormat(char_format({}))
        self.edit.blockSignals(False)
        self.sync()
        self.changed.emit()

    def set_text(self, text):
        """Обычный текст: каждая строка - абзац."""
        self.set_paragraphs([{'runs': [{'text': line}]} for line in text.splitlines()])

    # ---------- Кнопки ----------

    def merge(self, fmt):
        self.edit.mergeCurrentCharFormat(fmt)
        self.edit.setFocus()

    def set_bold(self, on):
        fmt = QTextCharFormat()
        fmt.setFontWeight(QFont.Bold if on else QFont.Normal)
        self.merge(fmt)

    def set_italic(self, on):
        fmt = QTextCharFormat()
        fmt.setFontItalic(on)
        self.merge(fmt)

    def set_underline(self, on):
        fmt = QTextCharFormat()
        fmt.setFontUnderline(on)
        self.merge(fmt)

    def set_size(self, index):
        fmt = QTextCharFormat()
        fmt.setFontPointSize(screen_size(self.size.itemData(index)))
        self.merge(fmt)

    def set_align(self, name):
        fmt = QTextBlockFormat()
        fmt.setAlignment(dict((a[0], a[1]) for a in ALIGNS)[name])
        self.edit.textCursor().mergeBlockFormat(fmt)  # все абзацы выделения
        self.edit.setFocus()
        self.changed.emit()

    def set_indent(self, on):
        fmt = QTextBlockFormat()
        fmt.setTextIndent(INDENT_PX if on else 0)
        self.edit.textCursor().mergeBlockFormat(fmt)
        self.edit.setFocus()
        self.changed.emit()

    def insert_break(self):
        self.edit.textCursor().insertText(LINE_SEPARATOR)
        self.edit.setFocus()

    def clear_format(self):
        cursor = self.edit.textCursor()
        cursor.beginEditBlock()
        if cursor.hasSelection():
            cursor.setCharFormat(char_format({}))
        cursor.mergeBlockFormat(block_format())
        cursor.endEditBlock()
        self.edit.setCurrentCharFormat(char_format({}))
        self.edit.setFocus()
        self.sync()
        self.changed.emit()

    def sync(self):
        """Кнопки показывают оформление текста под курсором."""
        fmt = self.edit.currentCharFormat()
        self.bold.setChecked(fmt.fontWeight() >= QFont.DemiBold)
        self.italic.setChecked(fmt.fontItalic())
        self.underline.setChecked(fmt.fontUnderline())
        size = int((fmt.fontPointSize() or screen_size(None)) / SHEET_SCALE + 0.5)
        index = self.size.findData(size)
        self.size.setCurrentIndex(index if index >= 0 else self.size.findData(docx.FONT_SIZE))
        block = self.edit.textCursor().blockFormat()
        self.align_actions[align_name(block.alignment())].setChecked(True)
        self.indent.setChecked(block.textIndent() > 0)
