"""Общие части интерфейса: лист записки, графики, оформление."""
import html
import math
import os
import subprocess
import sys

from PyQt5.QtCore import QRectF, QSize, QUrl, Qt
from PyQt5.QtGui import QColor, QDesktopServices, QFont, QPainter, QPainterPath, QPen
from PyQt5.QtWidgets import (QFrame, QHBoxLayout, QLabel, QSizePolicy, QTextBrowser, QToolTip,
                             QVBoxLayout, QWidget)

from .. import docx

ACCENT = '#2f5bd3'
MUTED = '#6b7385'
PLACEHOLDER = '#a0a7b5'

STYLE = """
QMainWindow, QWidget#content { background: #f4f5f8; }
QWidget { font-size: 10pt; color: #1d2433; }
QWidget#side { background: #1f2a44; }
QWidget#side QLabel { color: #c9d1e3; }
QLabel#appName { color: #ffffff; font-size: 12pt; font-weight: 600; }
QLabel#sideFolder { color: #8d98b3; font-size: 8pt; }
QPushButton#nav { color: #c9d1e3; background: transparent; border: none; text-align: left;
    padding: 9px 14px; border-radius: 6px; font-size: 10.5pt; }
QPushButton#nav:hover { background: #2b3858; color: #ffffff; }
QPushButton#nav:checked { background: #3a4a72; color: #ffffff; font-weight: 600; }
QPushButton#sideLink { color: #8d98b3; background: transparent; border: none; text-align: left;
    padding: 2px 0; font-size: 8pt; text-decoration: underline; }
QPushButton#sideLink:hover { color: #ffffff; }
QLabel#title { font-size: 16pt; font-weight: 600; }
QLabel#reg { font-size: 16pt; font-weight: 600; color: #1d2433; }
QLabel#muted { color: #6b7385; }
QLabel#fieldLabel { color: #3c4457; font-weight: 600; }
QLabel#hint { color: #6b7385; font-size: 9pt; }
QLabel#warn { background: #fff4e5; border: 1px solid #f3c98b; border-radius: 6px; padding: 8px 10px; }
QLabel#flashError { background: #fdecec; border: 1px solid #f0a9a9; color: #8a1f1f;
    border-radius: 6px; padding: 9px 12px; }
QFrame#panel { background: #ffffff; border: 1px solid #dde1e9; border-radius: 8px; }
QFrame#panel QLabel { background: transparent; }
QLabel#heroValue { font-size: 30pt; font-weight: 600; color: #2f5bd3; }
QLabel#tileValue { font-size: 16pt; font-weight: 600; }
QLabel#panelTitle { font-size: 11pt; font-weight: 600; }
QPushButton { background: #ffffff; border: 1px solid #c8cedb; border-radius: 6px; padding: 6px 14px; }
QPushButton:hover { border-color: #2f5bd3; }
QPushButton:disabled { color: #a0a7b5; border-color: #e1e4ea; }
QPushButton#primary { background: #2f5bd3; border-color: #2f5bd3; color: #ffffff; font-weight: 600; }
QPushButton#primary:hover { background: #2649ad; }
QPushButton#primary:disabled { background: #9db0e6; border-color: #9db0e6; }
QPushButton#danger { color: #b42318; }
QPushButton#danger:hover { border-color: #b42318; }
QPushButton#link { border: none; background: transparent; color: #2f5bd3; padding: 0;
    text-decoration: underline; }
QLineEdit, QPlainTextEdit, QComboBox, QDateEdit, QListWidget {
    background: #ffffff; border: 1px solid #c8cedb; border-radius: 6px; padding: 4px 6px; }
QLineEdit:focus, QPlainTextEdit:focus, QComboBox:focus, QDateEdit:focus, QListWidget:focus {
    border-color: #2f5bd3; }
QTableWidget { background: #ffffff; border: 1px solid #dde1e9; border-radius: 8px;
    gridline-color: #eceef3; selection-background-color: #dfe7fb; selection-color: #1d2433; }
QHeaderView::section { background: #f7f8fa; border: none; border-bottom: 1px solid #dde1e9;
    padding: 6px; font-weight: 600; color: #3c4457; }
QTabWidget::pane { border: none; }
QTabBar::tab { background: transparent; padding: 8px 16px; border: none;
    border-bottom: 2px solid transparent; color: #3c4457; }
QTabBar::tab:selected { border-bottom: 2px solid #2f5bd3; color: #1d2433; font-weight: 600; }
QScrollArea { border: none; background: transparent; }
QWidget#scrollBody { background: #f4f5f8; }
QFrame#formatBar { background: #ffffff; border: 1px solid #c8cedb; border-radius: 6px; }
QFrame#formatBar QToolButton { padding: 3px; border-radius: 4px; color: #3c4457; }
QFrame#formatBar QToolButton:hover { background: #eef1f7; }
QFrame#formatBar QToolButton:checked { background: #dfe7fb; color: #1d2433; }
QFrame#formatBar QComboBox { padding: 1px 4px; min-width: 44px; }
QTextEdit#richEdit { background: #ffffff; border: 1px solid #c8cedb; border-radius: 6px; padding: 4px 6px; }
QTextEdit#richEdit:focus { border-color: #2f5bd3; }
QFrame#sheetWrap { background: #e4e7ee; border-radius: 8px; }
QTextBrowser#sheet { background: #ffffff; border: 1px solid #d3d8e2; }
"""

# Сообщение над страницей: (успех, ошибка)
FLASH_STYLES = (
    'background: #e7f6ec; border: 1px solid #9ad3ad; color: #1d5b33; border-radius: 6px; padding: 9px 12px;',
    'background: #fdecec; border: 1px solid #f0a9a9; color: #8a1f1f; border-radius: 6px; padding: 9px 12px;',
)


def esc(value):
    return html.escape(str(value if value is not None else ''))


def label(text='', name=None, wrap=False):
    w = QLabel(text)
    if name:
        w.setObjectName(name)
    w.setWordWrap(wrap)
    return w


def button(text, name=None, on_click=None):
    from PyQt5.QtWidgets import QPushButton
    b = QPushButton(text)
    if name:
        b.setObjectName(name)
    b.setCursor(Qt.PointingHandCursor)
    if on_click:
        b.clicked.connect(lambda checked=False: on_click())
    return b


def panel(layout_cls=QVBoxLayout):
    frame = QFrame()
    frame.setObjectName('panel')
    lay = layout_cls(frame)
    lay.setContentsMargins(16, 14, 16, 14)
    lay.setSpacing(8)
    return frame, lay


def hbox(*widgets, stretch_end=True, spacing=8):
    w = QWidget()
    lay = QHBoxLayout(w)
    lay.setContentsMargins(0, 0, 0, 0)
    lay.setSpacing(spacing)
    for x in widgets:
        if x is None:
            lay.addStretch(1)
        else:
            lay.addWidget(x)
    if stretch_end:
        lay.addStretch(1)
    return w


def open_file(path):
    QDesktopServices.openUrl(QUrl.fromLocalFile(path))


def show_in_folder(path):
    if sys.platform == 'win32':
        subprocess.Popen('explorer /select,"%s"' % os.path.normpath(path))
    else:
        open_file(os.path.dirname(path))


# ---------- Лист служебной записки (как будет в Word) ----------

SHEET_SCALE = 0.8  # лист на экране чуть мельче документа


class SheetView(QTextBrowser):
    """Лист строится из тех же блоков, что и документ Word (docx.memo_blocks)."""

    def __init__(self):
        super().__init__()
        self.setObjectName('sheet')
        self.setOpenLinks(False)
        self.setMinimumWidth(380)
        self.document().setDocumentMargin(34)
        font = QFont(docx.FONT)
        font.setPointSizeF(docx.FONT_SIZE * SHEET_SCALE)
        self.document().setDefaultFont(font)

    def show_memo(self, memo, addressees, preview=False):
        self.setHtml(sheet_html(docx.memo_blocks(memo, addressees, preview=preview)))


def sheet_html(blocks):
    def t(x):
        if isinstance(x, dict):
            return '<span style="color:%s">%s</span>' % (PLACEHOLDER, esc(x['ph']))
        return esc(x)

    parts = []
    for b in blocks:
        size = (b.get('size') or docx.FONT_SIZE) * SHEET_SCALE
        style = 'margin:0; font-size:%.1fpt;' % size
        if b.get('blank'):
            parts.append('<p style="%s">&nbsp;</p>' % style)
        elif b.get('split'):
            left, right = b['split']
            parts.append('<table width="100%%" border="0" cellspacing="0" cellpadding="0" style="%s">'
                         '<tr><td>%s</td><td align="right">%s</td></tr></table>' % (style, t(left), t(right)))
        else:
            align = {'center': 'center', 'right': 'right', 'both': 'justify'}.get(b.get('align'), 'left')
            if b.get('first_line'):
                style += ' text-indent:%dpx;' % round(1.25 * 37.8 * SHEET_SCALE)
            runs = b.get('runs') or [{'text': b.get('text'), 'bold': b.get('bold')}]
            text = ''.join(run_html(r, t) for r in runs)
            parts.append('<p align="%s" style="%s">%s</p>' % (align, style, text))
    return ''.join(parts)


def run_html(run, t=esc):
    text = t(run['text']).replace('\n', '<br>')
    for flag, tag in (('bold', 'b'), ('italic', 'i'), ('underline', 'u')):
        if run.get(flag):
            text = '<%s>%s</%s>' % (tag, text, tag)
    if run.get('size'):
        text = '<span style="font-size:%.1fpt">%s</span>' % (run['size'] * SHEET_SCALE, text)
    return text


def sheet_frame(view):
    wrap = QFrame()
    wrap.setObjectName('sheetWrap')
    lay = QVBoxLayout(wrap)
    lay.setContentsMargins(18, 18, 18, 18)
    lay.addWidget(view)
    return wrap


# ---------- Графики ----------

def nice_step(maximum):
    raw = max(1, maximum) / 4
    pow10 = 10 ** math.floor(math.log10(raw))
    step = next(k * pow10 for k in (1, 2, 5, 10) if k * pow10 >= raw)
    return max(1, int(step))


class ColumnChart(QWidget):
    """Столбцы по месяцам с подписями значений и подсказкой при наведении."""

    def __init__(self):
        super().__init__()
        self.values = []
        self.labels = []
        self.tips = []
        self.highlight = -1
        self.setMouseTracking(True)
        self.setMinimumHeight(220)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

    def sizeHint(self):
        return QSize(720, 240)

    def set_data(self, values, labels, tips, highlight=-1):
        self.values, self.labels, self.tips, self.highlight = values, labels, tips, highlight
        self.update()

    def _geometry(self):
        left, right, top, bottom = 34, 8, 22, 28
        plot_w = self.width() - left - right
        plot_h = self.height() - top - bottom
        return left, top, plot_w, plot_h

    def paintEvent(self, event):
        if not self.values:
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        left, top, plot_w, plot_h = self._geometry()
        step = nice_step(max(self.values))
        peak = max(step, -(-max(self.values) // step) * step)

        def y(v):
            return top + plot_h - v / peak * plot_h

        small = QFont(self.font())
        small.setPointSizeF(8.5)
        p.setFont(small)
        for v in range(0, peak + 1, step):
            p.setPen(QPen(QColor('#e6e9ef'), 1))
            p.drawLine(left, int(y(v)), left + plot_w, int(y(v)))
            p.setPen(QColor(MUTED))
            p.drawText(QRectF(0, y(v) - 8, left - 8, 16), Qt.AlignRight | Qt.AlignVCenter, str(v))
        band = plot_w / len(self.values)
        col_w = min(26, band * 0.56)
        for i, v in enumerate(self.values):
            cx = left + band * i + band / 2
            if v > 0:
                path = QPainterPath()
                rect = QRectF(cx - col_w / 2, y(v), col_w, y(0) - y(v))
                path.addRoundedRect(rect, 3, 3)
                p.fillPath(path, QColor(ACCENT))
                p.fillRect(QRectF(rect.left(), rect.bottom() - min(3, rect.height()), col_w,
                                  min(3, rect.height())), QColor(ACCENT))
                p.setPen(QColor('#1d2433'))
                p.drawText(QRectF(cx - 30, y(v) - 20, 60, 18), Qt.AlignCenter, str(v))
            bold = QFont(small)
            bold.setBold(i == self.highlight)
            p.setFont(bold)
            p.setPen(QColor('#1d2433' if i == self.highlight else MUTED))
            p.drawText(QRectF(cx - band / 2, top + plot_h + 4, band, 20), Qt.AlignCenter, self.labels[i])
            p.setFont(small)

    def mouseMoveEvent(self, event):
        left, top, plot_w, plot_h = self._geometry()
        if self.values and left <= event.x() < left + plot_w:
            i = int((event.x() - left) / (plot_w / len(self.values)))
            QToolTip.showText(event.globalPos(), self.tips[i], self)
        else:
            QToolTip.hideText()


class BarList(QWidget):
    """Горизонтальные полосы: подпись, полоса, значение."""
    ROW = 30

    def __init__(self):
        super().__init__()
        self.rows = []
        self.total = 0
        self.setMouseTracking(True)

    def set_data(self, rows, total):
        self.rows, self.total = rows, total
        self.setFixedHeight(max(1, len(rows)) * self.ROW)
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        if not self.rows:
            p.setPen(QColor(MUTED))
            p.drawText(self.rect(), Qt.AlignLeft | Qt.AlignVCenter, 'Нет данных.')
            return
        label_w = int(self.width() * 0.42)
        track_w = self.width() - label_w - 44
        peak = max(r['n'] for r in self.rows)
        metrics = p.fontMetrics()
        for i, r in enumerate(self.rows):
            top = i * self.ROW
            p.setPen(QColor('#1d2433'))
            text = metrics.elidedText(str(r['label']), Qt.ElideRight, label_w - 10)
            p.drawText(QRectF(0, top, label_w - 10, self.ROW), Qt.AlignLeft | Qt.AlignVCenter, text)
            w = max(2, track_w * r['n'] / peak)
            path = QPainterPath()
            path.addRoundedRect(QRectF(label_w, top + 8, w, self.ROW - 16), 3, 3)
            p.fillPath(path, QColor(ACCENT))
            p.drawText(QRectF(label_w + w + 6, top, 40, self.ROW), Qt.AlignLeft | Qt.AlignVCenter, str(r['n']))

    def mouseMoveEvent(self, event):
        i = event.y() // self.ROW
        if 0 <= i < len(self.rows) and self.total:
            r = self.rows[i]
            QToolTip.showText(event.globalPos(), '%s: %d (%d%% за год)'
                              % (r['label'], r['n'], round(r['n'] / self.total * 100)), self)
        else:
            QToolTip.hideText()
