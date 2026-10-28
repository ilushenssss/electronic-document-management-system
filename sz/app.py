"""Запуск программы."""
import configparser
import os
import sys
import traceback

import PyQt5
from PyQt5.QtCore import QCoreApplication, QLibraryInfo, QLocale, QSettings, Qt, QTranslator
from PyQt5.QtWidgets import QApplication, QMessageBox

from .ui.common import STYLE
from .ui.window import MainWindow

# Папка программы: в ней лежат SZ.bat, main.pyw, sz/ и environment/
PROGRAM_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INI_FILE = os.path.join(PROGRAM_DIR, 'sz.ini')


def resolve_folder():
    """Общая папка с базой и документами.

    1. Папка, выбранная пользователем в программе (запоминается для пользователя Windows).
    2. Параметр folder в sz.ini рядом с программой - чтобы задать папку сразу для всех.
    3. Папка самой программы.
    """
    chosen = QSettings().value('folder', '')
    if chosen:
        return chosen
    folder = ini_folder()
    if folder:
        return os.path.join(PROGRAM_DIR, folder)  # относительный путь - от папки программы
    return PROGRAM_DIR


def ini_folder():
    if not os.path.isfile(INI_FILE):
        return ''
    # Блокнот сохраняет в UTF-8, а в старых Windows - в кодировке ANSI
    for encoding in ('utf-8-sig', 'cp1251'):
        ini = configparser.ConfigParser()
        try:
            ini.read(INI_FILE, encoding=encoding)
            return ini.get('main', 'folder', fallback='').strip()
        except UnicodeDecodeError:
            continue
        except configparser.Error:
            return ''
    return ''


def qt_path(name, kind):
    """Папка модулей или переводов Qt внутри PyQt5.

    PyQt5 сообщает Qt путь к ним сам, но портит в нём не латинские буквы: из папки
    вроде «Общие\\Служебные записки» Qt не находит модуль окна и программа
    не запускается. Поэтому путь берётся напрямую из расположения PyQt5.
    """
    path = os.path.join(os.path.dirname(os.path.abspath(PyQt5.__file__)), 'Qt5', name)
    return path if os.path.isdir(path) else QLibraryInfo.location(kind)


def excepthook(kind, value, tb):
    text = ''.join(traceback.format_exception(kind, value, tb))
    if sys.stderr:
        sys.stderr.write(text)
    if QApplication.instance():
        QMessageBox.critical(None, 'Служебные записки', 'Непредвиденная ошибка:\n\n%s' % text)


def main():
    QCoreApplication.setLibraryPaths([qt_path('plugins', QLibraryInfo.PluginsPath)])
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setOrganizationName('SZ')
    app.setApplicationName('Служебные записки')
    sys.excepthook = excepthook

    # Русские надписи в стандартных окнах Qt (кнопки «Да»/«Нет», календарь)
    QLocale.setDefault(QLocale(QLocale.Russian, QLocale.Russia))
    translator = QTranslator(app)
    if translator.load('qtbase_ru', qt_path('translations', QLibraryInfo.TranslationsPath)):
        app.installTranslator(translator)

    app.setStyle('Fusion')
    app.setStyleSheet(STYLE)
    win = MainWindow(resolve_folder)
    win.show()
    win.connect()
    return app.exec_()
