"""Служебные записки. Запуск: SZ.bat или environment\\python\\pythonw.exe main.pyw"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# pythonw.exe запускается без консоли: ошибки пишем в журнал во временной папке пользователя
if sys.stderr is None:
    try:
        sys.stderr = sys.stdout = open(os.path.join(tempfile.gettempdir(), 'sz-errors.log'), 'a',
                                       encoding='utf-8', buffering=1)
    except OSError:
        pass

from sz.app import main  # noqa: E402

if __name__ == '__main__':
    sys.exit(main())
