"""Сборка папки environment/python: Python 3.8 для Windows x64 и PyQt5.

Нужен только для обновления окружения на компьютере с интернетом; на рабочих
компьютерах ничего скачивать не нужно - готовое окружение лежит в репозитории.

Запуск (из любой ОС, где есть Python 3 и pip):
    python environment/build_environment.py

Что делает:
1. Скачивает официальную сборку Python с python.org из NuGet (пакет «python»)
   и колёса PyQt5 для Windows x64 с PyPI, проверяет их SHA-256.
2. Распаковывает Python в environment/python и ставит в него PyQt5.
3. Убирает модули Qt, которые программе не нужны (QML, мультимедиа, сеть, ...).
4. Если установлен pefile (pip install pefile), проверяет, что всем DLL
   хватает библиотек внутри окружения и системных библиотек Windows.
"""
import hashlib
import io
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
TARGET = os.path.join(HERE, 'python')

PYTHON_VERSION = '3.8.10'
PYTHON_URL = 'https://api.nuget.org/v3-flatcontainer/python/{0}/python.{0}.nupkg'.format(PYTHON_VERSION)
PYTHON_SHA256 = 'c63a2fc8fc62612b5abd391fda99ae1d90bad42ed8a9e99b1bd81c7ffdb4fefa'

WHEELS = {
    'PyQt5-5.15.11-cp38-abi3-win_amd64.whl':
        'bdde598a3bb95022131a5c9ea62e0a96bd6fb28932cc1619fd7ba211531b7517',
    'PyQt5_Qt5-5.15.2-py3-none-win_amd64.whl':
        '750b78e4dba6bdf1607febedc08738e318ea09e9b10aea9ff0d73073f11f6962',
    'PyQt5_sip-12.15.0-cp38-cp38-win_amd64.whl':
        'f600ae6f03e4bff91153c0dc7ebe52f90bd2b6afda58fd580e6990b3b951adc0',
}

# Что оставить от PyQt5 (остальное удаляется)
KEEP_PYQT = {'__init__.py', 'QtCore.pyd', 'QtGui.pyd', 'QtWidgets.pyd', 'sip.cp38-win_amd64.pyd', 'Qt5'}
KEEP_QT_BIN = {
    'Qt5Core.dll', 'Qt5Gui.dll', 'Qt5Widgets.dll',
    # OpenGL через DirectX (ANGLE) - Qt подгружает их сам, если понадобится
    'libEGL.dll', 'libGLESv2.dll', 'd3dcompiler_47.dll',
    # Библиотеки Visual C++, чтобы не требовать их установки в Windows
    'concrt140.dll', 'msvcp140.dll', 'msvcp140_1.dll', 'msvcp140_2.dll',
    'vcruntime140.dll', 'vcruntime140_1.dll',
}
KEEP_QT_PLUGINS = {
    'platforms': {'qwindows.dll', 'qminimal.dll', 'qoffscreen.dll'},
    'styles': {'qwindowsvistastyle.dll'},
    'imageformats': {'qgif.dll', 'qico.dll', 'qjpeg.dll'},
}
KEEP_QT_TRANSLATIONS = {'qtbase_ru.qm'}

# Пути поиска модулей: только окружение, без переменных PYTHONPATH/PYTHONHOME
PTH = 'DLLs\nLib\n.\nimport site\n'


def sha256(data):
    return hashlib.sha256(data).hexdigest()


def download(url):
    print('Скачивание', url)
    with urllib.request.urlopen(url) as r:
        return r.read()


def install_python():
    data = download(PYTHON_URL)
    if sha256(data) != PYTHON_SHA256:
        sys.exit('Неверная контрольная сумма пакета Python')
    if os.path.exists(TARGET):
        shutil.rmtree(TARGET)
    prefix = 'tools/'
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for info in z.infolist():
            if not info.filename.startswith(prefix) or info.is_dir():
                continue
            rel = info.filename[len(prefix):]
            # Заголовки и библиотеки для сборки расширений и демонстрации не нужны
            if rel.split('/')[0] in ('include', 'libs', 'Tools'):
                continue
            path = os.path.join(TARGET, *rel.split('/'))
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with z.open(info) as src, open(path, 'wb') as dst:
                shutil.copyfileobj(src, dst)
    with open(os.path.join(TARGET, 'python38._pth'), 'w', newline='\r\n') as f:
        f.write(PTH)


def install_pyqt():
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.check_call([
            sys.executable, '-m', 'pip', 'download', '--only-binary=:all:', '--no-deps',
            '--platform', 'win_amd64', '--python-version', '3.8', '--implementation', 'cp',
            '--abi', 'cp38', '--abi', 'abi3', '--abi', 'none', '-d', tmp,
            'PyQt5==5.15.11', 'PyQt5-Qt5==5.15.2', 'PyQt5-sip==12.15.0'])
        for name, digest in WHEELS.items():
            with open(os.path.join(tmp, name), 'rb') as f:
                if sha256(f.read()) != digest:
                    sys.exit('Неверная контрольная сумма ' + name)
        subprocess.check_call([
            sys.executable, '-m', 'pip', 'install', '--no-compile', '--no-index', '--no-deps',
            '--find-links', tmp, '--target', os.path.join(TARGET, 'Lib', 'site-packages'),
            '--platform', 'win_amd64', '--python-version', '3.8', '--implementation', 'cp',
            '--only-binary=:all:'] + [os.path.join(tmp, n) for n in WHEELS])


def remove(path):
    if os.path.isdir(path):
        shutil.rmtree(path)
    elif os.path.exists(path):
        os.remove(path)


def keep_only(directory, names):
    for name in os.listdir(directory):
        if name not in names:
            remove(os.path.join(directory, name))


def trim_pyqt():
    site = os.path.join(TARGET, 'Lib', 'site-packages')
    remove(os.path.join(site, 'bin'))  # pyuic5, pyrcc5 и т. п. - для разработчиков
    pyqt = os.path.join(site, 'PyQt5')
    keep_only(pyqt, KEEP_PYQT)
    qt = os.path.join(pyqt, 'Qt5')
    keep_only(qt, {'bin', 'plugins', 'translations'})
    keep_only(os.path.join(qt, 'bin'), KEEP_QT_BIN)
    keep_only(os.path.join(qt, 'plugins'), set(KEEP_QT_PLUGINS))
    for sub, names in KEEP_QT_PLUGINS.items():
        keep_only(os.path.join(qt, 'plugins', sub), names)
    keep_only(os.path.join(qt, 'translations'), KEEP_QT_TRANSLATIONS)
    for root, dirs, _ in os.walk(TARGET):
        for d in dirs:
            if d == '__pycache__':
                shutil.rmtree(os.path.join(root, d))


# Библиотеки, которые есть в любой Windows 7 и новее
SYSTEM_DLLS = {
    'advapi32.dll', 'comctl32.dll', 'comdlg32.dll', 'crypt32.dll', 'd3d11.dll', 'd3d9.dll', 'dwmapi.dll',
    'dxgi.dll', 'gdi32.dll', 'imm32.dll', 'kernel32.dll', 'mpr.dll', 'msi.dll', 'msvcrt.dll', 'netapi32.dll',
    'ntdll.dll', 'ole32.dll', 'oleaut32.dll', 'opengl32.dll', 'rpcrt4.dll', 'setupapi.dll', 'shell32.dll',
    'shlwapi.dll', 'user32.dll', 'userenv.dll', 'uxtheme.dll', 'version.dll', 'winmm.dll', 'winspool.drv',
    'ws2_32.dll', 'wtsapi32.dll', 'bcrypt.dll', 'cabinet.dll', 'iphlpapi.dll', 'secur32.dll', 'shcore.dll',
    'wintrust.dll', 'mswsock.dll', 'dnsapi.dll', 'psapi.dll', 'dbghelp.dll', 'glu32.dll',
}


def check_dlls():
    try:
        import pefile
    except ImportError:
        print('pefile не установлен - проверка DLL пропущена (pip install pefile)')
        return
    files = []
    for root, _, names in os.walk(TARGET):
        files += [os.path.join(root, n) for n in names if n.lower().endswith(('.dll', '.pyd', '.exe'))]
    bundled = {os.path.basename(f).lower() for f in files}
    problems = []
    for f in files:
        pe = pefile.PE(f, fast_load=True)
        pe.parse_data_directories(directories=[pefile.DIRECTORY_ENTRY['IMAGE_DIRECTORY_ENTRY_IMPORT']])
        for entry in getattr(pe, 'DIRECTORY_ENTRY_IMPORT', []):
            dll = entry.dll.decode().lower()
            if dll in bundled or dll in SYSTEM_DLLS or dll.startswith(('api-ms-win-', 'ext-ms-')):
                continue
            problems.append('%s -> %s' % (os.path.relpath(f, TARGET), dll))
        pe.close()
    if problems:
        sys.exit('Не хватает библиотек:\n' + '\n'.join(problems))
    print('Проверка DLL: все зависимости на месте (%d файлов)' % len(files))


def main():
    install_python()
    install_pyqt()
    trim_pyqt()
    check_dlls()
    print('Готово:', TARGET)


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--check':
        check_dlls()
    else:
        main()
