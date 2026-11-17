"""Общие функции точного предпросмотра; без зависимостей от Qt и Office API Python."""
import base64
import hashlib
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import zipfile

SCRIPT = Path(__file__).with_name('word_preview.ps1')


def document_digest(data):
    """Игнорируем время упаковки ZIP: оно не влияет на вёрстку DOCX."""
    digest = hashlib.sha256()
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for name in sorted(archive.namelist()):
            digest.update(name.encode('utf-8'))
            digest.update(b'\0')
            digest.update(archive.read(name))
    return digest.hexdigest()


def powershell_command(script=None):
    if sys.platform != 'win32':
        raise RuntimeError('Точный предпросмотр доступен в Windows с установленным Microsoft Word.')
    exe = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
    if not exe.is_file():
        raise RuntimeError('Не найден Windows PowerShell. Проверьте установку Windows PowerShell 5.1.')
    code = SCRIPT.read_text(encoding='utf-8') if script is None else script
    return [str(exe), '-NoProfile', '-NonInteractive', '-STA', '-WindowStyle', 'Hidden', '-EncodedCommand',
            base64.b64encode(code.encode('utf-16-le')).decode('ascii')]


def read_pages(directory):
    root = Path(directory).resolve()
    result = json.loads((root / 'result.json').read_text(encoding='utf-8-sig'))
    if result.get('error'):
        raise RuntimeError(result['error'])
    pages = []
    for name in result.get('pages', []):
        path = (root / name).resolve()
        if path.parent != root or path.suffix.lower() != '.png' or not path.is_file():
            raise RuntimeError('Word не вернул изображение страницы.')
        with path.open('rb') as file:
            signature = file.read(8)
        if signature != b'\x89PNG\r\n\x1a\n':
            raise RuntimeError('Повреждено изображение страницы предпросмотра.')
        pages.append(path)
    if not pages:
        raise RuntimeError('Word не вернул страницы документа.')
    return pages


def stop_owned_word(directory):
    """Закрываем только новый WINWORD с проверенным PID и временем запуска.

    Документы пользователя и Word, запущенный до задания, не затрагиваются.
    """
    marker = Path(directory) / 'owned-word.json'
    if not marker.is_file() or sys.platform != 'win32':
        return
    code = r'''
$ErrorActionPreference = 'Stop'
$owner = Get-Content -LiteralPath $env:SZ_PREVIEW_OWNER -Raw | ConvertFrom-Json
$process = Get-Process -Id ([int]$owner.pid) -ErrorAction SilentlyContinue
if ($process -and $process.ProcessName -eq 'WINWORD' -and
    $process.StartTime.ToUniversalTime().Ticks -eq [long]$owner.start_ticks) {
    Stop-Process -Id $process.Id -Force
}
'''
    env = os.environ.copy()
    env['SZ_PREVIEW_OWNER'] = str(marker)
    try:
        subprocess.run(powershell_command(code), env=env, timeout=8,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0), check=False)
    except (OSError, subprocess.TimeoutExpired, RuntimeError):
        pass
