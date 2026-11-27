"""Открывает штатный диалог печати Word в отдельном процессе."""
import base64
import os
import subprocess

SUPPORTED = os.name == 'nt'

# Show() применяет настройки диалога только после подтверждения пользователем.
PRINT_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
$result = 1
try {
    $word = New-Object -ComObject Word.Application
    $result = 2
    $word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    $word.Visible = $true
    # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
    $doc = $word.Documents.Open($env:SZ_PRINT_DOCX, $false, $true, $false)
    $doc.Activate()
    $dialog = $word.Dialogs.Item(88)  # wdDialogFilePrint
    if ($dialog.Show() -eq -1) {
        # Даём Word передать задание в очередь до закрытия открытой копии.
        while ($word.BackgroundPrintingStatus -gt 0) {
            Start-Sleep -Milliseconds 200
        }
    }
    $result = 0
} catch {
    # Код завершения передаёт приложению этап ошибки.
} finally {
    if ($null -ne $doc) {
        try { $doc.Close(0) } catch {}
    }
    if ($null -ne $word) {
        try {
            if ($word.Documents.Count -eq 0) { $word.Quit(0) }
        } catch {}
    }
}
exit $result
"""


def start(path):
    """Запускает Windows PowerShell без окна консоли и без ожидания в интерфейсе."""
    root = os.environ.get('SystemRoot', r'C:\Windows')
    executable = os.path.join(root, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
    command = base64.b64encode(PRINT_SCRIPT.encode('utf-16-le')).decode('ascii')
    env = dict(os.environ, SZ_PRINT_DOCX=os.path.abspath(path))
    return subprocess.Popen(
        [executable, '-NoProfile', '-NonInteractive', '-STA', '-EncodedCommand', command],
        env=env, creationflags=subprocess.CREATE_NO_WINDOW,
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
