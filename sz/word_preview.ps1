# Запускается приложением через Windows PowerShell, без сторонних модулей.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$root = $env:SZ_PREVIEW_DIRECTORY
$utf8 = [System.Text.UTF8Encoding]::new($false)
$word = $null
$document = $null
$ownsWord = $false
$result = @{pages = @(); error = $null}

try {
    if (-not $root -or -not (Test-Path -LiteralPath (Join-Path $root 'input.docx'))) {
        throw 'Не найден временный документ предпросмотра.'
    }
    Add-Type -AssemblyName System.Drawing
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class SzPreviewWindow {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
    $previous = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $word = New-Object -ComObject Word.Application
    [uint32]$wordProcessId = 0
    $null = [SzPreviewWindow]::GetWindowThreadProcessId([IntPtr]$word.Hwnd, [ref]$wordProcessId)
    if (-not $wordProcessId -or $previous -contains [int]$wordProcessId) {
        throw 'Не удалось создать отдельный экземпляр Word. Открытые документы не затронуты.'
    }
    $ownsWord = $true
    $process = Get-Process -Id $wordProcessId
    $owner = @{pid = $process.Id; start_ticks = $process.StartTime.ToUniversalTime().Ticks}
    [IO.File]::WriteAllText((Join-Path $root 'owned-word.json'), ($owner | ConvertTo-Json), $utf8)
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3  # не выполнять макросы
    # Открываем локальную копию только для чтения, без списка недавних файлов.
    $document = $word.Documents.Open((Join-Path $root 'input.docx'), $false, $true, $false)
    $document.ActiveWindow.View.Type = 3  # wdPrintView
    $document.ActiveWindow.View.ShowFieldCodes = $false
    $document.ActiveWindow.View.ShowAll = $false
    $document.ShowSpellingErrors = $false
    $document.ShowGrammaticalErrors = $false
    $null = $document.Fields.Update()
    $document.Repaginate()
    foreach ($section in $document.Sections) {
        foreach ($footer in $section.Footers) {
            $null = $footer.Range.Fields.Update()
        }
    }
    $document.Repaginate()
    $pages = $document.ActiveWindow.Panes.Item(1).Pages
    $count = $pages.Count
    for ($i = 1; $i -le $count; $i++) {
        if (Test-Path -LiteralPath (Join-Path $root 'cancel')) {
            throw 'Предпросмотр отменён.'
        }
        # Page.EnhMetaFileBits — изображение страницы, сформированное самим Word.
        $page = $pages.Item($i)
        [byte[]]$bits = $page.EnhMetaFileBits
        $stream = [IO.MemoryStream]::new($bits)
        $metafile = $null
        $bitmap = $null
        $graphics = $null
        try {
            $metafile = [Drawing.Imaging.Metafile]::new($stream)
            $width = [Math]::Max(1, [int][Math]::Ceiling($page.Width * 2))
            $height = [Math]::Max(1, [int][Math]::Ceiling($page.Height * 2))
            $bitmap = [Drawing.Bitmap]::new($width, $height)
            $graphics = [Drawing.Graphics]::FromImage($bitmap)
            $graphics.Clear([Drawing.Color]::White)
            $graphics.DrawImage($metafile, 0, 0, $width, $height)
            $name = 'page-{0:D4}.png' -f $i
            $bitmap.Save((Join-Path $root $name), [Drawing.Imaging.ImageFormat]::Png)
            $result.pages += $name
        }
        finally {
            if ($graphics) { $graphics.Dispose() }
            if ($bitmap) { $bitmap.Dispose() }
            if ($metafile) { $metafile.Dispose() }
            $stream.Dispose()
            $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($page)
        }
    }
}
catch {
    $result.error = 'Не удалось получить страницы Microsoft Word. ' + $_.Exception.Message
    $result.pages = @()
}
finally {
    if ($document) {
        try { $document.Close(0) } catch {}
        $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    }
    if ($word) {
        if ($ownsWord) { try { $word.Quit(0) } catch {} }
        $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [IO.File]::WriteAllText((Join-Path $root 'result.json'), ($result | ConvertTo-Json -Depth 4), $utf8)
}
