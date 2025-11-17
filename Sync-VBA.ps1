# -*- coding: utf-8 -*-
<#
    Sync-VBA.ps1  v2025.11.17r1
    Основной скрипт: экспорт / импорт VBA с поддержкой x86 Excel,
    автодетектом кодировок и аккуратным управлением Excel.
#>

param(
    [int]$Mode = 0,
    [string]$ProjectPath = (Get-Location)
)

# Глобальный путь к лог-файлу для всех функций
$script:SyncVba_LogFile = $null

<# ======================= ЛОГИРОВАНИЕ ======================= #>
function Write-Log {
    <#
        Многострочный комментарий:
        Функция логирования в консоль и файл SyncVBA.log.
        Цветной вывод в консоль, в файл пишем всегда серым текстом.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[{0}] {1}" -f $timestamp, $Message

    Write-Host $line -ForegroundColor $Color

    if ($script:SyncVba_LogFile) {
        Add-Content -Encoding UTF8 -Path $script:SyncVba_LogFile -Value $line
    }
}

<# ======================= КОДИРОВКИ / MOJIBAKE ======================= #>
function Test-Mojibake {
    <#
        Проверка строки на типичные "кракозябры" после перепутанной
        UTF-8 / ANSI кодировки.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return $false
    }

    return ($Text -match '[ÃÐÑâ€“â€”â€œâ€â€˜â€™¢™€]')
}

function Fix-Mojibake {
    <#
        Попытка "починить" кракозябры:
        считаем, что текст ошибочно прочитали как 1252 вместо UTF-8,
        и перекодируем обратно.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $bytes = [System.Text.Encoding]::GetEncoding(1252).GetBytes($Text)
    return [System.Text.Encoding]::UTF8.GetString($bytes)
}

function Write-UTF8BOM {
    <#
        Запись текста в файл с явным UTF-8 BOM,
        как требует твой проект.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $encoding = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $Text, $encoding)
}

function Convert-TextFile-ToUtf8Bom {
    <#
        Чтение текстового файла в системной ANSI-кодировке,
        починка кракозябр (если есть) и запись в UTF-8 BOM.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        return
    }

    $ansiEncoding = [System.Text.Encoding]::Default
    $raw = [System.IO.File]::ReadAllText($Path, $ansiEncoding)
    if (Test-Mojibake -Text $raw) {
        $raw = Fix-Mojibake -Text $raw
    }

    Write-UTF8BOM -Path $Path -Text $raw
}

<# ======================= СРЕДА / x86 ПЕРЕЗАПУСК ======================= #>
function Write-EnvironmentInfo {
    <#
        Логируем архитектуру PowerShell и Excel.
        Используем ключ реестра Excel 2016 (16.0).
    #>
    [CmdletBinding()]
    param()

    $psArch = if ([Environment]::Is64BitProcess) { "x64" } else { "x86" }
    $excelArch = ""

    try {
        $key = "HKLM:\SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot"
        if (Test-Path -Path $key) {
            $path = (Get-ItemProperty -Path $key).Path
            $excelArch = if ($path -match "Program Files \(x86\)") { "x86" } else { "x64" }
        }
    }
    catch {
        # Тихая ошибка, просто не знаем архитектуру Excel
    }

    Write-Log ("📊 Среда: Excel={0}, PowerShell={1}" -f $excelArch, $psArch)
}

function Stop-ExcelAll {
    <#
        Принудительное завершение всех процессов Excel.
        Используется в режиме KillExcel.
    #>
    [CmdletBinding()]
    param()

    Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
}

function Invoke-32BitSelf {
    <#
        Перезапуск текущего скрипта в 32-битной версии PowerShell
        для работы с 32-битным Excel.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Mode,

        [Parameter(Mandatory = $true)]
        [string]$ProjectPath
    )

    if (-not [Environment]::Is64BitProcess) {
        return
    }

    $wowPath = Join-Path -Path $env:SystemRoot -ChildPath 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -Path $wowPath)) {
        return
    }

    Write-Log "Перезапуск в 32-битной версии PowerShell для совместимости с Excel..." ([ConsoleColor]::Yellow)

    $argumentList = @(
        '-NoExit',
		'-ExecutionPolicy', 'Bypass',
        '-NoProfile',
        '-File', "`"$PSCommandPath`"",
        '-Mode', $Mode,
        '-ProjectPath', "`"$ProjectPath`""
    )

    Start-Process -FilePath $wowPath -ArgumentList $argumentList -Wait
    exit
}

<# ======================= ВЫБОР КНИГИ EXCEL ======================= #>
function Select-Workbook {
    <#
        Логика выбора Excel-книги:
        1) если книги уже открыты — даём выбрать;
        2) если нет — ищем .xlsm в ProjectPath;
        3) если не нашли — просим путь руками.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Excel,

        [Parameter(Mandatory = $true)]
        [string]$ProjectPath
    )

    $workbooks = @($Excel.Workbooks)

    if ($workbooks.Count -eq 0) {
        $xlsmFiles = Get-ChildItem -Path $ProjectPath -Filter '*.xlsm' -ErrorAction SilentlyContinue
        if ($xlsmFiles.Count -eq 1) {
            return $Excel.Workbooks.Open($xlsmFiles.FullName)
        }
        elseif ($xlsmFiles.Count -gt 1) {
            Write-Host "`nНайдено несколько файлов Excel:" -ForegroundColor Yellow
            for ($index = 0; $index -lt $xlsmFiles.Count; $index++) {
                Write-Host ("  {0}. {1}" -f ($index + 1), $xlsmFiles[$index].Name)
            }
            $selection = Read-Host "Введите номер файла"
            $selectedIndex = [int]$selection - 1
            return $Excel.Workbooks.Open($xlsmFiles[$selectedIndex].FullName)
        }
        else {
            Write-Host "Нет открытых книг и .xlsm не найдено. Укажи путь:" -ForegroundColor Yellow
            $path = Read-Host "Полный путь к .xlsm"
            return $Excel.Workbooks.Open($path)
        }
    }
    elseif ($workbooks.Count -eq 1) {
        return $workbooks.Item(1)
    }
    else {
        Write-Host "`nНайдено несколько открытых книг:" -ForegroundColor Yellow
        for ($index = 0; $index -lt $workbooks.Count; $index++) {
            Write-Host ("  {0}. {1}" -f ($index + 1), $workbooks[$index].Name)
        }
        $selection = Read-Host "Введите номер файла"
        $selectedIndex = [int]$selection
        return $workbooks.Item($selectedIndex)
    }
}

<# ======================= СЕССИЯ EXCEL ======================= #>
function Start-ExcelSession {
    <#
        Создаём или находим Excel, выбираем книгу,
        подготавливаем папку VBA и возвращаем объект с параметрами сессии.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectPath
    )

    Write-Log "🧭 Поиск активного Excel..." ([ConsoleColor]::Gray)

    $excel = $null
    $createdNewExcel = $false

    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        Write-Log "📎 Подключились к активному Excel." ([ConsoleColor]::Green)
    }
    catch {
        $running = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
        if ($running) {
            Write-Log "⚠ Excel запущен, но COM недоступен — создаём новый экземпляр." ([ConsoleColor]::Yellow)
        }
        else {
            Write-Log "⚠ Excel не найден — создаём новый экземпляр." ([ConsoleColor]::Yellow)
        }
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $createdNewExcel = $true
    }

    $excel.DisplayAlerts  = $false
    $excel.EnableEvents   = $false
    $excel.ScreenUpdating = $false
    $excel.Interactive    = $false

    $workbook = Select-Workbook -Excel $excel -ProjectPath $ProjectPath
    $workbookName      = $workbook.Name
    $workbookBaseName  = [System.IO.Path]::GetFileNameWithoutExtension($workbookName)

    Write-Log ("📘 Активная книга: {0}" -f $workbookName)

    $exportPath = Join-Path -Path $ProjectPath -ChildPath 'VBA'
    if (-not (Test-Path -Path $exportPath)) {
        New-Item -ItemType Directory -Path $exportPath | Out-Null
    }

    return [pscustomobject]@{
        Excel             = $excel
        Workbook          = $workbook
        WorkbookName      = $workbookName
        WorkbookBaseName  = $workbookBaseName
        ProjectPath       = $ProjectPath
        ExportPath        = $exportPath
        CreatedNewExcel   = $createdNewExcel
    }
}


function Stop-ExcelSession {
    <#
        Аккуратно сохраняем книгу, восстанавливаем параметры Excel,
        закрываем созданный экземпляр (если мы его создавали).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    $excel          = $Session.Excel
    $workbook       = $Session.Workbook
    $createdNewExcel = $Session.CreatedNewExcel

    try {
        if ($workbook -ne $null) {
            Write-Log "💾 Сохраняем книгу..." ([ConsoleColor]::Gray)
            $workbook.Save()
        }
    }
    catch {
        Write-Log ("⚠ Ошибка при сохранении книги: {0}" -f $_.Exception.Message) ([ConsoleColor]::Red)
    }
    finally {
        if ($workbook -ne $null) {
            try {
                $workbook.Close($true) | Out-Null
            }
            catch { }
        }

        if ($excel -ne $null) {
            try {
                $excel.DisplayAlerts  = $true
                $excel.EnableEvents   = $true
                $excel.ScreenUpdating = $true
                $excel.Interactive    = $true

                if ($createdNewExcel) {
                    $excel.Quit()
                }
            }
            catch { }

            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}

<# ======================= ЭКСПОРТ / ИМПОРТ ======================= #>
function Export-VBAModules {
    <#
        Экспорт всех модулей, классов и форм VBA
        в папку VBA с сохранением UTF-8 BOM.
        Имена файлов: <ИмяКнигиБезРасширения>_<ИмяМодуля>.bas/.cls/.frm
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    $workbook          = $Session.Workbook
    $exportPath        = $Session.ExportPath
    $workbookBaseName  = $Session.WorkbookBaseName

    Write-Log ">>> Экспорт VBA-компонентов..." ([ConsoleColor]::Gray)

    $vbComponents = @($workbook.VBProject.VBComponents | Where-Object { $_.Type -ne 100 })
    $total = $vbComponents.Count
    $index = 0

    foreach ($vbComponent in $vbComponents) {
        $index++
        $percent = if ($total -gt 0) { [int](($index / $total) * 100) } else { 0 }

        Write-Progress -Activity "Экспорт VBA" -Status $vbComponent.Name -PercentComplete $percent

        try {
            switch ($vbComponent.Type) {
                1 { $extension = ".bas" }  # стандартный модуль
                2 { $extension = ".cls" }  # класс
                3 { $extension = ".frm" }  # форма
                default { $extension = ".bas" }
            }

            # Имя файла: <ИмяКнигиБезРасширения>_<ИмяМодуля>.ext
            $fileName   = "{0}_{1}{2}" -f $workbookBaseName, $vbComponent.Name, $extension
            $targetPath = Join-Path -Path $exportPath -ChildPath $fileName

            if ($extension -in @(".bas", ".cls")) {
                $lineCount = $vbComponent.CodeModule.CountOfLines
                if ($lineCount -gt 0) {
                    $codeText = $vbComponent.CodeModule.Lines(1, $lineCount)
                    if (Test-Mojibake -Text $codeText) {
                        $codeText = Fix-Mojibake -Text $codeText
                    }
                    Write-UTF8BOM -Path $targetPath -Text $codeText
                    Write-Log ("✔ Экспортирован модуль: {0}" -f $fileName) ([ConsoleColor]::Green)
                }
            }
            elseif ($extension -eq ".frm") {
                $vbComponent.Export($targetPath)
                Convert-TextFile-ToUtf8Bom -Path $targetPath
                Write-Log ("✔ Экспортирована форма: {0}" -f $fileName) ([ConsoleColor]::Green)
            }
        }
        catch {
            Write-Log ("⚠ Ошибка при экспорте {0}: {1}" -f $vbComponent.Name, $_.Exception.Message) ([ConsoleColor]::Red)
        }
    }

    Write-Progress -Activity "Экспорт VBA" -Completed -Status "Готово"
    Write-Log "✅ Все модули успешно экспортированы." ([ConsoleColor]::Cyan)
}


function Import-VBAModules {
    <#
        Импорт модулей, классов и форм VBA из папки VBA в книгу.
        Ожидаемый шаблон имени файла:
        <ИмяКнигиБезРасширения>_<ИмяМодуля>.bas/.cls/.frm
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Session
    )

    $workbook          = $Session.Workbook
    $exportPath        = $Session.ExportPath
    $workbookBaseName  = $Session.WorkbookBaseName

    Write-Log ">>> Импорт VBA-компонентов..." ([ConsoleColor]::Gray)

    # Берём только файлы, относящиеся к данной книге
    $patternBas = "{0}_*.bas" -f $workbookBaseName
    $patternCls = "{0}_*.cls" -f $workbookBaseName
    $patternFrm = "{0}_*.frm" -f $workbookBaseName

    $files = Get-ChildItem -Path $exportPath -Include $patternBas, $patternCls, $patternFrm -File -ErrorAction SilentlyContinue

    foreach ($file in $files) {
        $fileBaseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $extension    = $file.Extension.ToLowerInvariant()
        $prefix       = "{0}_" -f $workbookBaseName

        if (-not $fileBaseName.StartsWith($prefix)) {
            continue
        }

        # Имя модуля = всё после "<ИмяКниги>_"
        $moduleName = $fileBaseName.Substring($prefix.Length)
        if ([string]::IsNullOrWhiteSpace($moduleName)) {
            continue
        }

        try {
            if ($extension -in @(".bas", ".cls")) {
                $text = Get-Content -Raw -Encoding UTF8 -Path $file.FullName
                if (Test-Mojibake -Text $text) {
                    $text = Fix-Mojibake -Text $text
                }

                $vbComponent = $workbook.VBProject.VBComponents | Where-Object { $_.Name -eq $moduleName }
                if (-not $vbComponent) {
                    $componentType = if ($extension -eq ".cls") { 2 } else { 1 }
                    $vbComponent = $workbook.VBProject.VBComponents.Add($componentType)
                    $vbComponent.Name = $moduleName
                }

                $codeModule = $vbComponent.CodeModule
                $linesCount = $codeModule.CountOfLines
                if ($linesCount -gt 0) {
                    $codeModule.DeleteLines(1, $linesCount)
                }
                $codeModule.AddFromString($text)

                Write-Log ("✔ Импортирован модуль: {0} ({1})" -f $moduleName, $file.Name) ([ConsoleColor]::Green)
            }
            elseif ($extension -eq ".frm") {
                $existing = $workbook.VBProject.VBComponents | Where-Object { $_.Name -eq $moduleName }
                if ($existing) {
                    $workbook.VBProject.VBComponents.Remove($existing)
                }

                $null = $workbook.VBProject.VBComponents.Import($file.FullName)

                $frxPath = [System.IO.Path]::ChangeExtension($file.FullName, ".frx")
                if (Test-Path -Path $frxPath) {
                    $targetFrx = Join-Path -Path ([System.IO.Path]::GetDirectoryName($workbook.FullName)) -ChildPath ([System.IO.Path]::GetFileName($frxPath))
                    Copy-Item -Path $frxPath -Destination $targetFrx -Force
                }

                Write-Log ("✔ Импортирована форма: {0} ({1})" -f $moduleName, $file.Name) ([ConsoleColor]::Green)
            }
        }
        catch {
            Write-Log ("⚠ Ошибка при импорте {0} из {1}: {2}" -f $moduleName, $file.Name, $_.Exception.Message) ([ConsoleColor]::Red)
        }
    }

    Write-Log "✅ Импорт завершён, книга сохранена." ([ConsoleColor]::Cyan)
}

<# ======================= VS CODE / ФИНАЛ ======================= #>
ffunction Open-VbaInEditor {
    <#
        Открываем результаты экспорта:
        1) Если установлен Notepad++ по пути C:\Program Files\Notepad++\notepad++.exe,
           открываем в нём все .bas-файлы текущей книги.
        2) Если Notepad++ не найден или .bas нет,
           просто открываем папку VBA в Проводнике.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExportPath,

        [Parameter(Mandatory = $true)]
        [string]$WorkbookBaseName
    )

    $notepadPath = "C:\Program Files\Notepad++\notepad++.exe"

    $patternBas = "{0}_*.bas" -f $WorkbookBaseName
    $basFiles = Get-ChildItem -Path $ExportPath -Filter $patternBas -File -ErrorAction SilentlyContinue

    if ((Test-Path -Path $notepadPath) -and $basFiles -and $basFiles.Count -gt 0) {
        # Открываем все соответствующие .bas в Notepad++
        $args = $basFiles.FullName
        Start-Process -FilePath $notepadPath -ArgumentList $args
        Write-Log ("📄 Открыто в Notepad++ файлов: {0}" -f $basFiles.Count) ([ConsoleColor]::DarkGray)
    }
    else {
        # Fallback: просто открыть папку в Проводнике
        Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$ExportPath`""
        Write-Log "📂 Открыт каталог VBA в Проводнике." ([ConsoleColor]::DarkGray)
    }
}

function Show-FinishMessage {
    <#
        Финальное сообщение и пауза, чтобы окно не закрывалось мгновенно.
    #>
    [CmdletBinding()]
    param()

    Write-Host "`n=== Работа завершена. Нажми любую клавишу для выхода... ===" -ForegroundColor Gray
    Pause
}

<# ======================= ГЛАВНАЯ ФУНКЦИЯ ======================= #>
function Invoke-SyncVbaMain {
    <#
        Главная точка входа: меню, перезапуск в x86, запуск экспорта/импорта.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Mode,

        [Parameter(Mandatory = $true)]
        [string]$ProjectPath
    )

    $script:SyncVba_LogFile = Join-Path -Path $ProjectPath -ChildPath 'SyncVBA.log'
    Add-Content -Encoding UTF8 -Path $script:SyncVba_LogFile -Value "`n=== Run $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

    chcp 65001 > $null
    [Console]::InputEncoding  = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

    $today = Get-Date
    if ($today.Month -eq 11 -and $today.Day -eq 11) {
        Write-Log "🎂 С днём рождения, инженер Александр!" ([ConsoleColor]::Magenta)
        [Console]::Beep(880,150); [Console]::Beep(988,150); [Console]::Beep(1047,250)
    }

    Write-EnvironmentInfo

    $effectiveMode = $Mode
    if ($effectiveMode -eq 0) {
        Write-Host "`n 1-Экспорт  2-Импорт  3-Оба  4-KillExcel" -ForegroundColor Cyan
        $inputValue = Read-Host "Введите режим"
        [void][int]::TryParse($inputValue, [ref]$effectiveMode)
    }

    switch ($effectiveMode) {
        1 { Write-Log "🚀 Режим: ЭКСПОРТ" ([ConsoleColor]::Cyan) }
        2 { Write-Log "🚀 Режим: ИМПОРТ" ([ConsoleColor]::Cyan) }
        3 { Write-Log "🚀 Режим: ЭКСПОРТ+ИМПОРТ" ([ConsoleColor]::Cyan) }
        4 {
            Write-Log "💀 Завершаем все процессы Excel..." ([ConsoleColor]::Yellow)
            Stop-ExcelAll
            Write-Log "✅ Все экземпляры Excel завершены." ([ConsoleColor]::Green)
            Show-FinishMessage
            return
        }
        Default {
            Write-Log "❌ Неизвестный режим." ([ConsoleColor]::Red)
            Show-FinishMessage
            return
        }
    }

    Invoke-32BitSelf -Mode $effectiveMode -ProjectPath $ProjectPath

    $session    = Start-ExcelSession -ProjectPath $ProjectPath
    $exportDone = $false
    $importDone = $false

    try {
        if ($effectiveMode -eq 1 -or $effectiveMode -eq 3) {
            Export-VBAModules -Session $session
            $exportDone = $true
        }

        if ($effectiveMode -eq 2 -or $effectiveMode -eq 3) {
            Import-VBAModules -Session $session
            $importDone = $true
        }
    }
    finally {
        Stop-ExcelSession -Session $session
    }

    if ($exportDone -and -not $importDone) {
        Open-VbaInEditor -ExportPath $session.ExportPath -WorkbookBaseName $session.WorkbookBaseName
    }

    Show-FinishMessage
}


# Автоматический запуск только если скрипт запускают как .\Sync-VBA.ps1,
# а не dot-source (в обёртках Export-VBA.ps1 / Import-VBA.ps1).
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-SyncVbaMain -Mode $Mode -ProjectPath $ProjectPath
}
