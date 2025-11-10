# -*- coding: utf-8 -*-
<#
  Sync-VBA.ps1  v2025.11.11r10
  Автоматизация экспорта и импорта VBA с автодетектом кодировок,
  логом, прогрессбаром и безопасным управлением Excel.
#>

param(
    [int]$Mode = 0,
    [string]$ProjectPath = (Get-Location)
)

# ---------- Лог ----------
$LogFile = Join-Path $ProjectPath "SyncVBA.log"
function Write-Log([string]$msg,[ConsoleColor]$color="Gray"){
    $ts=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line="[$ts] $msg"
    Write-Host $line -ForegroundColor $color
    Add-Content -Encoding UTF8 -Path $LogFile -Value $line
}
Add-Content -Encoding UTF8 -Path $LogFile -Value "`n=== Run $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

# ---------- Кодировки ----------
chcp 65001 > $null
[Console]::InputEncoding=[System.Text.Encoding]::UTF8
[Console]::OutputEncoding=[System.Text.Encoding]::UTF8

# ---------- Пасхалка ----------
$today = Get-Date
if($today.Month -eq 11 -and $today.Day -eq 11){
    Write-Log "🎂 С днём рождения, инженер Александр!" Magenta
    [console]::beep(880,150); [console]::beep(988,150); [console]::beep(1047,250)
}

# ---------- Вспомогательные функции ----------
function Test-Mojibake([string]$s){
    if([string]::IsNullOrEmpty($s)){return $false}
    return ($s -match '[ÃÐÑâ€“â€”â€œâ€â€˜â€™¢™€]')
}
function Fix-Mojibake([string]$s){
    $bytes=[Text.Encoding]::GetEncoding(1252).GetBytes($s)
    return [Text.Encoding]::UTF8.GetString($bytes)
}
function Write-UTF8BOM([string]$path,[string]$text){
    $utf8bom=New-Object System.Text.UTF8Encoding($true)
    [IO.File]::WriteAllText($path,$text,$utf8bom)
}
function Preview-FirstLines([string]$text,[int]$n=8){
    $lines=($text -split "`r?`n")[0..([Math]::Min($n,(($text -split "`r?`n").Count))-1)]
    return ($lines -join "`n")
}
function Convert-TextFile-ToUtf8Bom([string]$path){
    if(!(Test-Path $path)){return}
    $ansi=[Text.Encoding]::Default
    $raw=[IO.File]::ReadAllText($path,$ansi)
    if(Test-Mojibake $raw){$raw=Fix-Mojibake $raw}
    Write-UTF8BOM $path $raw
}

# ---------- Перезапуск в 32-бит ----------
if([Environment]::Is64BitProcess -and (Test-Path "$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0\powershell.exe")){
    Write-Log "Launching 32-bit PowerShell..." Yellow
    $wow="$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    Start-Process -FilePath $wow -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`" -Mode $Mode -ProjectPath `"$ProjectPath`"" -Wait
    Pause
    exit
}

# ---------- Информация о среде ----------
$psArch=if([Environment]::Is64BitProcess){"x64"}else{"x86"}
$excelArch=""
try{
    $key="HKLM:\SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot"
    if(Test-Path $key){
        $path=(Get-ItemProperty $key).Path
        $excelArch=if($path -match "Program Files \(x86\)"){"x86"}else{"x64"}
    }
}catch{}
Write-Log "📊 Среда: Excel=$excelArch, PowerShell=$psArch"

# ---------- Меню ----------
if($Mode -eq 0){
    Write-Host "`n 1-Экспорт  2-Импорт  3-Оба  4-KillExcel" -ForegroundColor Cyan
    $Mode=Read-Host "Введите режим"
}
switch($Mode){
    1{Write-Log "🚀 Режим: ЭКСПОРТ" Cyan}
    2{Write-Log "🚀 Режим: ИМПОРТ" Cyan}
    3{Write-Log "🚀 Режим: ЭКСПОРТ+ИМПОРТ" Cyan}
    4{
        Write-Log "💀 Завершаем все процессы Excel..." Yellow
        Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
        Write-Log "✅ Все экземпляры Excel завершены." Green
        Pause
        exit
    }
    default{Write-Log "❌ Неизвестный режим." Red; Pause; exit}
}

# ---------- Подключение к Excel ----------
Write-Log "🧭 Поиск активного Excel..."
try{
    $excel=[Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    Write-Log "📎 Подключились к активному Excel." Green
}catch{
    $running=Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    if($running){
        Write-Log "⚠ Excel запущен, но COM недоступен — ждём..." Yellow
        Start-Sleep -Seconds 2
        try{
            $excel=[Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            Write-Log "📎 Подключились после ожидания." Green
        }catch{
            Write-Log "⚠ COM не отвечает — создаём новый Excel." Yellow
            $excel=New-Object -ComObject Excel.Application
            $excel.Visible=$true
        }
    }else{
        Write-Log "⚠ Excel не найден — создаём новый экземпляр." Yellow
        $excel=New-Object -ComObject Excel.Application
        $excel.Visible=$true
    }
}
$excel.DisplayAlerts=$false
$excel.EnableEvents=$false
$excel.ScreenUpdating=$false
$excel.Interactive=$false

# ---------- Определение книги ----------
$books=@($excel.Workbooks)
if($books.Count -eq 0){
    $xlsm=Get-ChildItem -Path $ProjectPath -Filter *.xlsm -ErrorAction SilentlyContinue
    if($xlsm.Count -eq 1){
        $wb=$excel.Workbooks.Open($xlsm.FullName)
    }elseif($xlsm.Count -gt 1){
        Write-Host "`nНайдено несколько файлов Excel:" -ForegroundColor Yellow
        for($i=0;$i -lt $xlsm.Count;$i++){Write-Host "  $($i+1). $($xlsm[$i].Name)"}
        $sel=Read-Host "Введите номер файла"
        $wb=$excel.Workbooks.Open($xlsm[[int]$sel-1].FullName)
    }else{
        Write-Host "Нет открытых книг и .xlsm не найдено. Укажи путь:" -ForegroundColor Yellow
        $path=Read-Host "Полный путь к .xlsm"
        $wb=$excel.Workbooks.Open($path)
    }
}elseif($books.Count -eq 1){
    $wb=$books.Item(1)
}else{
    Write-Host "`nНайдено несколько открытых книг:"
    for($i=0;$i -lt $books.Count;$i++){Write-Host "  $($i+1). $($books[$i].Name)"}
    $sel=Read-Host "Введите номер файла"
    $wb=$books.Item([int]$sel)
}
$wbName=$wb.Name
Write-Log "📘 Активная книга: $wbName"

# ---------- Папка для модулей ----------
$ExportPath=Join-Path $ProjectPath "VBA"
if(!(Test-Path $ExportPath)){New-Item -ItemType Directory -Path $ExportPath|Out-Null}

# ---------- ЭКСПОРТ ----------
if($Mode -in 1,3){
    Write-Log ">>> Экспорт VBA-компонентов..."
    $vbComps=@($wb.VBProject.VBComponents | Where-Object { $_.Type -ne 100 })
    $total=$vbComps.Count; $i=0
    foreach($vbComp in $vbComps){
        $i++; $p=[int](($i/$total)*100)
        Write-Progress -Activity "Экспорт VBA" -Status "$($vbComp.Name)" -PercentComplete $p
        try{
            switch($vbComp.Type){
                1 { $ext=".bas" } 2 { $ext=".cls" } 3 { $ext=".frm" } default { $ext=".bas" }
            }
            $target=Join-Path $ExportPath ($vbComp.Name+$ext)
            if($ext -in ".bas",".cls"){
                $lines=$vbComp.CodeModule.CountOfLines
                if($lines -gt 0){
                    $raw=$vbComp.CodeModule.Lines(1,$lines)
                    $text=$raw
                    if(Test-Mojibake $text){
                        Write-Log "⚠ Кракозябры в $($vbComp.Name) — перекодировка..." Yellow
                        $text=Fix-Mojibake $text
                    }
                    Write-UTF8BOM $target $text
                    $prev=Preview-FirstLines $text 6
                    Write-Log "✔ Exported: $($vbComp.Name)$ext" Green
                    Write-Host $prev -ForegroundColor DarkGray
                }
            }else{
                $vbComp.Export($target)
                Convert-TextFile-ToUtf8Bom $target
                $frx=[IO.Path]::ChangeExtension($target,".frx")
                if(Test-Path $frx){Copy-Item $frx -Dest $ExportPath -Force}
                Write-Log "✔ Exported form: $($vbComp.Name)$ext" Green
            }
        }catch{
            Write-Log ("⚠ Ошибка при экспорте "+$vbComp.Name+": "+$_.Exception.Message) Red
        }
    }
    Write-Progress -Activity "Экспорт VBA" -Completed
    Write-Log "✅ Все модули успешно экспортированы." Cyan
    if(Get-Command code -ErrorAction SilentlyContinue){
        Start-Process "code" -ArgumentList "-r `"$ExportPath`""
        Write-Log "📂 Открыт каталог в текущем окне VS Code." DarkGray
    }

    if ($Mode -eq 1) {
        $excel.Interactive = $true
        Write-Host "`n=== Работа завершена. Нажми любую клавишу для выхода... ===" -ForegroundColor Gray
        Pause
        return
    }
}

# ---------- ИМПОРТ ----------
if($Mode -in 2,3){
    Write-Log ">>> Импорт VBA..."
    $files=Get-ChildItem -Path $ExportPath -Include *.bas,*.cls,*.frm -ErrorAction SilentlyContinue
    foreach($file in $files){
        $name=[IO.Path]::GetFileNameWithoutExtension($file)
        $ext=$file.Extension.ToLower()
        try{
            if($ext -in ".bas",".cls"){
                $text=Get-Content -Raw -Encoding UTF8 $file
                if(Test-Mojibake $text){$text=Fix-Mojibake $text}
                $vbComp=$wb.VBProject.VBComponents|Where-Object{$_.Name -eq $name}
                if(-not $vbComp){$vbComp=$wb.VBProject.VBComponents.Add(1);$vbComp.Name=$name}
                $vbComp.CodeModule.DeleteLines(1,$vbComp.CodeModule.CountOfLines)
                $vbComp.CodeModule.AddFromString($text)
                Write-Log "✔ Импортирован: $name$ext" Green
            }else{
                $old=$wb.VBProject.VBComponents|?{$_.Name -eq $name}
                if($old){try{$wb.VBProject.VBComponents.Remove($old)}catch{}}
                $wb.VBProject.VBComponents.Import($file.FullName)
                $frx=[IO.Path]::ChangeExtension($file.FullName,".frx")
                if(Test-Path $frx){
                    $tfrx=Join-Path ([IO.Path]::GetDirectoryName($wb.FullName)) ($file.BaseName+".frx")
                    Copy-Item $frx -Dest $tfrx -Force
                }
                Write-Log "✔ Импортирован: $name$ext" Green
            }
        }catch{
            Write-Log ("⚠ Ошибка при импорте "+$name+$ext+": "+$_.Exception.Message) Red
        }
    }
    $wb.Save()
    Write-Log "✅ Импорт завершён, книга сохранена." Cyan
}

# ---------- Финал ----------
$excel.Interactive = $true
Write-Host "`n=== Работа завершена. Нажми любую клавишу для выхода... ===" -ForegroundColor Gray
Pause
