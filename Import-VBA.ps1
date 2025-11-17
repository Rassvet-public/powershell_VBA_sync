# -*- coding: utf-8 -*-
<#
    Import-VBA.ps1
    Узкоспециализированный запуск: только импорт VBA-модулей
    из папки VBA в указанную Excel-книгу.
#>

param(
    [string]$ProjectPath = (Get-Location)
)

<#
    Многострочный комментарий:
    1) Подключаем основной Sync-VBA.ps1.
    2) Вызываем Invoke-SyncVbaMain в режиме 2 (импорт).
#>

$scriptRoot = Split-Path -Parent $PSCommandPath
$mainPath   = Join-Path -Path $scriptRoot -ChildPath 'Sync-VBA.ps1'

if (-not (Test-Path -Path $mainPath)) {
    Write-Host "Не найден основной скрипт Sync-VBA.ps1 рядом с Import-VBA.ps1" -ForegroundColor Red
    Pause
    exit 1
}

. $mainPath

Invoke-SyncVbaMain -Mode 2 -ProjectPath $ProjectPath
