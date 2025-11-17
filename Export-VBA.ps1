# -*- coding: utf-8 -*-
<#
    Export-VBA.ps1
    Узкоспециализированный запуск: только экспорт VBA-модулей
    из Excel-книги в папку VBA.
#>

param(
    [string]$ProjectPath = (Get-Location)
)

<#
    Многострочный комментарий:
    1) Определяем путь к основному скрипту Sync-VBA.ps1.
    2) Подключаем его как библиотеку (dot-source).
    3) Вызываем Invoke-SyncVbaMain в режиме 1 (экспорт).
#>

$scriptRoot = Split-Path -Parent $PSCommandPath
$mainPath   = Join-Path -Path $scriptRoot -ChildPath 'Sync-VBA.ps1'

if (-not (Test-Path -Path $mainPath)) {
    Write-Host "Не найден основной скрипт Sync-VBA.ps1 рядом с Export-VBA.ps1" -ForegroundColor Red
    Pause
    exit 1
}

. $mainPath

Invoke-SyncVbaMain -Mode 1 -ProjectPath $ProjectPath
