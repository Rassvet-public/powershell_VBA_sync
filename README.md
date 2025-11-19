# Sync-VBA PowerShell Toolkit

PowerShell-скрипты для безопасной синхронизации VBA-кода между Excel-книгой и файловой структурой проекта. Набор включает
экспорт, импорт и комбинированную синхронизацию модулей, строгие проверки кодировок и полноценный CI/CD на GitHub Actions.

## Содержание
1. [Назначение проекта](#назначение-проекта)
2. [Структура репозитория](#структура-репозитория)
3. [Скрипты и сценарии](#скрипты-и-сценарии)
4. [Как пользоваться Sync-VBA.ps1](#как-пользоваться-sync-vbaps1)
5. [Инфраструктура CI/CD](#инфраструктура-cicd)
   1. [Триггеры и правила ветвления](#триггеры-и-правила-ветвления)
   2. [Workflow «PSScriptAnalyzer Check v7»](#workflow-psscriptanalyzer-check-v7)
   3. [Workflow «UTF-8 Encoding Check v6»](#workflow-utf-8-encoding-check-v6)
   4. [RuleSet для PSScriptAnalyzer](#ruleset-для-psscriptanalyzer)
6. [Roadmap по CI/CD и качеству](#roadmap-по-cicd-и-качеству)
7. [Требования к окружению](#требования-к-окружению)
8. [Практические советы](#практические-советы)

## Назначение проекта
* Экспорт и импорт модулей, классов и форм VBA в кодировке UTF-8 BOM.
* Автоматическое исправление кракозябр, появившихся из-за ANSI/Windows-1252.
* Управление экземплярами Excel, отключение предупреждений и событий COM.
* Цветной журнал `SyncVBA.log`, предпросмотр первых строк кода при экспорте.
* Опциональное открытие каталога `VBA` в VS Code и удаление старых версий `.frx` при импорте форм.

## Структура репозитория
```
.
├── Export-VBA.ps1             # Односторонний экспорт модулей в каталог VBA
├── Import-VBA.ps1             # Односторонний импорт модулей из каталога VBA
├── Sync-VBA.ps1               # Интерактивная синхронизация (экспорт/импорт/kill Excel)
├── README.md                  # Этот документ
├── RoadMap_powershell_VBA_sync.md # Подробная история настройки CI/CD
├── .github/
│   ├── workflows/
│   │   ├── analyze_ps_v7.yml  # Проверка AST + PSScriptAnalyzer
│   │   └── check_encoding.yml # Контроль кодировок UTF-8/UTF-8 BOM
│   └── PSScriptAnalyzerRuleSet.psd1 # Набор строгих правил анализатора
└── .gitignore
```

## Скрипты и сценарии
| Скрипт          | Что делает | Когда использовать |
|-----------------|------------|--------------------|
| `Sync-VBA.ps1`  | Интерактивно запускает экспорт/импорт VBA, показывает прогресс, ведёт журнал, управляет Excel и кодировками. | Основной сценарий синхронизации перед сборкой или ревью VBA-кода. |
| `Export-VBA.ps1`| Выгружает все модули и формы из выбранной книги в папку `VBA`, гарантируя UTF-8 BOM и чистку Mojibake. | Автономный экспорт для CI/backup. |
| `Import-VBA.ps1`| Загружает модули/формы из папки `VBA` обратно в книгу Excel, копирует `.frx`, удаляет старые версии. | Автономный импорт после ревью или из шаблона. |

## Как пользоваться Sync-VBA.ps1
1. Скопируйте `Sync-VBA.ps1` и вспомогательные скрипты в корень проекта рядом с `.xlsm`.
2. Запустите PowerShell с правами, позволяющими работать с COM (при необходимости `Set-ExecutionPolicy Bypass -Scope Process`).
3. Выполните
   ```powershell
   .\Sync-VBA.ps1
   ```
4. Выберите режим:
   * `1` — экспорт модулей в папку `VBA`.
   * `2` — импорт модулей из `VBA` в книгу.
   * `3` — экспорт, затем импорт (идеально для синхронизации перед сборкой).
   * `4` — завершение всех процессов Excel (`KillExcel`).

Скрипт автоматически ищет `.xlsm` в текущем каталоге. Если нужная книга не найдена, введите путь вручную. После экспорта можно открыть VS Code в каталоге `VBA` для ревью.

## Инфраструктура CI/CD
Цель пайплайна — не пропускать синтаксические ошибки, нарушения правил кодирования и проблемы с кодировками.

### Триггеры и правила ветвления
* Активные ветки: `master` (стабильная), `dev` (основная разработка), `codex` (песочница).
* Оба workflow запускаются на `push` и `pull_request` к указанным веткам.
* Включена защита ветки `master`: обязательные PR, запрет прямых push, требование статуса `PSScriptAnalyzer Check v7 / pssa` и `UTF-8 Encoding Check v6 / encoding_check`, линейная история.

### Workflow `PSScriptAnalyzer Check v7`
Файл: [`.github/workflows/analyze_ps_v7.yml`](.github/workflows/analyze_ps_v7.yml)

| Этап | Что происходит |
|------|----------------|
| Checkout | `actions/checkout@v4` получает код. |
| Установка PowerShell 7 | Скачивается архив PowerShell 7.4.2 x64, добавляется в `PATH`. |
| Вывод версии | Команда `pwsh -c "Write-Host (Get-Host).Version"`. |
| Установка PSScriptAnalyzer | `Install-Module PSScriptAnalyzer -Scope CurrentUser`. |
| AST + PSSA | Для каждого `.ps1` файла: парсинг AST (`Parser.ParseFile`) и запуск `Invoke-ScriptAnalyzer` с RuleSet. Ошибки в любом файле ломают job. |

Особенности:
* Исключены файлы внутри временной папки `pwsh7`.
* Журналирование выводит отдельный блок на файл и общий итог «SUCCESS/FAIL».

### Workflow `UTF-8 Encoding Check v6`
Файл: [`.github/workflows/check_encoding.yml`](.github/workflows/check_encoding.yml)

Проверяет кодировки без запуска внешних утилит:
* `.ps1`, `.psm1`, `.psd1` — обязаны иметь **UTF-8 BOM** (функция `Test-Utf8Bom`).
* `.md`, `.txt`, `.yml`, `.yaml` — должны быть **UTF-8 без BOM** (функция `Test-Utf8WithoutBom`).
* Результат выводится в виде таблицы; при обнаружении ошибки job завершается со статусом `exit 1`.

### RuleSet для PSScriptAnalyzer
Файл: [`.github/PSScriptAnalyzerRuleSet.psd1`](.github/PSScriptAnalyzerRuleSet.psd1)

| Правило | Уровень | Назначение |
|---------|---------|------------|
| PSAvoidAssignmentToAutomaticVariable | Error | Запрещает присваивать значения встроенным переменным PowerShell. |
| PSMisleadingBacktick | Error | Ловит случайные обратные апострофы в конце строк. |
| PSMisleadingPipeRedirect | Warning | Предупреждает о путанице между конвейером и перенаправлением. |
| PSPossibleIncorrectComparisonWithNull | Error | Проверяет корректные сравнения с `$null`. |
| PSPossibleIncorrectUsageOfAssignmentOperator | Error | Ловит перепутанные `=` и `-eq`. |
| PSReturnCorrectTypes | Warning | Требует одинаковые типы возвращаемых значений. |
| PSAvoidUsingEmptyCatchBlock | Error | Запрещает пустые `catch`. |
| PSAvoidUsingInvokeExpression | Error | Блокирует `Invoke-Expression`. |
| PSAvoidUsingPlainTextForPassword | Warning | Предупреждает о хранении паролей в явном виде. |
| PSAvoidUsingConvertToSecureStringWithPlainText | Warning | Контролирует конвертацию паролей без шифрования. |
| PSUseApprovedVerbs | Warning | Требует стандартных глаголов PowerShell. |
| PSUseShouldProcessForStateChangingFunctions | Warning | Требует `SupportsShouldProcess` и `WhatIf/Confirm`. |

Все правила применяются с `Severity = Error` или `Warning`; workflow трактует любые findings как причину провала job.

## Roadmap по CI/CD и качеству
Актуальное состояние и планы основаны на документе `RoadMap_powershell_VBA_sync.md`.

### Уже сделано
* Настроены ветки `master` и `dev`, локальная конфигурация git.
* Созданы и обновлены workflows `analyze_ps_v7` и `check_encoding` с универсальными триггерами.
* Добавлен AST-анализ и RuleSet уровня «A».
* Внедрены проверки UTF-8 BOM для `.ps1/.psd1` и отсутствие BOM для документации.
* Включена защита ветки `master`, обязательные статусы и запрет прямых push.

### Логичные следующие шаги
1. **Автоматическое удаление веток после merge** — включить опцию в настройках репозитория.
2. **Pester-тесты** — добавить job `Invoke-Pester` для `Import-VBA`/`Export-VBA`.
3. **Автосборка релизов** — при merge в `master` создавать тег и GitHub Release.
4. **YAML lint** — запуск `yamllint` для всех workflow-файлов.
5. **Markdown/JSON lint** — автоматические проверки форматирования документации.
6. **Дополнительный PSScriptAnalyzer job (Warnings)** — второй job, публикующий рекомендации без блокировки PR.
7. **Проверка whitespace** — контроль табуляции, лишних пробелов и конца файла.

## Требования к окружению
* Windows с установленным Excel 2016/365 (COM-доступ к VBA разрешён в Центре управления безопасностью).
* PowerShell 7 (x64/x86). Скрипт автоматически перезапускается в 32-битной версии при подключении к 32-битному Excel.
* Разрешения на выполнение скриптов (`Set-ExecutionPolicy Bypass -Scope Process`).

## Практические советы
* Закрывайте Excel перед импортом форм (`.frm/.frx`), чтобы файлы не блокировались.
* Добавьте папку `VBA` в систему контроля версий — так код VBA участвует в ревью вместе с остальными файлами.
* Файл `SyncVBA.log` можно добавить в `.gitignore`, если журнал не нужен в истории репозитория.
* При ошибках кодировки запускайте workflow локально или в Actions — он точно укажет файл и ожидаемый формат.
