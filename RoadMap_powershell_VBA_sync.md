Готово — вот **краткая, технически точная, без воды** маршрутная карта твоего пути по настройке CI/CD GitHub Actions для PowerShell-проекта.
Даю в **markdown**, полностью готово для копирования в проект.

---

# 📘 Маршрутная карта (Roadmap) по настройке CI/CD GitHub Actions

### *Проект: Sync-VBA PowerShell CI | PSScriptAnalyzer | UTF-8 BOM | Branch Protection*

---

## 1. Подготовка репозитория

* Создан репозиторий `powershell_VBA_sync`
* Приведены в порядок ветки:

  * `master` — стабильная
  * `dev` — рабочая
* Настроено локальное окружение:

  ```
  git init
  git remote add origin …
  git config user.name / user.email
  ```

---

## 2. Структура CI

Создана структура:

```
.github/
   workflows/
       analyze_ps_v5.yml
       check_encoding.yml
   PSScriptAnalyzerRuleSet.psd1
```

---

## 3. Первая версия workflow (v4)

* Создан YAML для PSScriptAnalyzer
* Были ошибки:

  * старый trigger `push: dev`
  * не работало на master
  * не работала проверка синтаксиса
  * использован неправильный параметр `-CustomRulePath`

---

## 4. Исправление triggers

Workflow обновлён:

```
on:
  push:
    branches: [ dev, master ]
  pull_request:
    branches: [ dev, master ]
```

Теперь CI запускается везде, где нужно.

---

## 5. Исправление ошибки парсинга PowerShell

Добавлен анализ синтаксиса:

```
[System.Management.Automation.Language.Parser]::ParseFile()
```

Теперь любые мусорные строки (`dfsd sd e3`) гарантированно ломают CI.

---

## 6. Создание RuleSet (набор А)

Создан файл:

```
PSScriptAnalyzerRuleSet.psd1
```

Формат:

* Кодировка: **UTF-8 BOM**
* Концы строк: **CRLF**

В RuleSet включены только важные правила:

* PSAvoidAssignmentToAutomaticVariable
* PSMisleadingBacktick
* PSMisleadingPipeRedirect
* PSPossibleIncorrectComparisonWithNull
* PSPossibleIncorrectUsageOfAssignmentOperator
* PSReturnCorrectTypes
* PSAvoidUsingEmptyCatchBlock
* PSAvoidUsingInvokeExpression
* PSAvoidUsingPlainTextForPassword
* PSAvoidUsingConvertToSecureStringWithPlainText
* PSUseApprovedVerbs
* PSUseShouldProcessForStateChangingFunctions

---

## 7. Создание рабочего workflow (v5)

В analyze_ps_v5:

* убран `-CustomRulePath`
* добавлен `-Settings PSScriptAnalyzerRuleSet.psd1`
* добавлен синтаксический анализ
* добавлена обработка ошибок на уровне файла
* единый лог работы

---

## 8. Удаление устаревших workflow

Удалён:

```
.github/workflows/analyze_ps_v4.yml
```

чтобы не было дублирующихся checks.

---

## 9. Настройка проверки UTF-8 BOM

Создан workflow:

```
check_encoding.yml
```

Он проверяет:

* `.ps1` → UTF-8 BOM
* `.txt`, `.md`, `.yml` → UTF-8 без BOM
* `.psd1` → UTF-8 BOM

---

## 10. Настройка Branch Protection

Включено:

* Require PR before merging
* Require status checks to pass
* Required checks:

  * **PSScriptAnalyzer Check v5 / pssa**
  * **UTF-8 BOM Encoding Check / encoding_check**
* Disallow direct push to master
* Linear history

Теперь master защищён, ошибки не проходят.

---

## 11. Проверка pipeline

Тестовые ошибки (вставка мусора) корректно ломают pipeline → PR запрещён.
После исправления — merge проходит.

---

# ✔ Итоговое состояние CI

| Компонент                       | Статус      |
| ------------------------------- | ----------- |
| PSScriptAnalyzer (RuleSet A)    | ✔ Работает  |
| Ловит синтаксические ошибки     | ✔           |
| Ловит критические анти-паттерны | ✔           |
| UTF-8 BOM check                 | ✔           |
| Triggers для dev + master       | ✔           |
| Branch Protection               | ✔ настроено |
| Старые workflow удалены         | ✔           |
| Ошибки не попадают в master     | ✔           |
| PR pipeline стабилен            | ✔           |

---

# 🔧 Доработки, которые логично добавить дальше

## 1. Автоматическое удаление веток после merge

Включается в GitHub Settings.

## 2. Pester-тесты PowerShell

Добавить:

```
Invoke-Pester
```

в CI → можно тестировать Import-VBA / Export-VBA функцию.

## 3. Workflow для автоматического создания релизов

При merge в master → tag + GitHub Release + архив.

## 4. YAML lint

Добавить проверку валидности YAML файлов:

```
yamllint
```

## 5. Проверка JSON/Markdown линтером

Проблемы форматирования будут ловиться автоматически.

## 6. Анализ кода PowerShell по Warning-правилам (не Error)

Можно добавить второй job, чтобы рекомендации показывались, но PR не блокировался.

## 7. Проверка whitespace/indentation

Если требуется строгий стиль.

---

# 📦 Готовый раздел для проекта ChatGPT

Да — **YAML и RuleSet нужно добавлять в проект ChatGPT**, чтобы:

* модель понимала структуру твоего CI
* могла улучшать workflows
* могла генерировать корректные изменения
* не ошибалась в путях и именах файлов

Добавлять:

```
.github/workflows/analyze_ps_v5.yml
.github/workflows/check_encoding.yml
.github/PSScriptAnalyzerRuleSet.psd1
```

---

Готов продолжить и собрать полноценную документацию CI/CD или добавить релизы, тесты и автогенерацию changelog.
