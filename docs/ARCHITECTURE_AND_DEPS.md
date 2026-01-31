# Структура и зависимости кода

Обзор основных файлов, глобальных объектов и порядка загрузки для разработки. См. также детальные спецификации в конце документа.

## Основные файлы и роль

| Файл | Роль |
|------|------|
| **Code.js** | Точка входа: меню (onOpen), обработчики пунктов меню, вызовы Setup, синхронизации raw, импорта, тестов и т.д. |
| **Schema.js** | Схемы листов: PF_TRANSACTIONS_SCHEMA, PF_ACCOUNTS_SCHEMA, PF_CATEGORIES_SCHEMA и др. |
| **Constants.js** | Константы: PF_TRANSACTION_STATUS, PF_TRANSACTION_TYPE, PF_IMPORT_SOURCE, PF_DEFAULT_CURRENCY, PF_SUPPORTED_CURRENCIES, PF_IMPORT_BATCH_SIZE и т.д. |
| **I18n.js** | Локализация: PF_SHEET_KEYS, PF_I18N (ru/en), pfT_(), pfGetLanguage_(), смена языка. |
| **Sheets.js** | Утилиты листов: pfFindSheetByKey_(), поиск по имени, работа с заголовками. |
| **Setup.js** | Создание и инициализация листов: PF_NAMED_RANGES, PF_SETUP_KEYS, pfColumnIndex_(), pfSetup(), создание листов Транзакции, Счета, Категории, Настройки, Отчёты, Дашборд и т.д. |
| **Settings.js** | Настройки таблицы: pfGetSetting_(), pfSetSetting_(), чтение/запись с листа «Настройки». |
| **RawSheets.js** | Синхронизация raw-листов: чтение листов с префиксом raw, парсинг даты/времени/суммы, формирование SourceId, дедупликация, запись в «Транзакции». |
| **Import.js** | Импорт транзакций: pfCanonicalDedupeKey_(), pfGenerateDedupeKey_(), предпросмотр, коммит из листа «Импорт (черновик)», запись в «Транзакции». |
| **ImportCsv.js** | Парсинг CSV, маппинг колонок, формирование транзакций для импорта. |
| **ImportPdf.js**, **ImportPdfSberbank.js**, **ImportPdfYandex.js** | Парсинг PDF-выписок банков. |
| **ImportSberbank.js** | Импорт выписок Сбербанка (формат CSV и др.). |
| **Validation.js** | Валидация транзакций, правил категоризации. |
| **ErrorHandler.js** | Обработка ошибок, логирование, pfLogError_, pfLogWarning_. |
| **Dashboard.js** | Построение дашборда: KPI, графики по данным из «Транзакции». |
| **Reports.js** | Построение отчётов: сводка по периодам, топ категорий, остатки по счетам. |
| **Budgets.js** | Бюджеты: расчёт факта, остатка, статусов. |
| **RecurringTransactions.js** | Регулярные платежи: создание транзакций по расписанию. |
| **CategoryRules.js** | Правила категоризации: применение правил по паттернам к транзакциям. |
| **Search.js** | Поиск транзакций, диалог поиска (SearchUI.html). |
| **QuickEntry.js** | Быстрый ввод транзакции (QuickEntryUI.html). |
| **Export.js** | Экспорт в CSV/JSON. |
| **Archive.js** | Архивация старых транзакций. |
| **Help.js** | Содержимое листа «Инструкция» (справка). |
| **Template.js** | Создание шаблона таблицы для нового пользователя. |
| **DateUtils.js** | Работа с датами: форматирование, парсинг. |
| **DemoData.js**, **TestData.js** | Тестовые и демо-данные. |
| **Tests.js** | Юнит-тесты: pfRunAllTests(), тесты констант, дедупликации, валидации, RawSheets и др. |

HTML-файлы (ImportUI.html, QuickEntryUI.html, SearchUI.html) — клиентский UI для импорта, быстрого ввода и поиска.

## Глобальные объекты и источник

| Объект / функция | Файл |
|------------------|------|
| PF_SHEET_KEYS | I18n.js |
| PF_SETUP_KEYS | Setup.js |
| PF_NAMED_RANGES | Setup.js |
| PF_TRANSACTIONS_SCHEMA, PF_ACCOUNTS_SCHEMA, PF_CATEGORIES_SCHEMA | Schema.js |
| pfColumnIndex_() | Setup.js |
| pfFindSheetByKey_() | Sheets.js |
| pfGetSetting_() | Settings.js |
| pfT_(), pfGetLanguage_() | I18n.js |
| pfCanonicalDedupeKey_(), pfGenerateDedupeKey_() | Import.js |

## Порядок загрузки

В Google Apps Script порядок файлов в проекте может влиять на инициализацию: скрипты выполняются в одном глобальном контексте, и функции/переменные должны быть объявлены до использования.

Рекомендуемый порядок (при добавлении новых файлов):

1. **Константы и схемы:** Constants.js, Schema.js  
2. **Листы и настройки:** I18n.js, Sheets.js, Setup.js, Settings.js  
3. **Модули данных и UI:** RawSheets.js, Import.js, ImportCsv.js, ImportPdf*.js, Validation.js, ErrorHandler.js  
4. **Аналитика и фичи:** Dashboard.js, Reports.js, Budgets.js, RecurringTransactions.js, CategoryRules.js, Search.js, QuickEntry.js, Export.js, Archive.js, Help.js, Template.js, DateUtils.js, DemoData.js, TestData.js  
5. **Точка входа и тесты:** Code.js, Tests.js  

Точный порядок в проекте задаётся списком файлов в корне Apps Script (при использовании clasp — содержимое папки `src/`; порядок при `clasp push` может совпадать с алфавитным). Важно: Import.js загружается до RawSheets.js (алфавитный порядок), поэтому pfCanonicalDedupeKey_() из Import.js доступна в RawSheets.js.

## Ссылки на детальные документы

- [RAW_SHEETS_ARCHITECTURE.md](RAW_SHEETS_ARCHITECTURE.md) — raw-листы, формат колонок, синхронизация  
- [TRANSACTIONS_SCHEMA.md](TRANSACTIONS_SCHEMA.md) — схема листа «Транзакции», ключ дедупликации  
- [TECHNICAL_SPECIFICATION.md](TECHNICAL_SPECIFICATION.md) — техническая спецификация этапов  
- [PROTOTYPE_REVIEW_AND_IMPLEMENTATION_PLAN.md](PROTOTYPE_REVIEW_AND_IMPLEMENTATION_PLAN.md) — ревью прототипа и план внедрения  
- [REFACTORING_PLAN_FOR_PM.md](REFACTORING_PLAN_FOR_PM.md) — план рефакторинга для PM  
