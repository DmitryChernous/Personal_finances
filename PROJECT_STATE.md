# Project state (snapshot)

Дата: 2026-01-24

Этот файл фиксирует текущее состояние репозитория перед началом разработки основной логики.

## Git

- Ветка: `main` (tracking: `origin/main`)
- Remote: `https://github.com/DmitryChernous/Personal_finances.git`
- Рабочая директория: см. `git status`

## Локальная привязка Google Apps Script (clasp)

- В репозитории **есть локальный** файл `.clasp.json` (он **в `.gitignore`** и не должен коммититься).
- `.clasp.json` содержит привязку к конкретному Apps Script проекту и контейнерной таблице:
  - `scriptId`: (локально)
  - `parentId` (Spreadsheet ID): (локально)
  - `rootDir`: `src`
- Шаблон для переноса/восстановления: `.clasp.json.example`

## Node / зависимости

- Node.js: `v24.13.0`
- npm: `11.6.2`
- Основная dev-зависимость: `@google/clasp@3.1.3`

## Структура проекта

- `src/appsscript.json` — манифест Apps Script (V8, Europe/Moscow)
- `src/Code.js` — меню в Google Sheets + `pfSetup()` (локаль + язык + листы)
- `src/Sheets.js` — утилиты для листов + применение локализации
- `src/I18n.js` — словари RU/EN (имена листов, заголовки, меню)
- `src/Settings.js` — хранение настроек в листе Settings/Настройки (пока только язык)
- `src/Schema.js` — схема листов (ключи колонок для локализации)
- `TRANSACTIONS_SCHEMA.md` — спецификация схемы транзакций
- `README.md` — инструкции по подключению и работе через `clasp`

## Что сейчас реализовано в Apps Script

- `onOpen()` добавляет меню **Personal finances** в Google Sheets.
- Пункт меню **Setup (создать листы)** вызывает `pfSetup()`, который создаёт (или находит) листы:
  - `Транзакции` / `Transactions`
  - `Категории` / `Categories`
  - `Счета` / `Accounts`
  - `Настройки` / `Settings`
- В меню есть выбор языка (RU/EN); выбранный язык сохраняется в листе настроек и применяется к именам листов и заголовкам.

## Команды (npm scripts)

- `npm run login` — авторизация clasp
- `npm run create:sheets` — создать Google Sheet + контейнерный Apps Script, rootDir=`src`
- `npm run push` / `npm run push:force` — залить код в Apps Script
- `npm run pull` — скачать код из Apps Script
- `npm run open` — открыть контейнерный проект в браузере
- `npm run logs` — посмотреть логи

## Примечания перед разработкой основной части

- Идентификаторы `scriptId` / `parentId` **не сохраняем в git** (оставляем только локально в `.clasp.json`).
- Таблица **русскоязычная**; синтаксис формул/форматы должны быть совместимы с локалью `ru_RU` (см. `SHEETS_LOCALE.md`).
- Язык интерфейса (RU/EN) должен определять имена листов/заголовки и расширяться через словари в `src/I18n.js`.
- Следующий шаг разработки: расширить `pfSetup()` до полной инициализации таблицы (шапки, форматы, валидации, именованные диапазоны, и т.п.).

