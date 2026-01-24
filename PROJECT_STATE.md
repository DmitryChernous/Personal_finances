# Project state (snapshot)

Дата: 2026-01-24

Этот файл фиксирует текущее состояние репозитория перед началом разработки основной логики.

## Git

- Ветка: `main` (tracking: `origin/main`)
- Remote: `git@github.com:DmitryChernous/Personal_finances.git`
- HEAD: `a063e727e9a273712840415eb029edd062e8b3dc`
- Рабочая директория: чистая (`git status` без изменений)

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
- `src/Code.js` — меню в Google Sheets + минимальный `pfSetup()`
- `src/Sheets.js` — утилита `getOrCreateSheet_()`
- `README.md` — инструкции по подключению и работе через `clasp`

## Что сейчас реализовано в Apps Script

- `onOpen()` добавляет меню **Personal finances** в Google Sheets.
- Пункт меню **Setup (создать листы)** вызывает `pfSetup()`, который создаёт (или находит) листы:
  - `Transactions`
  - `Categories`
  - `Accounts`

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
- Следующий шаг разработки: определить модель данных (колонки/валидации), и расширить `pfSetup()` до полной инициализации таблицы (шапки, форматы, именованные диапазоны, защиты и т.п.).

