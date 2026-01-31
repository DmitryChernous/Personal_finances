# Personal_finances

Проект инструмента для учета личных финансов на базе **Google Sheets** + **Google Apps Script**.

## Цель этого репозитория

- Хранить исходники Apps Script в GitHub.
- Синхронизировать код с Google Apps Script проектом через **clasp**.
- Быстро “привязать” проект к Google Sheets (контейнерный проект).

## Требования

- Windows
- Git (у тебя уже есть)
- Node.js (установлен через `winget`)

## Структура

- `src/` — исходники Apps Script
  - `appsscript.json` — манифест
  - `Code.js` — меню/точка входа
  - `Sheets.js` — утилиты для работы с листами
- Обзор структуры и зависимостей кода для разработки: [docs/ARCHITECTURE_AND_DEPS.md](docs/ARCHITECTURE_AND_DEPS.md)
- `.clasp.json` — локальная привязка к Script ID (в `.gitignore`)
- `.clasp.json.example` — шаблон

## Важно: локаль и формулы (русскоязычные Google Sheets)

Проект рассчитан на русскоязычные Google Sheets, поэтому **синтаксис формул и форматы** зависят от локали (часто разделитель аргументов `;`, десятичный `,`).

Подробно зафиксировано в `SHEETS_LOCALE.md`.

## Импорт транзакций

**Надёжный способ:** выгрузка выписки в **CSV или Excel** из личного кабинета банка (например, Сбербанк Онлайн) и импорт через меню приложения. PDF-выписки поддерживаются для нескольких банков, но распознавание может ошибаться. Подробнее и альтернативы — в [docs/IMPORT_ALTERNATIVES.md](docs/IMPORT_ALTERNATIVES.md).

## Подключение к Google Sheets (через clasp)

### 0) Включить Apps Script API (один раз)

`clasp` использует Apps Script API. Если она выключена, команды `create/push/pull` будут падать.

- Открой: `https://script.google.com/home/usersettings`
- Включи переключатель **Google Apps Script API**
- Подожди 1–3 минуты (иногда нужно время на “прокидывание”)

### 1) Войти в Google (один раз)

Из корня репозитория:

```powershell
npm run login
```

Откроется браузер/ссылка для авторизации Google аккаунта.

### 2) Создать новую Google Таблицу + связанный Apps Script проект

В корне репозитория:

```powershell
npm run create:sheets
```

Эта команда:
- создаст новую таблицу в Google Drive
- создаст контейнерный Apps Script проект
- создаст локальный файл `.clasp.json` со `scriptId` и `rootDir: "src"`

### 3) Залить код в Apps Script

```powershell
npm run push
```

Если `push` пишет `Skipping push.`, сделай принудительный пуш манифеста:

```powershell
npm run push:force
```

### 4) Открыть проект в браузере

```powershell
npm run open
```

## Если хочешь привязать к уже существующей Таблице

Нужен ID таблицы (из URL вида `https://docs.google.com/spreadsheets/d/<ID>/edit`).

Команда:

```powershell
npx clasp create --parentId "<SHEET_ID>" --title "Personal finances" --rootDir src
```

Дальше те же шаги: сделать `npm run push` (или `npm run push:force`, если нужно).

## Быстрая проверка “связь работает”

После `push` открой таблицу → обнови страницу.
В меню появится пункт **Personal finances** → **Setup (создать листы)**.

## Запуск тестов

В таблице: **Personal finances → Запустить тесты** (или **Run tests** при английском языке). Результаты юнит-тестов (константы, дедупликация, RawSheets-парсинг, валидация и др.) отобразятся в диалоге. Ручной чеклист сценариев — в [docs/TESTING_CHECKLIST.md](docs/TESTING_CHECKLIST.md).

## Troubleshooting

### ECONNRESET / Connection was reset при `push/pull`

Иногда это бывает из-за IPv6. Быстрый фикс — заставить Node предпочитать IPv4.

В **PowerShell**:

```powershell
$env:NODE_OPTIONS="--dns-result-order=ipv4first"
npm run push
```

В **cmd**:

```bat
set NODE_OPTIONS=--dns-result-order=ipv4first
npm run push
```

