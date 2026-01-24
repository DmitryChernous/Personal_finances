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
- `.clasp.json` — локальная привязка к Script ID (в `.gitignore`)
- `.clasp.json.example` — шаблон

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
- создаст локальный файл `.clasp.json` со `scriptId`

### 3) Указать `rootDir` в `.clasp.json`

`clasp create` обычно создаёт `.clasp.json` только с `scriptId`.
Нужно добавить `rootDir`, чтобы синхронизировать `src/`.

Пример содержимого (можно взять из `.clasp.json.example`):

```json
{
  "scriptId": "PASTE_SCRIPT_ID_HERE",
  "rootDir": "src"
}
```

### 4) Залить код в Apps Script

```powershell
npm run push
```

### 5) Открыть проект в браузере

```powershell
npm run open
```

## Если хочешь привязать к уже существующей Таблице

Нужен ID таблицы (из URL вида `https://docs.google.com/spreadsheets/d/<ID>/edit`).

Команда:

```powershell
npx clasp create --parentId "<SHEET_ID>" --title "Personal finances"
```

Дальше те же шаги: добавить `rootDir` и сделать `npm run push`.

## Быстрая проверка “связь работает”

После `push` открой таблицу → обнови страницу.
В меню появится пункт **Personal finances** → **Setup (создать листы)**.

