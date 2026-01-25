# Архитектурный обзор и анализ проблем

**Дата:** 2025-01-XX  
**Проект:** Personal Finances (Google Sheets + Apps Script)  
**Статус:** Критические проблемы требуют немедленного исправления

---

## Оглавление

1. [Критические проблемы](#критические-проблемы)
2. [Архитектурные проблемы](#архитектурные-проблемы)
3. [Проблемы производительности](#проблемы-производительности)
4. [Проблемы надежности](#проблемы-надежности)
5. [Проблемы безопасности](#проблемы-безопасности)
6. [Рекомендации по улучшению](#рекомендации-по-улучшению)

---

## Критические проблемы

### 1. КРИТИЧЕСКАЯ ОШИБКА: Синтаксическая ошибка в ImportSberbank.js

**Местоположение:** `src/ImportSberbank.js`, строки 71-129

**Проблема:**
```javascript
for (var i = startRow; i < endRow; i++) {
  var line = lines[i].trim();
  
  // Skip empty lines and page breaks
  if (line === '' || ...) {
    continue;
  }
  
  // Stop at footer
  if (line.indexOf('Для проверки подлинности') !== -1 || ...) {
    // ...
    break; // Выход из цикла
  }

  // ПРОБЛЕМА: Парсинг CSV находится ВНЕ цикла из-за неправильной структуры
  // Parse CSV line (handle quoted fields)
  var fields = this._parseCSVLine_(line, delimiter);
  
  // Check if this is a new transaction...
  // ...
}
```

**Описание:** 
Код парсинга CSV (строка 94) находится внутри цикла `for (var i = startRow; i < endRow; i++)`, но из-за неправильной структуры условий он выполняется только для последней строки перед `break`, а не для всех строк секции. Также есть лишняя закрывающая скобка на строке 129, которая может вызывать синтаксические ошибки.

**Проблема:**
```javascript
// Строка 82-91: проверка условий для остановки парсинга
if (line.indexOf('Для проверки подлинности') !== -1 || ...) {
  // ...
  break; // Выход из цикла
}

// Строка 93: Парсинг CSV - НО ЭТО ВНЕ ЦИКЛА!
var fields = this._parseCSVLine_(line, delimiter);
```

**Описание:** После `break` на строке 90 код продолжается на строке 93, но это уже вне цикла `for (var i = startRow; i < endRow; i++)`. Это приводит к тому, что парсинг выполняется только для последней строки перед `break`, а не для всех строк секции.

**Последствия:**
- Парсер обрабатывает только первую транзакцию в каждой секции
- Остальные транзакции игнорируются
- Результат: вместо 532 транзакций обрабатывается только 9 (по одной на страницу)

**Решение:**
```javascript
for (var i = startRow; i < endRow; i++) {
  var line = lines[i].trim();
  
  // Skip empty lines and page breaks
  if (line === '' || 
      line.indexOf('Продолжение на следующей странице') !== -1 ||
      line.indexOf('Страница') !== -1 && line.indexOf('из') !== -1) {
    continue; // Пропустить, но продолжить цикл
  }
  
  // Stop at footer
  if (line.indexOf('Для проверки подлинности') !== -1 ||
      line.indexOf('Действителен') !== -1 ||
      (line.indexOf('Выписка по счёту') !== -1 && line.indexOf('Страница') !== -1)) {
    // End of this section, but continue to next section
    if (currentTransaction) {
      transactions.push(currentTransaction);
      currentTransaction = null;
    }
    break; // Выход из цикла секции
  }
  
  // Парсинг CSV - ВНУТРИ ЦИКЛА
  var fields = this._parseCSVLine_(line, delimiter);
  
  // ... остальная логика парсинга
}
```

---

### 2. Неправильная передача данных между клиентом и сервером

**Местоположение:** `src/ImportUI.html`, строка 360; `src/Import.js`, строка 472

**Проблема:**
```javascript
// ImportUI.html:360
processDataInBatches(parseResult.rawData, formatInfo.importerType, options, parseResult.count);

// Import.js:472
function pfParseFileContent(fileContent, importerType, options) {
  // ...
  return {
    rawData: rawData,  // Большой массив (532 транзакции)
    count: rawData.length,
    errors: errors
  };
}
```

**Описание:** 
- `pfParseFileContent` возвращает полный массив `rawData` через `google.script.run`
- Google Apps Script имеет лимит на размер данных, передаваемых через `google.script.run` (обычно ~50MB, но на практике может быть меньше)
- Для 532 транзакций с многострочными описаниями это может быть проблемой
- Кроме того, весь массив передается обратно в клиент, а затем снова на сервер по частям

**Последствия:**
- Возможные ошибки "Data too large" или таймауты
- Неэффективное использование памяти
- Медленная передача данных

**Решение:**
```javascript
// Вариант 1: Хранить rawData на сервере, передавать только ID сессии
function pfParseFileContent(fileContent, importerType, options) {
  // ...
  var sessionId = Utilities.getUuid();
  // Сохранить rawData в ScriptProperties или CacheService
  CacheService.getScriptCache().put(sessionId, JSON.stringify(rawData), 3600); // 1 час
  
  return {
    sessionId: sessionId,  // Вместо rawData
    count: rawData.length,
    errors: errors
  };
}

// Вариант 2: Обрабатывать на сервере, передавать только прогресс
function pfProcessImportWithSession(sessionId, importerType, options, batchSize, startIndex) {
  var rawDataJson = CacheService.getScriptCache().get(sessionId);
  var rawData = JSON.parse(rawDataJson);
  // Обработать батч
  // ...
}
```

---

### 3. КРИТИЧЕСКАЯ ОШИБКА: Неопределенная переменная в Import.js

**Местоположение:** `src/Import.js`, строка 542

**Проблема:**
```javascript
var cacheKey = 'pf_import_existing_keys';
var cachedKeys = PropertiesService.getScriptProperties().getProperty(cacheKey);

if (cachedKeys) {  // ОК
  // ...
} else {
  existingKeys = pfGetExistingTransactionKeys_();
  PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(existingKeys));
}

// НО на строке 542 используется переменная cachedKeys, которая не была объявлена выше!
```

**Описание:** 
В коде отсутствует строка `var cachedKeys = PropertiesService.getScriptProperties().getProperty(cacheKey);` перед использованием переменной `cachedKeys` на строке 542. Это приводит к ошибке "ReferenceError: cachedKeys is not defined".

**Решение:**
```javascript
var cacheKey = 'pf_import_existing_keys';
var cachedKeys = PropertiesService.getScriptProperties().getProperty(cacheKey); // ДОБАВИТЬ ЭТУ СТРОКУ

if (cachedKeys) {
  try {
    existingKeys = JSON.parse(cachedKeys);
  } catch (e) {
    existingKeys = pfGetExistingTransactionKeys_();
    PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(existingKeys));
  }
} else {
  existingKeys = pfGetExistingTransactionKeys_();
  PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(existingKeys));
}
```

---

### 4. Проблема с кэшированием ключей дедупликации

**Местоположение:** `src/Import.js`, строки 538-593

**Проблема:**
```javascript
// Get existing keys only once (cache it in ScriptProperties for persistence across calls)
var existingKeys = null;
var cacheKey = 'pf_import_existing_keys';
var cachedKeys = PropertiesService.getScriptProperties().getProperty(cacheKey);
if (cachedKeys) {
  try {
    existingKeys = JSON.parse(cachedKeys);
  } catch (e) {
    existingKeys = pfGetExistingTransactionKeys_();
    PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(existingKeys));
  }
} else {
  existingKeys = pfGetExistingTransactionKeys_();
  PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(existingKeys));
}

// ... обработка батча ...

// Update cache
PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(existingKeys));
```

**Описание:**
- `PropertiesService` имеет лимит 9KB на значение
- Для большого количества транзакций ключи дедупликации могут превысить этот лимит
- `PropertiesService.setProperty()` и `getProperty()` - синхронные операции, которые могут быть медленными
- Кэш обновляется после каждого батча, что неэффективно
- Нет очистки кэша после завершения импорта

**Последствия:**
- Ошибки "Value too large" при большом количестве транзакций
- Медленная обработка из-за частых операций с PropertiesService
- Утечка памяти (кэш не очищается)

**Решение:**
```javascript
// Вариант 1: Использовать CacheService (лимит 100KB, быстрее)
var cache = CacheService.getScriptCache();
var cacheKey = 'pf_import_keys_' + Utilities.getUuid(); // Уникальный ключ для сессии

// Вариант 2: Загружать ключи один раз в начале, хранить в памяти функции
// (но это не работает для батчевой обработки через разные вызовы)

// Вариант 3: Использовать временный лист для хранения ключей
function pfGetExistingTransactionKeysBatch_(startRow, endRow) {
  // Загружать только нужный диапазон ключей
}

// Вариант 4: Оптимизировать структуру ключей (использовать Set вместо Object)
var existingKeysSet = new Set(); // Но Set не поддерживается в Apps Script V8 напрямую
// Использовать Map или оптимизированный объект
```

**Рекомендуемое решение:**
```javascript
function pfProcessDataBatch(rawDataJson, importerType, options, batchSize, startIndex) {
  // ...
  
  // Загрузить ключи один раз в начале сессии, хранить в options
  if (!options._existingKeysLoaded) {
    options._existingKeys = pfGetExistingTransactionKeys_();
    options._existingKeysLoaded = true;
  }
  var existingKeys = options._existingKeys;
  
  // Обработать батч
  for (var i = 0; i < rawData.length; i++) {
    // ... обработка ...
    // Обновлять existingKeys в памяти (не сохранять в PropertiesService)
  }
  
  // НЕ сохранять в PropertiesService после каждого батча
  // Сохранить только в конце всей обработки (если нужно)
  
  return {
    transactions: transactions,
    stats: stats,
    processed: processed,
    total: totalCount,
    hasMore: processed < totalCount,
    existingKeys: existingKeys // Вернуть обновленные ключи клиенту
  };
}

// В клиенте:
function processDataInBatches(rawData, importerType, options, totalCount) {
  // ...
  var existingKeys = null; // Хранить на клиенте
  
  function processNextBatch() {
    // ...
    google.script.run
      .withSuccessHandler(function(batchResult) {
        // ...
        existingKeys = batchResult.existingKeys; // Обновить ключи
        // Передать в следующий батч
        batchOptions._existingKeys = existingKeys;
        // ...
      })
      .pfProcessDataBatch(JSON.stringify(batchData), importerType, batchOptions, batchSize, processed);
  }
}
```

---

### 5. Неправильная логика батчевой обработки

**Местоположение:** `src/Import.js`, строки 512-606; `src/ImportUI.html`, строки 380-485

**Проблема:**
```javascript
// ImportUI.html:430
var batchData = rawData.slice(processed, batchEnd);

// Import.js:555
for (var i = 0; i < rawData.length; i++) {
  // Обрабатывает весь rawData, а не только батч
}
```

**Описание:**
- Клиент передает только батч (`batchData`), но серверная функция `pfProcessDataBatch` все равно пытается обработать весь `rawData.length`
- Параметр `startIndex` не используется правильно
- Логика расчета `processed` и `hasMore` может быть неправильной

**Последствия:**
- Неправильный подсчет обработанных транзакций
- Возможные дубликаты или пропуски
- Неправильный прогресс

**Решение:**
```javascript
function pfProcessDataBatch(rawDataJson, importerType, options, batchSize, startIndex) {
  var rawData = JSON.parse(rawDataJson); // rawData - это уже батч, не весь массив
  // startIndex не нужен, так как rawData уже содержит только батч
  
  // Обработать все элементы в батче
  for (var i = 0; i < rawData.length; i++) {
    // ... обработка ...
  }
  
  // Правильно рассчитать processed
  var processed = (options._startIndex || 0) + rawData.length;
  var totalCount = options._totalCount || rawData.length;
  
  return {
    transactions: transactions,
    stats: stats,
    processed: processed,
    total: totalCount,
    hasMore: processed < totalCount
  };
}
```

---

### 6. Отсутствие обработки таймаутов и ошибок

**Местоположение:** `src/ImportUI.html`, строки 437-480

**Проблема:**
```javascript
google.script.run
  .withSuccessHandler(function(batchResult) {
    // ...
    if (batchResult.hasMore) {
      setTimeout(processNextBatch, 50);
    }
  })
  .withFailureHandler(function(error) {
    hideProgress();
    showError('Ошибка...');
  })
  .pfProcessDataBatch(...);
```

**Описание:**
- Нет таймаута для вызова `google.script.run`
- Если серверная функция зависнет, клиент будет ждать бесконечно
- Нет механизма повторных попыток
- Нет логирования ошибок для отладки

**Последствия:**
- UI зависает без обратной связи
- Невозможно диагностировать проблемы
- Плохой пользовательский опыт

**Решение:**
```javascript
function processNextBatch() {
  // ...
  
  var timeoutId = setTimeout(function() {
    hideProgress();
    showError('Таймаут при обработке. Попробуйте уменьшить размер файла или разбить его на части.');
    document.getElementById('previewBtn').disabled = false;
  }, 300000); // 5 минут таймаут
  
  google.script.run
    .withSuccessHandler(function(batchResult) {
      clearTimeout(timeoutId);
      // ... обработка результата ...
    })
    .withFailureHandler(function(error) {
      clearTimeout(timeoutId);
      console.error('Batch processing error:', error);
      hideProgress();
      showError('Ошибка при обработке данных (строка ' + (processed + 1) + '): ' + 
                (error.message || error.toString()) + 
                '\nДетали: ' + JSON.stringify(error));
      document.getElementById('previewBtn').disabled = false;
    })
    .pfProcessDataBatch(...);
}
```

---

## Архитектурные проблемы

### 7. Смешение ответственности в функциях

**Местоположение:** `src/Import.js`, функция `pfProcessDataBatch`

**Проблема:**
Функция `pfProcessDataBatch` выполняет слишком много задач:
- Парсинг JSON
- Выбор импортера
- Загрузка ключей дедупликации
- Нормализация транзакций
- Проверка дубликатов
- Подсчет статистики
- Сохранение кэша

**Решение:**
Разделить на отдельные функции:
```javascript
function pfProcessDataBatch(rawDataJson, importerType, options, batchSize, startIndex) {
  var rawData = JSON.parse(rawDataJson);
  var importer = pfGetImporter_(importerType);
  var existingKeys = pfLoadExistingKeys_(options);
  
  var results = pfNormalizeBatch_(rawData, importer, options);
  var deduplicated = pfDeduplicateBatch_(results.transactions, existingKeys);
  var stats = pfCalculateStats_(deduplicated);
  
  return {
    transactions: deduplicated.transactions,
    stats: stats,
    processed: processed,
    total: totalCount,
    hasMore: hasMore,
    existingKeys: deduplicated.updatedKeys
  };
}
```

---

### 8. Отсутствие валидации входных данных

**Местоположение:** Все функции импорта

**Проблема:**
- Нет проверки типов параметров
- Нет проверки на null/undefined
- Нет проверки размеров данных

**Решение:**
```javascript
function pfProcessDataBatch(rawDataJson, importerType, options, batchSize, startIndex) {
  // Валидация
  if (!rawDataJson || typeof rawDataJson !== 'string') {
    throw new Error('rawDataJson must be a non-empty string');
  }
  if (!importerType || !['sberbank', 'csv'].includes(importerType)) {
    throw new Error('Invalid importerType: ' + importerType);
  }
  if (batchSize && (batchSize < 1 || batchSize > 1000)) {
    throw new Error('batchSize must be between 1 and 1000');
  }
  
  // ... остальной код ...
}
```

---

### 9. Неэффективное использование памяти

**Местоположение:** `src/ImportUI.html`, функция `processDataInBatches`

**Проблема:**
```javascript
var allTransactions = []; // Накапливает все транзакции в памяти клиента
// ...
allTransactions = allTransactions.concat(batchResult.transactions);
// ...
.pfWritePreview(JSON.stringify(allTransactions)); // Передает весь массив на сервер
```

**Описание:**
- Все транзакции накапливаются в памяти браузера
- Для 532 транзакций это может быть проблемой на слабых устройствах
- Финальная передача всего массива на сервер может превысить лимиты

**Решение:**
```javascript
// Вариант 1: Записывать на сервер по частям
function pfAppendPreviewBatch(transactionsJson, isLast) {
  var transactions = JSON.parse(transactionsJson);
  var stagingSheet = pfEnsureImportRawSheet_(SpreadsheetApp.getActiveSpreadsheet());
  // Записать батч в конец листа
  // ...
}

// Вариант 2: Использовать серверную сессию
function pfStartPreviewSession() {
  var sessionId = Utilities.getUuid();
  // Создать пустой staging sheet с sessionId
  return { sessionId: sessionId };
}

function pfAppendToPreview(sessionId, transactionsJson) {
  // Добавить транзакции к существующему preview
}
```

---

## Проблемы производительности

### 10. Множественные обращения к SpreadsheetApp

**Местоположение:** `src/Import.js`, функция `pfGetExistingTransactionKeys_`

**Проблема:**
```javascript
function pfGetExistingTransactionKeys_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet || txSheet.getLastRow() <= 1) return {};
  
  // Одно большое чтение
  var data = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
  // ...
}
```

**Описание:**
- Чтение всех транзакций при каждом импорте
- Для больших таблиц это может быть медленно
- Вызывается при каждом батче (хотя и кэшируется)

**Решение:**
- Кэшировать результат на время сессии импорта
- Использовать инкрементальную загрузку ключей
- Оптимизировать структуру данных для быстрого поиска

---

### 11. Неоптимальная обработка больших файлов

**Местоположение:** Весь процесс импорта

**Проблема:**
- Парсинг всего файла в память сразу
- Нет потоковой обработки
- Нет ограничения на размер файла

**Решение:**
- Добавить проверку размера файла перед обработкой
- Использовать потоковую обработку (если возможно в Apps Script)
- Разбивать большие файлы на части

---

## Проблемы надежности

### 12. Отсутствие транзакционности

**Местоположение:** `src/Import.js`, функция `pfCommitImport_`

**Проблема:**
Если импорт прервется на середине, часть транзакций будет добавлена, а часть - нет.

**Решение:**
- Использовать транзакционный подход (если возможно)
- Добавить механизм отката
- Использовать флаги для отслеживания состояния импорта

---

### 13. Нет обработки конкурентных импортов

**Проблема:**
Если пользователь запустит два импорта одновременно, могут возникнуть конфликты.

**Решение:**
- Добавить блокировку на время импорта
- Проверять наличие активного импорта перед началом нового

---

## Проблемы безопасности

### 14. Отсутствие валидации данных от пользователя

**Проблема:**
- Нет проверки на вредоносный код в CSV
- Нет ограничения на размер файла
- Нет санитизации входных данных

**Решение:**
- Добавить валидацию размера файла
- Санитизировать все строковые данные
- Ограничить размер полей

---

## Рекомендации по улучшению

### Приоритет 1 (Критично - исправить немедленно)

1. **Исправить синтаксическую ошибку в ImportSberbank.js** (проблема #1)
2. **Исправить неопределенную переменную cachedKeys** (проблема #3)
3. **Исправить логику батчевой обработки** (проблема #5)
4. **Оптимизировать передачу данных** (проблема #2)
5. **Исправить кэширование ключей** (проблема #4)

### Приоритет 2 (Важно - исправить в ближайшее время)

6. **Добавить обработку таймаутов** (проблема #6)
7. **Разделить ответственность функций** (проблема #7)
8. **Добавить валидацию входных данных** (проблема #8)
9. **Оптимизировать использование памяти** (проблема #9)

### Приоритет 3 (Желательно - улучшить позже)

10. **Оптимизировать обращения к SpreadsheetApp** (проблема #10)
11. **Улучшить обработку больших файлов** (проблема #11)
12. **Добавить транзакционность** (проблема #12)
13. **Обработать конкурентные импорты** (проблема #13)
14. **Улучшить безопасность** (проблема #14)

---

## План действий

### Шаг 1: Исправить критические ошибки
1. Исправить синтаксическую ошибку в `ImportSberbank.js` (проблема #1)
2. Исправить неопределенную переменную `cachedKeys` в `Import.js` (проблема #3)
3. Исправить логику батчевой обработки (проблема #5)
4. Протестировать на реальном файле

### Шаг 2: Оптимизировать передачу данных
1. Реализовать сессионное хранение `rawData` на сервере
2. Изменить клиент для работы с сессиями
3. Протестировать производительность

### Шаг 3: Улучшить надежность
1. Добавить таймауты
2. Добавить обработку ошибок
3. Добавить логирование

### Шаг 4: Рефакторинг
1. Разделить большие функции
2. Добавить валидацию
3. Улучшить документацию

---

## Заключение

Текущая архитектура импорта имеет несколько критических проблем, которые препятствуют корректной работе с большими файлами. Наиболее критичными являются синтаксическая ошибка в парсере Сбербанка и неправильная логика батчевой обработки. После исправления этих проблем система должна работать корректно.

Рекомендуется начать с исправления проблем Приоритета 1, затем перейти к проблемам Приоритета 2, и только после этого заниматься оптимизацией (Приоритет 3).
