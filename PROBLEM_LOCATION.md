# 🔴 МЕСТО ПРОБЛЕМЫ - ЗАТЫК В КОДЕ

## Проблемная функция: `pfGetExistingTransactionKeys_()`

**Файл:** `src/Import.js`  
**Строки:** 255-302

### Код, который зависает:

```javascript
function pfGetExistingTransactionKeys_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet || txSheet.getLastRow() <= 1) return {};
  
  var keys = {};
  var sourceCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Source');
  var sourceIdCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'SourceId');
  var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
  var accountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Account');
  var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  var typeCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
  
  if (!sourceCol || !dateCol || !accountCol || !amountCol || !typeCol) return {};
  
  // ⚠️ ПРОБЛЕМА ЗДЕСЬ - строка 270:
  // Читает ВСЕ транзакции из таблицы за один раз!
  var data = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
  
  // ⚠️ ПРОБЛЕМА ЗДЕСЬ - строка 272-299:
  // Обрабатывает каждую транзакцию и вычисляет MD5 хеш для каждой
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var source = row[sourceCol - 1];
    var sourceId = sourceIdCol ? row[sourceIdCol - 1] : null;
    
    if (sourceId) {
      keys[source + ':' + sourceId] = true;
    } else {
      // ⚠️ ОЧЕНЬ МЕДЛЕННО - вычисление MD5 для каждой транзакции
      var date = row[dateCol - 1];
      var account = row[accountCol - 1];
      var amount = row[amountCol - 1];
      var type = row[typeCol - 1];
      
      var keyFields = [
        date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        account || '',
        String(amount || ''),
        type || ''
      ].join('|');
      
      // ⚠️ МЕДЛЕННАЯ ОПЕРАЦИЯ - MD5 хеш для каждой транзакции
      var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, keyFields).map(function(b) {
        return ('0' + (b & 0xFF).toString(16)).slice(-2);
      }).join('');
      
      keys[(source || 'unknown') + ':' + hash] = true;
    }
  }
  
  return keys;
}
```

## Где вызывается эта функция:

**Файл:** `src/Import.js`  
**Строка:** 593

```javascript
function pfProcessDataBatch(rawDataJson, importerType, options, batchSize, startIndex) {
  // ...
  
  // ⚠️ ПРОБЛЕМА ЗДЕСЬ - строка 586-594:
  // При первом батче options._existingKeys не установлен,
  // поэтому вызывается pfGetExistingTransactionKeys_()
  // которая читает ВСЕ транзакции и зависает!
  var existingKeys = null;
  if (options._existingKeys && typeof options._existingKeys === 'object') {
    existingKeys = options._existingKeys;
  } else {
    // ⚠️ ВОТ ЗДЕСЬ ЗАВИСАНИЕ!
    existingKeys = pfGetExistingTransactionKeys_(); // ← ЗАТЫК!
  }
  
  // ...
}
```

## Почему зависает:

1. **При первом вызове** `pfProcessDataBatch` параметр `options._existingKeys` не установлен
2. Вызывается `pfGetExistingTransactionKeys_()`
3. Эта функция читает **ВСЕ транзакции** из таблицы (строка 270)
4. Для каждой транзакции без `sourceId` вычисляется **MD5 хеш** (строка 293)
5. Если в таблице много транзакций (например, 1000+), это занимает **очень много времени**
6. Apps Script имеет лимит времени выполнения (6 минут), но функция может зависнуть раньше

## Решение:

1. **Не вызывать `pfGetExistingTransactionKeys_()` при первом батче**
2. **Начинать с пустого объекта ключей** `{}`
3. **Загружать ключи только если действительно нужно** (например, если в таблице уже есть транзакции)
4. **Или оптимизировать функцию** - загружать ключи порциями, использовать кэш

## Быстрое исправление:

Заменить в `src/Import.js` строки 586-594:

```javascript
// БЫЛО (зависает):
var existingKeys = null;
if (options._existingKeys && typeof options._existingKeys === 'object') {
  existingKeys = options._existingKeys;
} else {
  existingKeys = pfGetExistingTransactionKeys_(); // ← ЗАТЫК!
}

// ДОЛЖНО БЫТЬ (не зависает):
var existingKeys = null;
if (options._existingKeys && typeof options._existingKeys === 'object') {
  existingKeys = options._existingKeys;
} else {
  // Начинаем с пустого объекта - ключи будут накапливаться по мере обработки
  existingKeys = {};
}
```

Или еще лучше - загружать ключи только один раз в начале, перед первым батчем, и передавать через options.
