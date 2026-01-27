/**
 * Утилиты для работы с Google Sheets
 */

/**
 * Получить лист по имени, создать если не существует
 * @param {string} sheetName - Имя листа
 * @return {Sheet} Объект листа
 */
function getOrCreateSheet(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log('Создан новый лист: ' + sheetName);
  }
  
  return sheet;
}

/**
 * Проверить, пуст ли лист
 * @param {Sheet} sheet - Лист
 * @return {boolean} true если лист пуст
 */
function isSheetEmpty(sheet) {
  var lastRow = sheet.getLastRow();
  return lastRow === 0;
}

/**
 * Получить все данные из листа
 * @param {Sheet} sheet - Лист
 * @return {Array<Array>} Массив строк
 */
function getSheetData(sheet) {
  if (isSheetEmpty(sheet)) {
    return [];
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) {
    return [];
  }
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

/**
 * Добавить данные в конец листа
 * @param {Sheet} sheet - Лист
 * @param {Array<Array>} data - Данные для добавления
 */
function appendDataToSheet(sheet, data) {
  if (!data || data.length === 0) {
    Logger.log('Нет данных для записи');
    return;
  }
  
  var lastRow = sheet.getLastRow();
  var startRow = lastRow === 0 ? 1 : lastRow + 1;
  
  sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
  Logger.log('Добавлено строк: ' + data.length);
}

/**
 * Установить заголовки листа
 * @param {Sheet} sheet - Лист
 * @param {Array<string>} headers - Массив заголовков
 */
function setSheetHeaders(sheet, headers) {
  if (isSheetEmpty(sheet)) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    // Форматирование заголовков
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e0e0e0');
  }
}

/**
 * Проверить, существует ли дата в листе
 * @param {Sheet} sheet - Лист
 * @param {string} dateStr - Дата в формате YYYY-MM-DD
 * @param {number} dateColumnIndex - Индекс столбца с датой (начиная с 1)
 * @return {boolean} true если дата уже существует
 */
function dateExistsInSheet(sheet, dateStr, dateColumnIndex) {
  if (isSheetEmpty(sheet)) {
    return false;
  }
  
  var tz = sheet.getParent().getSpreadsheetTimeZone();
  var targetDateStr = normalizeDateToIso(dateStr, tz);
  if (!targetDateStr) {
    return false;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { // Только заголовок
    return false;
  }
  
  var dateColumn = sheet.getRange(2, dateColumnIndex, lastRow - 1, 1).getValues();
  
  for (var i = 0; i < dateColumn.length; i++) {
    var cellValue = dateColumn[i][0];
    var cellDateStr = normalizeDateToIso(cellValue, tz);
    
    if (cellDateStr && cellDateStr === targetDateStr) {
      return true;
    }
  }
  
  return false;
}

/**
 * Нормализовать дату в строку YYYY-MM-DD без изменения формата в листе
 * @param {*} value - Дата (string/Date/number)
 * @param {string} tz - Часовой пояс
 * @return {string} Дата в формате YYYY-MM-DD или пустая строка
 */
function normalizeDateToIso(value, tz) {
  if (!value && value !== 0) {
    return '';
  }
  
  var timezone = tz || Session.getScriptTimeZone();
  
  if (value instanceof Date) {
    return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
  }
  
  var valueType = typeof value;
  
  if (valueType === 'number') {
    var dateObj = new Date((value - 25569) * 86400 * 1000);
    if (!isNaN(dateObj.getTime())) {
      return Utilities.formatDate(dateObj, timezone, 'yyyy-MM-dd');
    }
    return '';
  }
  
  if (valueType !== 'string') {
    return '';
  }
  
  var cleanValue = value.replace(/^'/, '').trim();
  if (!cleanValue) {
    return '';
  }
  
  var datePart = cleanValue.split('T')[0].split(' ')[0];
  
  // Формат M/D/YYYY или MM/DD/YYYY
  var slashParts = datePart.split('/');
  if (slashParts.length === 3) {
    var m = parseInt(slashParts[0], 10);
    var d = parseInt(slashParts[1], 10);
    var y = parseInt(slashParts[2], 10);
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31 && y >= 1000) {
      var mm = String(m).padStart(2, '0');
      var dd = String(d).padStart(2, '0');
      return y + '-' + mm + '-' + dd;
    }
  }
  
  // Формат YYYY-MM-DD
  if (datePart.indexOf('-') !== -1) {
    var parsedIso = parseDate(datePart);
    if (parsedIso && !isNaN(parsedIso.getTime())) {
      var isoParts = datePart.split('-');
      if (isoParts.length === 3) {
        var isoYear = isoParts[0];
        var isoMonth = String(parseInt(isoParts[1], 10)).padStart(2, '0');
        var isoDay = String(parseInt(isoParts[2], 10)).padStart(2, '0');
        return isoYear + '-' + isoMonth + '-' + isoDay;
      }
      return Utilities.formatDate(parsedIso, timezone, 'yyyy-MM-dd');
    }
  }
  
  // Формат DD.MM.YYYY
  var dotParts = datePart.split('.');
  if (dotParts.length === 3) {
    var day = parseInt(dotParts[0], 10);
    var month = parseInt(dotParts[1], 10);
    var year = parseInt(dotParts[2], 10);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 1000) {
      var monthStr = String(month).padStart(2, '0');
      var dayStr = String(day).padStart(2, '0');
      return year + '-' + monthStr + '-' + dayStr;
    }
  }
  
  return '';
}

/**
 * Удалить строки с указанной датой из листа
 * @param {Sheet} sheet - Лист
 * @param {string} dateStr - Дата в формате YYYY-MM-DD
 * @param {number} dateColumnIndex - Индекс столбца с датой (начиная с 1)
 * @param {number} headerRowCount - Количество строк заголовков (обычно 1 или 2)
 * @return {number} Количество удаленных строк
 */
function deleteRowsByDate(sheet, dateStr, dateColumnIndex, headerRowCount) {
  headerRowCount = headerRowCount || 1; // По умолчанию 1 строка заголовков
  
  if (isSheetEmpty(sheet)) {
    Logger.log('deleteRowsByDate: Лист пуст');
    return 0;
  }
  
  var ss = sheet.getParent();
  var tz = ss.getSpreadsheetTimeZone();
  var targetDateStr = normalizeDateToIso(dateStr, tz);
  if (!targetDateStr) {
    Logger.log('deleteRowsByDate: Неверный формат даты: ' + dateStr);
    return 0;
  }
  var lastRow = sheet.getLastRow();
  Logger.log('deleteRowsByDate: lastRow=' + lastRow + ', headerRowCount=' + headerRowCount + ', dateStr=' + dateStr + ', targetDateStr=' + targetDateStr + ', tz=' + tz);
  
  if (lastRow <= headerRowCount) {
    Logger.log('deleteRowsByDate: Нет данных для проверки (только заголовки)');
    return 0;
  }
  
  // Получаем данные столбца с датой (начиная со строки после заголовков)
  var dataStartRow = headerRowCount + 1;
  var numRows = lastRow - headerRowCount;
  var dateColumn = sheet.getRange(dataStartRow, dateColumnIndex, numRows, 1).getValues();
  
  Logger.log('deleteRowsByDate: Проверяю ' + dateColumn.length + ' строк, начиная со строки ' + dataStartRow);
  
  // Находим индексы строк для удаления (в обратном порядке, чтобы не сбить нумерацию)
  var rowsToDelete = [];
  for (var i = dateColumn.length - 1; i >= 0; i--) {
    var cellValue = dateColumn[i][0];
    var rowNum = dataStartRow + i;
    
    // Пропускаем пустые значения
    if (!cellValue || cellValue === '') {
      continue;
    }
    
    // Логируем первые 3 строки для отладки
    if (i < 3) {
      Logger.log('deleteRowsByDate: Строка ' + rowNum + ', значение="' + cellValue + '", тип=' + (typeof cellValue) + ', instanceof Date=' + (cellValue instanceof Date));
    }
    
    var cellDateStr = normalizeDateToIso(cellValue, tz);
    
    // Сравниваем даты
    if (i < 3) {
      Logger.log('deleteRowsByDate: Строка ' + rowNum + ', сравнение: cellDateStr="' + cellDateStr + '" === targetDateStr="' + targetDateStr + '" = ' + (cellDateStr === targetDateStr));
    }
    if (cellDateStr && cellDateStr === targetDateStr) {
      rowsToDelete.push(rowNum);
      Logger.log('deleteRowsByDate: Строка ' + rowNum + ' помечена на удаление');
    }
  }
  
  // Удаляем строки (уже в обратном порядке, чтобы не сбить нумерацию)
  if (rowsToDelete.length > 0) {
    Logger.log('deleteRowsByDate: Найдено строк для удаления: ' + rowsToDelete.length);
    
    for (var j = 0; j < rowsToDelete.length; j++) {
      var rowToDelete = rowsToDelete[j];
      try {
        sheet.deleteRow(rowToDelete);
      } catch (e) {
        Logger.log('deleteRowsByDate: ОШИБКА при удалении строки ' + rowToDelete + ': ' + e.toString());
      }
    }
    
    Logger.log('deleteRowsByDate: Успешно удалено строк за дату ' + dateStr + ': ' + rowsToDelete.length);
  } else {
    Logger.log('deleteRowsByDate: Строк с датой ' + dateStr + ' не найдено');
  }
  
  return rowsToDelete.length;
}

/**
 * Полностью очистить лист и записать новые данные
 * @param {Sheet} sheet - Лист
 * @param {Array<string>} headers - Заголовки столбцов
 * @param {Array<Array>} data - Данные для записи
 */
function clearAndWriteSheet(sheet, headers, data) {
  Logger.log('clearAndWriteSheet: заголовков=' + (headers ? headers.length : 0) + ', строк данных=' + (data ? data.length : 0));
  
  // Очищаем весь лист
  sheet.clear();
  Logger.log('Лист очищен');
  
  // Записываем заголовки
  if (headers && headers.length > 0) {
    try {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      Logger.log('Заголовки записаны: ' + headers.length + ' столбцов');
      
      // Форматирование заголовков
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e0e0e0');
    } catch (e) {
      Logger.log('Ошибка записи заголовков: ' + e.toString());
      throw e;
    }
  }
  
  // Записываем данные
  if (data && data.length > 0) {
    try {
      var startRow = headers && headers.length > 0 ? 2 : 1;
      var numCols = data[0] ? data[0].length : 0;
      Logger.log('Запись данных: строка начала=' + startRow + ', строк=' + data.length + ', столбцов=' + numCols);
      
      if (numCols > 0) {
        sheet.getRange(startRow, 1, data.length, numCols).setValues(data);
        Logger.log('Данные записаны успешно: ' + data.length + ' строк');
      } else {
        Logger.log('ОШИБКА: Первая строка данных пустая!');
      }
    } catch (e) {
      Logger.log('Ошибка записи данных: ' + e.toString());
      Logger.log('Стек ошибки: ' + e.stack);
      throw e;
    }
  } else {
    Logger.log('Нет данных для записи');
  }
}
