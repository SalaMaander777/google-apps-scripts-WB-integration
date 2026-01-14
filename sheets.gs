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
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { // Только заголовок
    return false;
  }
  
  var dateColumn = sheet.getRange(2, dateColumnIndex, lastRow - 1, 1).getValues();
  
  for (var i = 0; i < dateColumn.length; i++) {
    var cellValue = dateColumn[i][0];
    var cellDateStr = '';
    
    if (cellValue instanceof Date) {
      // Если значение - объект Date
      cellDateStr = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof cellValue === 'string') {
      // Если значение - строка, пытаемся распарсить
      var parsed = parseDate(cellValue);
      if (parsed) {
        cellDateStr = Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        // Если не удалось распарсить, сравниваем как строку
        cellDateStr = cellValue.split('T')[0]; // Убираем время если есть
      }
    }
    
    if (cellDateStr === dateStr) {
      return true;
    }
  }
  
  return false;
}

/**
 * Полностью очистить лист и записать новые данные
 * @param {Sheet} sheet - Лист
 * @param {Array<string>} headers - Заголовки столбцов
 * @param {Array<Array>} data - Данные для записи
 */
function clearAndWriteSheet(sheet, headers, data) {
  // Очищаем весь лист
  sheet.clear();
  
  // Записываем заголовки
  if (headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    // Форматирование заголовков
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e0e0e0');
  }
  
  // Записываем данные
  if (data && data.length > 0) {
    var startRow = headers && headers.length > 0 ? 2 : 1;
    sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
    Logger.log('Записано строк: ' + data.length);
  }
}
