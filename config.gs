/**
 * Конфигурация проекта
 */

/**
 * Получить ID Google Таблицы из свойств скрипта или активной таблицы
 * @return {string} ID таблицы
 */
function getSpreadsheetId() {
  // Сначала пытаемся получить из свойств
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) {
    return id;
  }
  
  // Если не установлен, используем активную таблицу
  try {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSpreadsheet) {
      var activeId = activeSpreadsheet.getId();
      // Сохраняем для будущего использования
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', activeId);
      return activeId;
    }
  } catch (e) {
    Logger.log('Не удалось получить активную таблицу: ' + e.toString());
  }
  
  throw new Error('SPREADSHEET_ID не установлен и не удалось получить активную таблицу');
}

/**
 * Получить имя листа для ежедневных финансовых отчетов
 * @return {string} Имя листа
 */
function getFinanceDailySheetName() {
  return 'Ежедневные фин отчеты';
}

/**
 * Получить имя листа для ленты заказов
 * @return {string} Имя листа
 */
function getOrdersFeedSheetName() {
  return 'Лента заказов';
}

/**
 * Получить имя листа для остатков
 * @return {string} Имя листа
 */
function getStocksSheetName() {
  return 'Остатки';
}

/**
 * Получить Google Таблицу
 * @return {Spreadsheet} Объект таблицы
 */
function getSpreadsheet() {
  // Сначала пытаемся получить активную таблицу
  try {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSpreadsheet) {
      return activeSpreadsheet;
    }
  } catch (e) {
    Logger.log('Не удалось получить активную таблицу, используем ID: ' + e.toString());
  }
  
  // Если активной таблицы нет (например, при запуске по триггеру), используем ID
  return SpreadsheetApp.openById(getSpreadsheetId());
}
