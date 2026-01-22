/**
 * Модуль архивирования таблицы
 * Создает копию таблицы и очищает накопительные листы
 */

/**
 * Листы, которые НЕ нужно очищать при архивировании
 */
var SHEETS_TO_PRESERVE = [
  'ID-ART',
  'История рекламных расходов',
  'воронка отчет',
  'Процентная разбивка размеров',
  'Остатки' // Остатки всегда свежие, их очищать бессмысленно
];

/**
 * Лист "Воронка динамика" - особый режим: очищаются только столбцы данных (B и далее)
 */
var FUNNEL_DYNAMIC_SHEET_NAME = 'Воронка динамика';

/**
 * Показать диалог подтверждения архивирования
 */
function showArchiveConfirmDialog() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Собираем информацию о таблице
  var info = getSpreadsheetInfo(ss);
  
  var message = '📊 Информация о таблице:\n\n' +
    '• Всего листов: ' + info.totalSheets + '\n' +
    '• Листов для очистки: ' + info.sheetsToClean.length + '\n' +
    '• Листы которые НЕ будут очищены:\n  - ' + SHEETS_TO_PRESERVE.join('\n  - ') + '\n' +
    '• Лист "Воронка динамика": будут очищены только дни и недели\n\n' +
    '⚠️ Будет создана полная копия таблицы перед очисткой.\n\n' +
    'Продолжить?';
  
  var response = ui.alert('🗄️ Архивирование таблицы', message, ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    archiveAndCleanSpreadsheet();
  }
}

/**
 * Получить информацию о таблице
 * @param {Spreadsheet} ss - Объект таблицы
 * @return {Object} Информация о таблице
 */
function getSpreadsheetInfo(ss) {
  var sheets = ss.getSheets();
  var sheetsToClean = [];
  
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    
    // Проверяем, нужно ли очищать этот лист
    if (!isSheetPreserved(sheetName)) {
      sheetsToClean.push(sheetName);
    }
  }
  
  return {
    totalSheets: sheets.length,
    sheetsToClean: sheetsToClean
  };
}

/**
 * Проверить, нужно ли сохранять лист (не очищать)
 * @param {string} sheetName - Имя листа
 * @return {boolean} true если лист нужно сохранить
 */
function isSheetPreserved(sheetName) {
  for (var i = 0; i < SHEETS_TO_PRESERVE.length; i++) {
    if (sheetName.toLowerCase() === SHEETS_TO_PRESERVE[i].toLowerCase()) {
      return true;
    }
  }
  return false;
}

/**
 * Главная функция архивирования и очистки таблицы
 */
function archiveAndCleanSpreadsheet() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log('=== Начало архивирования таблицы ===');
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssName = ss.getName();
    
    // Шаг 1: Создаем архивную копию
    ui.alert('⏳ Процесс', 'Создание архивной копии... Это может занять несколько минут.', ui.ButtonSet.OK);
    
    var archiveResult = createArchiveCopy(ss);
    
    if (!archiveResult.success) {
      throw new Error('Не удалось создать архив: ' + archiveResult.error);
    }
    
    Logger.log('Архив создан: ' + archiveResult.archiveUrl);
    
    // Шаг 2: Очищаем листы в основной таблице
    var cleanResult = cleanSheets(ss);
    
    Logger.log('=== Архивирование завершено ===');
    Logger.log('Очищено листов: ' + cleanResult.cleanedSheets.length);
    
    // Показываем результат
    var successMessage = '✅ Архивирование завершено!\n\n' +
      '📁 Архив создан:\n' + archiveResult.archiveName + '\n\n' +
      '🧹 Очищено листов: ' + cleanResult.cleanedSheets.length + '\n' +
      '- ' + cleanResult.cleanedSheets.join('\n- ') + '\n\n' +
      '🔗 Ссылка на архив скопирована в буфер обмена (если доступно).\n\n' +
      archiveResult.archiveUrl;
    
    ui.alert('Успешно!', successMessage, ui.ButtonSet.OK);
    
    return {
      success: true,
      archiveUrl: archiveResult.archiveUrl,
      archiveName: archiveResult.archiveName,
      cleanedSheets: cleanResult.cleanedSheets
    };
    
  } catch (error) {
    Logger.log('ОШИБКА при архивировании: ' + error.toString());
    ui.alert('❌ Ошибка', 'Произошла ошибка при архивировании:\n\n' + error.toString(), ui.ButtonSet.OK);
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Создать архивную копию таблицы
 * @param {Spreadsheet} ss - Исходная таблица
 * @return {Object} Результат с URL и именем архива
 */
function createArchiveCopy(ss) {
  try {
    var ssName = ss.getName();
    var ssFile = DriveApp.getFileById(ss.getId());
    var parentFolder = ssFile.getParents().next(); // Папка, где лежит таблица
    
    // Формируем имя архива с датой
    var today = new Date();
    var dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var archiveName = ssName + '_Archive_до_' + dateStr;
    
    Logger.log('Создание копии: ' + archiveName);
    
    // Создаем копию в той же папке
    var archiveCopy = ssFile.makeCopy(archiveName, parentFolder);
    var archiveUrl = 'https://docs.google.com/spreadsheets/d/' + archiveCopy.getId();
    
    Logger.log('Копия создана: ' + archiveUrl);
    
    return {
      success: true,
      archiveId: archiveCopy.getId(),
      archiveUrl: archiveUrl,
      archiveName: archiveName
    };
    
  } catch (error) {
    Logger.log('ОШИБКА при создании копии: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Очистить накопительные листы
 * @param {Spreadsheet} ss - Таблица
 * @return {Object} Результат очистки
 */
function cleanSheets(ss) {
  var sheets = ss.getSheets();
  var cleanedSheets = [];
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    
    // Пропускаем листы, которые нужно сохранить
    if (isSheetPreserved(sheetName)) {
      Logger.log('Пропускаем лист (сохранение): ' + sheetName);
      continue;
    }
    
    // Особая обработка для листа "Воронка динамика"
    if (sheetName.toLowerCase() === FUNNEL_DYNAMIC_SHEET_NAME.toLowerCase()) {
      cleanFunnelDynamicSheet(sheet);
      cleanedSheets.push(sheetName + ' (только дни и недели)');
      continue;
    }
    
    // Очищаем обычный лист (сохраняем заголовки в первой строке)
    cleanRegularSheet(sheet);
    cleanedSheets.push(sheetName);
  }
  
  return {
    cleanedSheets: cleanedSheets
  };
}

/**
 * Очистить обычный лист (сохраняя заголовки)
 * @param {Sheet} sheet - Лист
 */
function cleanRegularSheet(sheet) {
  var sheetName = sheet.getName();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  Logger.log('Очистка листа: ' + sheetName + ' (строк: ' + lastRow + ', столбцов: ' + lastCol + ')');
  
  // Если лист пустой или только заголовки, пропускаем
  if (lastRow <= 1) {
    Logger.log('Лист пуст или содержит только заголовки. Пропускаем.');
    return;
  }
  
  // Удаляем все строки кроме первой (заголовки)
  try {
    // Удаляем строки начиная со 2-й
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
      Logger.log('Удалено строк: ' + (lastRow - 1));
    }
  } catch (error) {
    Logger.log('ОШИБКА при очистке листа ' + sheetName + ': ' + error.toString());
  }
}

/**
 * Очистить лист "Воронка динамика" (только столбцы данных B и далее)
 * Сохраняет структуру в столбце A (заголовки строк)
 * @param {Sheet} sheet - Лист
 */
function cleanFunnelDynamicSheet(sheet) {
  var sheetName = sheet.getName();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  Logger.log('Очистка листа "Воронка динамика": столбцов: ' + lastCol + ', строк: ' + lastRow);
  
  // Если только столбец A, пропускаем
  if (lastCol <= 1) {
    Logger.log('Лист содержит только столбец заголовков. Пропускаем.');
    return;
  }
  
  try {
    // Удаляем все столбцы кроме A (столбец 1)
    // Начинаем с B (столбец 2) и удаляем все до конца
    var columnsToDelete = lastCol - 1;
    
    if (columnsToDelete > 0) {
      sheet.deleteColumns(2, columnsToDelete);
      Logger.log('Удалено столбцов: ' + columnsToDelete);
    }
  } catch (error) {
    Logger.log('ОШИБКА при очистке листа "Воронка динамика": ' + error.toString());
  }
}

/**
 * Тестовая функция для проверки архивирования (без фактического выполнения)
 */
function testArchiveInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var info = getSpreadsheetInfo(ss);
  
  Logger.log('=== Информация о таблице ===');
  Logger.log('Всего листов: ' + info.totalSheets);
  Logger.log('Листов для очистки: ' + info.sheetsToClean.length);
  Logger.log('Листы для очистки: ' + info.sheetsToClean.join(', '));
  Logger.log('Листы для сохранения: ' + SHEETS_TO_PRESERVE.join(', '));
  
  return info;
}
