/**
 * Модуль для отчета "Аналитика продавца - Воронка продаж"
 * Выгружает данные через API за предыдущий день с дозаписыванием
 */

/**
 * Основная функция для синхронизации воронки продаж
 * Выгружает данные за предыдущий день и дозаписывает в таблицу
 */
function syncSalesFunnel() {
  try {
    Logger.log('=== Начало синхронизации воронки продаж ===');
    
    // 1. Получаем дату предыдущего дня
    var reportDate = getPreviousDay();
    Logger.log('Дата отчета: ' + reportDate);
    
    // 2. Проверяем, есть ли уже данные за эту дату
    var sheetName = getSalesFunnelSheetName();
    var sheet = getOrCreateSheet(sheetName);
    
    // Инициализируем заголовки если лист пустой
    if (isSheetEmpty(sheet)) {
      var headers = getSalesFunnelHeaders();
      setSheetHeaders(sheet, headers);
      Logger.log('Заголовки установлены');
    }
    
    // Проверяем дублирование данных (столбец 1 - дата)
    if (dateExistsInSheet(sheet, reportDate, 1)) {
      Logger.log('Данные за дату ' + reportDate + ' уже существуют в таблице. Пропускаем.');
      return;
    }
    
    // 3. Создаем задание на генерацию отчета
    var downloadId = createSalesFunnelReport(reportDate, reportDate);
    if (!downloadId) {
      throw new Error('Не удалось создать задание на генерацию отчета');
    }
    
    Logger.log('Создано задание с ID: ' + downloadId);
    
    // 4. Ждем готовности отчета
    var maxAttempts = 60; // Максимум 10 минут ожидания
    var attempt = 0;
    var status = '';
    
    while (attempt < maxAttempts) {
      attempt++;
      Logger.log('Проверка статуса, попытка ' + attempt + ' из ' + maxAttempts);
      
      status = checkSalesFunnelReportStatus(downloadId);
      Logger.log('Статус отчета: ' + status);
      
      if (status === 'DONE') {
        Logger.log('Отчет готов!');
        break;
      } else if (status === 'FAILED') {
        throw new Error('Генерация отчета завершилась с ошибкой. Попробуйте создать отчет повторно.');
      }
      
      // Если отчет еще не готов, ждем 10 секунд
      if (attempt < maxAttempts) {
        Logger.log('Отчет еще не готов, ждем 10 секунд...');
        Utilities.sleep(10000); // 10 секунд
      }
    }
    
    if (status !== 'DONE') {
      throw new Error('Превышено время ожидания готовности отчета. Последний статус: ' + status);
    }
    
    // 5. Получаем готовый отчет (CSV из ZIP)
    Logger.log('Загрузка готового отчета...');
    var csvData = downloadSalesFunnelReport(downloadId);
    
    if (!csvData || csvData.length === 0) {
      Logger.log('Нет данных в отчете за ' + reportDate);
      return;
    }
    
    Logger.log('Получено записей из CSV: ' + csvData.length);
    
    // 6. Форматируем данные для записи
    var formattedData = formatSalesFunnelData(csvData);
    
    if (!formattedData || formattedData.length === 0) {
      Logger.log('Нет данных для записи после форматирования');
      return;
    }
    
    // 7. Дозаписываем в таблицу
    appendDataToSheet(sheet, formattedData);
    
    Logger.log('=== Синхронизация завершена успешно. Записано строк: ' + formattedData.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации воронки продаж: ' + error.toString());
    Logger.log('Стек ошибки: ' + error.stack);
    throw error;
  }
}

/**
 * Создать задание на генерацию отчета воронки продаж
 * @param {string} startDate - Дата начала периода (YYYY-MM-DD)
 * @param {string} endDate - Дата конца периода (YYYY-MM-DD)
 * @return {string} ID задания (downloadId)
 */
function createSalesFunnelReport(startDate, endDate) {
  var url = 'https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads';
  var token = getWBApiToken();
  
  // Генерируем UUID для отчета
  var reportId = generateUUID();
  
  // Параметры отчета
  var payload = {
    id: reportId,
    reportType: 'DETAIL_HISTORY_REPORT',
    userReportName: 'Воронка продаж ' + startDate,
    params: {
      nmIDs: [], // Пустой массив = все товары
      startDate: startDate,
      endDate: endDate,
      timezone: 'Europe/Moscow',
      aggregationLevel: 'day'
    }
  };
  
  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    var statusCode = response.getResponseCode();
    
    if (statusCode !== 200) {
      var errorText = response.getContentText();
      Logger.log('Ошибка создания задания: ' + statusCode + ' - ' + errorText);
      throw new Error('Ошибка создания задания: ' + statusCode + ' - ' + errorText);
    }
    
    // API возвращает пустой ответ, используем сгенерированный ID
    return reportId;
    
  } catch (error) {
    Logger.log('Ошибка при создании задания: ' + error.toString());
    throw error;
  }
}

/**
 * Проверить статус готовности отчета
 * @param {string} downloadId - ID отчета
 * @return {string} Статус (PROCESSING, DONE, FAILED)
 */
function checkSalesFunnelReportStatus(downloadId) {
  var url = 'https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads';
  var token = getWBApiToken();
  
  // Формируем URL с параметрами фильтра
  var queryString = 'filter[downloadIds]=' + encodeURIComponent(downloadId);
  var fullUrl = url + '?' + queryString;
  
  try {
    var response = UrlFetchApp.fetch(fullUrl, {
      method: 'get',
      headers: {
        'Authorization': token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    var statusCode = response.getResponseCode();
    
    if (statusCode !== 200) {
      var errorText = response.getContentText();
      Logger.log('Ошибка проверки статуса: ' + statusCode + ' - ' + errorText);
      throw new Error('Ошибка проверки статуса: ' + statusCode + ' - ' + errorText);
    }
    
    var responseText = response.getContentText();
    var result = JSON.parse(responseText);
    
    if (!result || !result.data || !Array.isArray(result.data) || result.data.length === 0) {
      Logger.log('Отчет не найден в списке');
      return 'PROCESSING';
    }
    
    // Находим наш отчет по ID
    var report = result.data.find(function(item) {
      return item.id === downloadId;
    });
    
    if (!report) {
      Logger.log('Отчет с ID ' + downloadId + ' не найден');
      return 'PROCESSING';
    }
    
    return report.status || 'PROCESSING';
    
  } catch (error) {
    Logger.log('Ошибка при проверке статуса: ' + error.toString());
    throw error;
  }
}

/**
 * Скачать готовый отчет в формате CSV из ZIP архива
 * @param {string} downloadId - ID отчета
 * @return {Array<Object>} Массив объектов с данными из CSV
 */
function downloadSalesFunnelReport(downloadId) {
  var url = 'https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads/file/' + downloadId;
  var token = getWBApiToken();
  
  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': token
      },
      muteHttpExceptions: true
    });
    
    var statusCode = response.getResponseCode();
    
    if (statusCode === 204) {
      Logger.log('API вернул 204 - нет данных');
      return [];
    }
    
    if (statusCode !== 200) {
      var errorText = response.getContentText();
      Logger.log('Ошибка загрузки отчета: ' + statusCode + ' - ' + errorText);
      throw new Error('Ошибка загрузки отчета: ' + statusCode + ' - ' + errorText);
    }
    
    // Получаем ZIP архив
    var blob = response.getBlob();
    
    // Распаковываем ZIP
    var unzippedBlobs = Utilities.unzip(blob);
    
    if (!unzippedBlobs || unzippedBlobs.length === 0) {
      Logger.log('ZIP архив пустой');
      return [];
    }
    
    // Берем первый файл из архива (должен быть CSV)
    var csvBlob = unzippedBlobs[0];
    var csvContent = csvBlob.getDataAsString('UTF-8');
    
    // Парсим CSV
    var csvData = parseCSV(csvContent);
    
    return csvData;
    
  } catch (error) {
    Logger.log('Ошибка при загрузке отчета: ' + error.toString());
    throw error;
  }
}

/**
 * Парсинг CSV файла
 * @param {string} csvContent - Содержимое CSV файла
 * @return {Array<Object>} Массив объектов с данными
 */
function parseCSV(csvContent) {
  if (!csvContent) return [];
  
  var lines = csvContent.split('\n');
  if (lines.length === 0) return [];
  
  // Первая строка - заголовки
  var headers = lines[0].split(';').map(function(h) {
    return h.trim().replace(/^"/, '').replace(/"$/, '');
  });
  
  var result = [];
  
  // Парсим данные
  for (var i = 1; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    
    var values = line.split(';').map(function(v) {
      return v.trim().replace(/^"/, '').replace(/"$/, '');
    });
    
    if (values.length === headers.length) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = values[j];
      }
      result.push(row);
    }
  }
  
  return result;
}

/**
 * Форматировать данные из CSV в формат для таблицы
 * @param {Array<Object>} csvData - Массив объектов из CSV
 * @return {Array<Array>} Массив строк для Google Sheets
 */
function formatSalesFunnelData(csvData) {
  var data = [];
  
  for (var i = 0; i < csvData.length; i++) {
    var r = csvData[i];
    
    var row = [
      r['Дата'] || r['date'] || '',                          // 1. Дата
      r['Артикул WB'] || r['nmId'] || '',                    // 2. Артикул WB
      r['Артикул продавца'] || r['vendorCode'] || '',        // 3. Артикул продавца
      r['Наименование'] || r['name'] || '',                  // 4. Наименование
      r['Предмет'] || r['subject'] || '',                    // 5. Предмет
      r['Бренд'] || r['brand'] || '',                        // 6. Бренд
      parseNumber(r['Открытий карточек'] || r['cardViews'] || 0),       // 7. Открытий карточек (переходы в карточку)
      parseNumber(r['Добавлено в корзину'] || r['addedToCart'] || 0),   // 8. Добавлено в корзину
      parseNumber(r['Заказано товаров'] || r['orders'] || 0),           // 9. Заказано товаров
      parseNumber(r['Заказано на сумму'] || r['ordersSum'] || 0),       // 10. Заказано на сумму
      parseNumber(r['Выкупили товаров'] || r['buyouts'] || 0),          // 11. Выкупили товаров (если есть)
      parseNumber(r['Выкупили на сумму'] || r['buyoutsSum'] || 0)       // 12. Выкупили на сумму (если есть)
    ];
    
    data.push(row);
  }
  
  return data;
}

/**
 * Получить заголовки для листа воронки продаж
 * @return {Array<string>} Массив заголовков
 */
function getSalesFunnelHeaders() {
  return [
    'Дата',
    'Артикул WB',
    'Артикул продавца',
    'Наименование',
    'Предмет',
    'Бренд',
    'Переходы в карточку',
    'Положили в корзину',
    'Заказы',
    'Заказы на сумму',
    'Выкупы',
    'Выкупы на сумму'
  ];
}

/**
 * Парсинг числа из строки
 * @param {*} value - Значение для парсинга
 * @return {number} Число или 0
 */
function parseNumber(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  
  // Убираем пробелы и заменяем запятую на точку
  var cleaned = String(value).replace(/\s/g, '').replace(',', '.');
  var num = parseFloat(cleaned);
  
  return isNaN(num) ? 0 : num;
}

/**
 * Генерация UUID v4
 * @return {string} UUID
 */
function generateUUID() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0;
    var v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}
