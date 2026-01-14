/**
 * Модуль для работы с API Wildberries
 */

/**
 * Получить токен API из свойств скрипта
 * @return {string} API токен
 */
function getWBApiToken() {
  var token = PropertiesService.getScriptProperties().getProperty('WB_API_TOKEN');
  if (!token) {
    throw new Error('WB_API_TOKEN не установлен в свойствах скрипта. Установите токен через меню: Файл -> Свойства проекта -> Свойства скрипта');
  }
  return token;
}

/**
 * Выполнить запрос к API Wildberries
 * @param {string} url - URL запроса
 * @param {Object} params - Параметры запроса
 * @return {Object} Ответ API
 */
function fetchWBAPI(url, params) {
  var token = getWBApiToken();
  
  // Формируем URL с параметрами
  var queryString = Object.keys(params).map(function(key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');
  
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
    
    if (statusCode === 204) {
      // Нет данных
      Logger.log('API вернул 204 - нет данных');
      return null;
    }
    
    if (statusCode !== 200) {
      var errorText = response.getContentText();
      Logger.log('Ошибка API: ' + statusCode + ' - ' + errorText);
      throw new Error('Ошибка API: ' + statusCode + ' - ' + errorText);
    }
    
    var responseText = response.getContentText();
    if (!responseText) {
      return null;
    }
    
    return JSON.parse(responseText);
    
  } catch (error) {
    Logger.log('Ошибка при запросе к API: ' + error.toString());
    throw error;
  }
}

/**
 * Выполнить POST запрос к API Wildberries
 * @param {string} url - URL запроса
 * @param {Object} payload - Тело запроса (JSON объект)
 * @return {Object} Ответ API
 */
function fetchWBAPIPost(url, payload) {
  var token = getWBApiToken();
  
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
    
    if (statusCode === 204) {
      Logger.log('API вернул 204 - нет данных');
      return null;
    }
    
    if (statusCode !== 200) {
      var errorText = response.getContentText();
      Logger.log('Ошибка API: ' + statusCode + ' - ' + errorText);
      throw new Error('Ошибка API: ' + statusCode + ' - ' + errorText);
    }
    
    var responseText = response.getContentText();
    if (!responseText) {
      return null;
    }
    
    return JSON.parse(responseText);
    
  } catch (error) {
    Logger.log('Ошибка при POST запросе к API: ' + error.toString());
    throw error;
  }
}

/**
 * Получить детализацию отчета за период с пагинацией
 * @param {string} dateFrom - Дата начала в формате RFC3339
 * @param {string} dateTo - Дата конца в формате RFC3339
 * @param {string} period - Периодичность: "daily" или "weekly"
 * @return {Array<Object>} Массив записей отчета
 */
function getReportDetailByPeriod(dateFrom, dateTo, period) {
  period = period || 'daily';
  var allRecords = [];
  var rrdid = 0;
  var limit = 100000;
  var url = 'https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod';
  
  Logger.log('Начало загрузки отчета с ' + dateFrom + ' по ' + dateTo);
  
  while (true) {
    try {
      var params = {
        dateFrom: dateFrom,
        dateTo: dateTo,
        period: period,
        limit: limit,
        rrdid: rrdid
      };
      
      var response = fetchWBAPI(url, params);
      
      if (!response || response.length === 0) {
        Logger.log('Получены все данные. Всего записей: ' + allRecords.length);
        break;
      }
      
      allRecords = allRecords.concat(response);
      Logger.log('Загружено записей в этой итерации: ' + response.length + ', всего: ' + allRecords.length);
      
      // Получаем последний rrd_id для следующей итерации
      var lastRecord = response[response.length - 1];
      if (lastRecord && lastRecord.rrd_id) {
        rrdid = lastRecord.rrd_id;
      } else {
        // Если нет rrd_id, значит это последняя страница
        break;
      }
      
      // Небольшая задержка для соблюдения лимитов API (1 запрос в минуту)
      Utilities.sleep(61000); // 61 секунда
      
    } catch (error) {
      Logger.log('Ошибка при загрузке данных: ' + error.toString());
      throw error;
    }
  }
  
  return allRecords;
}

/**
 * Получить заказы за указанную дату с пагинацией
 * @param {string} dateFrom - Дата в формате RFC3339 (YYYY-MM-DD или YYYY-MM-DDTHH:mm:ss)
 * @param {number} flag - Флаг: 0 для пагинации по lastChangeDate, 1 для получения всех заказов за дату
 * @return {Array<Object>} Массив заказов
 */
function getOrders(dateFrom, flag) {
  flag = flag !== undefined ? flag : 1; // По умолчанию flag=1 для получения всех заказов за дату
  var allOrders = [];
  var currentDateFrom = dateFrom;
  var url = 'https://statistics-api.wildberries.ru/api/v1/supplier/orders';
  
  Logger.log('Начало загрузки заказов с ' + dateFrom + ', flag=' + flag);
  
  while (true) {
    try {
      var params = {
        dateFrom: currentDateFrom,
        flag: flag
      };
      
      var response = fetchWBAPI(url, params);
      
      if (!response || response.length === 0) {
        Logger.log('Получены все заказы. Всего заказов: ' + allOrders.length);
        break;
      }
      
      allOrders = allOrders.concat(response);
      Logger.log('Загружено заказов в этой итерации: ' + response.length + ', всего: ' + allOrders.length);
      
      // Если flag=0, используем lastChangeDate для пагинации
      if (flag === 0) {
        var lastRecord = response[response.length - 1];
        if (lastRecord && lastRecord.lastChangeDate) {
          currentDateFrom = lastRecord.lastChangeDate;
        } else {
          break;
        }
      } else {
        // Если flag=1, получаем все заказы за дату за один запрос
        break;
      }
      
      // Задержка для соблюдения лимитов API (1 запрос в минуту)
      Utilities.sleep(61000); // 61 секунда
      
    } catch (error) {
      Logger.log('Ошибка при загрузке заказов: ' + error.toString());
      throw error;
    }
  }
  
  return allOrders;
}

/**
 * Получить остатки товаров по складам
 * Использует новый трёхэтапный API:
 * 1. Создание задания на генерацию отчёта
 * 2. Проверка статуса задания
 * 3. Получение готового отчёта
 * @param {Object} options - Опции запроса (для совместимости, не используются)
 * @return {Object} Ответ API с данными по остаткам
 */
function getStocksReport(options) {
  options = options || {};
  
  Logger.log('=== Начало процесса получения отчёта об остатках ===');
  
  // Шаг 1: Создать задание на генерацию отчёта
  var taskId = createWarehouseRemainsTask();
  if (!taskId) {
    throw new Error('Не удалось создать задание на генерацию отчёта');
  }
  
  Logger.log('Создано задание с ID: ' + taskId);
  
  // Шаг 2: Ждём готовности отчёта
  var maxAttempts = 60; // Максимум 10 минут ожидания (60 попыток по 10 секунд)
  var attempt = 0;
  var status = '';
  
  while (attempt < maxAttempts) {
    attempt++;
    Logger.log('Проверка статуса, попытка ' + attempt + ' из ' + maxAttempts);
    
    status = checkWarehouseRemainsTaskStatus(taskId);
    Logger.log('Статус задания: ' + status);
    
    if (status === 'done') {
      Logger.log('Отчёт готов!');
      break;
    } else if (status === 'canceled' || status === 'purged') {
      throw new Error('Задание отклонено или удалено. Статус: ' + status);
    }
    
    // Если отчёт ещё не готов, ждём 10 секунд
    if (attempt < maxAttempts) {
      Logger.log('Отчёт ещё не готов, ждём 10 секунд...');
      Utilities.sleep(10000); // 10 секунд
    }
  }
  
  if (status !== 'done') {
    throw new Error('Превышено время ожидания готовности отчёта. Последний статус: ' + status);
  }
  
  // Шаг 3: Получить готовый отчёт
  Logger.log('Загрузка готового отчёта...');
  var products = downloadWarehouseRemainsReport(taskId);
  
  Logger.log('Загружено товаров: ' + (products ? products.length : 0));
  
  return {
    data: {
      products: products || []
    }
  };
}

/**
 * Создать задание на генерацию отчёта об остатках
 * @return {string} ID задания
 */
function createWarehouseRemainsTask() {
  var url = 'https://seller-analytics-api.wildberries.ru/api/v1/warehouse_remains';
  var token = getWBApiToken();
  
  // Параметры запроса - все группировки включены
  var params = {
    locale: 'ru',
    groupByBrand: true,
    groupBySubject: true,
    groupBySa: true,
    groupByNm: true,
    groupByBarcode: true,
    groupBySize: true
  };
  
  // Формируем URL с параметрами
  var queryString = Object.keys(params).map(function(key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');
  
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
      Logger.log('Ошибка создания задания: ' + statusCode + ' - ' + errorText);
      throw new Error('Ошибка создания задания: ' + statusCode + ' - ' + errorText);
    }
    
    var responseText = response.getContentText();
    var result = JSON.parse(responseText);
    
    if (!result || !result.data || !result.data.taskId) {
      Logger.log('Неверная структура ответа: ' + responseText);
      throw new Error('Не удалось получить ID задания из ответа API');
    }
    
    return result.data.taskId;
    
  } catch (error) {
    Logger.log('Ошибка при создании задания: ' + error.toString());
    throw error;
  }
}

/**
 * Проверить статус задания на генерацию отчёта
 * @param {string} taskId - ID задания
 * @return {string} Статус задания (new, processing, done, purged, canceled)
 */
function checkWarehouseRemainsTaskStatus(taskId) {
  var url = 'https://seller-analytics-api.wildberries.ru/api/v1/warehouse_remains/tasks/' + taskId + '/status';
  var token = getWBApiToken();
  
  try {
    var response = UrlFetchApp.fetch(url, {
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
    
    if (!result || !result.data || !result.data.status) {
      Logger.log('Неверная структура ответа: ' + responseText);
      throw new Error('Не удалось получить статус задания из ответа API');
    }
    
    return result.data.status;
    
  } catch (error) {
    Logger.log('Ошибка при проверке статуса: ' + error.toString());
    throw error;
  }
}

/**
 * Получить готовый отчёт об остатках
 * @param {string} taskId - ID задания
 * @return {Array<Object>} Массив товаров с остатками
 */
function downloadWarehouseRemainsReport(taskId) {
  var url = 'https://seller-analytics-api.wildberries.ru/api/v1/warehouse_remains/tasks/' + taskId + '/download';
  var token = getWBApiToken();
  
  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': token,
        'Content-Type': 'application/json'
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
      Logger.log('Ошибка загрузки отчёта: ' + statusCode + ' - ' + errorText);
      throw new Error('Ошибка загрузки отчёта: ' + statusCode + ' - ' + errorText);
    }
    
    var responseText = response.getContentText();
    if (!responseText) {
      return [];
    }
    
    var products = JSON.parse(responseText);
    
    // API возвращает массив напрямую
    if (!Array.isArray(products)) {
      Logger.log('Неожиданная структура ответа, ожидался массив: ' + responseText.substring(0, 200));
      return [];
    }
    
    return products;
    
  } catch (error) {
    Logger.log('Ошибка при загрузке отчёта: ' + error.toString());
    throw error;
  }
}
