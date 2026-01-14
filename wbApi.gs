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
 * Использует endpoint /api/v2/stocks-report/products/products для получения данных по товарам
 * @param {Object} options - Опции запроса
 * @param {Array<number>} options.nmIDs - Список артикулов WB для фильтрации (опционально)
 * @param {string} options.stockType - Тип складов: "" (все), "wb" (склады WB), "mp" (склады продавца)
 * @param {number} options.limit - Лимит записей (по умолчанию 1000, максимум 1000)
 * @return {Object} Ответ API с данными по остаткам
 */
function getStocksReport(options) {
  options = options || {};
  var url = 'https://seller-analytics-api.wildberries.ru/api/v2/stocks-report/products/products';
  
  // Получаем текущую дату для периода
  var today = new Date();
  var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  var payload = {
    currentPeriod: {
      start: todayStr,
      end: todayStr
    },
    stockType: options.stockType || '',
    skipDeletedNm: true,
    limit: options.limit || 1000,
    offset: 0
  };
  
  // Добавляем фильтры если указаны
  if (options.nmIDs && options.nmIDs.length > 0) {
    payload.nmIDs = options.nmIDs;
  }
  
  Logger.log('Загрузка остатков товаров...');
  
  var allProducts = [];
  var offset = 0;
  
  while (true) {
    try {
      payload.offset = offset;
      var response = fetchWBAPIPost(url, payload);
      
      if (!response || !response.data || !response.data.products || response.data.products.length === 0) {
        Logger.log('Получены все данные. Всего товаров: ' + allProducts.length);
        break;
      }
      
      allProducts = allProducts.concat(response.data.products);
      Logger.log('Загружено товаров в этой итерации: ' + response.data.products.length + ', всего: ' + allProducts.length);
      
      // Если получено меньше лимита, значит это последняя страница
      if (response.data.products.length < payload.limit) {
        break;
      }
      
      offset += payload.limit;
      
      // Задержка для соблюдения лимитов API (3 запроса в минуту, интервал 20 секунд)
      Utilities.sleep(21000); // 21 секунда
      
    } catch (error) {
      Logger.log('Ошибка при загрузке остатков: ' + error.toString());
      throw error;
    }
  }
  
  return {
    data: {
      products: allProducts
    }
  };
}
