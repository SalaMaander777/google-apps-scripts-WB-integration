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

/**
 * Получить список рекламных кампаний
 * @param {Object} options - Опции запроса
 * @param {string} options.ids - ID кампаний через запятую (максимум 50), опционально
 * @param {string} options.statuses - Статусы кампаний через запятую (опционально)
 * @param {string} options.payment_type - Тип оплаты: cpm или cpc (опционально)
 * @return {Array<Object>} Массив кампаний
 */
function getAdverts(options) {
  options = options || {};
  var url = 'https://advert-api.wildberries.ru/api/advert/v2/adverts';
  
  var params = {};
  if (options.ids) params.ids = options.ids;
  if (options.statuses) params.statuses = options.statuses;
  if (options.payment_type) params.payment_type = options.payment_type;
  
  try {
    var response = fetchWBAPI(url, params);
    
    if (!response) {
      Logger.log('Нет рекламных кампаний');
      return [];
    }
    
    // API возвращает объект с полем adverts
    if (response.adverts && Array.isArray(response.adverts)) {
      Logger.log('Получено кампаний: ' + response.adverts.length);
      return response.adverts;
    }
    
    Logger.log('Неожиданная структура ответа getAdverts');
    return [];
    
  } catch (error) {
    Logger.log('Ошибка при получении списка кампаний: ' + error.toString());
    throw error;
  }
}

/**
 * Получить полную статистику по рекламным кампаниям
 * @param {Array<number>} advertIds - Массив ID кампаний (максимум 50)
 * @param {string} beginDate - Дата начала в формате YYYY-MM-DD
 * @param {string} endDate - Дата конца в формате YYYY-MM-DD
 * @return {Array<Object>} Массив статистики по кампаниям
 */
function getAdvertFullStats(advertIds, beginDate, endDate) {
  if (!advertIds || advertIds.length === 0) {
    Logger.log('Нет ID кампаний для получения статистики');
    return [];
  }
  
  // API лимит - максимум 50 ID за раз
  if (advertIds.length > 50) {
    Logger.log('ПРЕДУПРЕЖДЕНИЕ: передано больше 50 ID кампаний, обрабатываем только первые 50');
    advertIds = advertIds.slice(0, 50);
  }
  
  var url = 'https://advert-api.wildberries.ru/adv/v3/fullstats';
  
  var params = {
    ids: advertIds.join(','),
    beginDate: beginDate,
    endDate: endDate
  };
  
  try {
    var response = fetchWBAPI(url, params);
    
    if (!response) {
      Logger.log('Нет статистики по кампаниям');
      return [];
    }
    
    // API возвращает массив напрямую
    if (Array.isArray(response)) {
      Logger.log('Получено статистик: ' + response.length);
      return response;
    }
    
    Logger.log('Неожиданная структура ответа getAdvertFullStats');
    return [];
    
  } catch (error) {
    Logger.log('Ошибка при получении статистики кампаний: ' + error.toString());
    // Для соблюдения лимитов API (3 запроса в минуту) не бросаем ошибку, возвращаем пустой массив
    Logger.log('Возвращаем пустой массив статистики');
    return [];
  }
}

/**
 * Получить историю затрат на рекламные кампании за период
 * @param {string} fromDate - Дата начала в формате YYYY-MM-DD
 * @param {string} toDate - Дата конца в формате YYYY-MM-DD (минимум 1 день, максимум 31 день)
 * @return {Array<Object>} Массив истории затрат
 */
function getAdvertCostsHistory(fromDate, toDate) {
  if (!fromDate || !toDate) {
    throw new Error('Необходимо указать даты начала и конца периода');
  }
  
  var url = 'https://advert-api.wildberries.ru/adv/v1/upd';
  
  var params = {
    from: fromDate,
    to: toDate
  };
  
  try {
    var response = fetchWBAPI(url, params);
    
    if (!response) {
      Logger.log('Нет данных по затратам на рекламу');
      return [];
    }
    
    // API возвращает массив напрямую
    if (Array.isArray(response)) {
      Logger.log('Получено записей затрат: ' + response.length);
      return response;
    }
    
    Logger.log('Неожиданная структура ответа getAdvertCostsHistory');
    return [];
    
  } catch (error) {
    Logger.log('Ошибка при получении истории затрат: ' + error.toString());
    throw error;
  }
}

/**
 * Получить статистику карточек товаров (воронка продаж) с пагинацией
 * @param {string} selectedStart - Дата начала запрашиваемого периода в формате YYYY-MM-DD
 * @param {string} selectedEnd - Дата конца запрашиваемого периода в формате YYYY-MM-DD
 * @param {string} pastStart - Дата начала периода для сравнения в формате YYYY-MM-DD
 * @param {string} pastEnd - Дата конца периода для сравнения в формате YYYY-MM-DD
 * @return {Array<Object>} Массив товаров со статистикой
 */
function getSalesFunnelData(selectedStart, selectedEnd, pastStart, pastEnd) {
  if (!selectedStart || !selectedEnd) {
    throw new Error('Необходимо указать период selectedPeriod (start и end)');
  }
  
  var url = 'https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products';
  
  var allProducts = [];
  var offset = 0;
  var limit = 1000; // Максимальный размер страницы
  
  Logger.log('Начало загрузки статистики карточек товаров за период ' + selectedStart + ' - ' + selectedEnd);
  
  while (true) {
    try {
      var payload = {
        selectedPeriod: {
          start: selectedStart,
          end: selectedEnd
        },
        nmIds: [],
        brandNames: [],
        subjectIds: [],
        tagIds: [],
        skipDeletedNm: false,
        orderBy: {
          field: 'openCard',
          mode: 'desc'
        },
        limit: limit,
        offset: offset
      };
      
      // Добавляем период для сравнения если указан
      if (pastStart && pastEnd) {
        payload.pastPeriod = {
          start: pastStart,
          end: pastEnd
        };
      }
      
      var response = fetchWBAPIPost(url, payload);
      
      if (!response) {
        Logger.log('Получены все данные. Всего товаров: ' + allProducts.length);
        break;
      }
      
      // API возвращает объект с data.products
      var products = [];
      if (response.data && response.data.products && Array.isArray(response.data.products)) {
        products = response.data.products;
      } else {
        Logger.log('Неожиданная структура ответа');
        break;
      }
      
      if (products.length === 0) {
        Logger.log('Получены все данные. Всего товаров: ' + allProducts.length);
        break;
      }
      
      allProducts = allProducts.concat(products);
      Logger.log('Загружено товаров в этой итерации: ' + products.length + ', всего: ' + allProducts.length);
      
      // Если получили меньше чем limit, значит это последняя страница
      if (products.length < limit) {
        Logger.log('Получена последняя страница');
        break;
      }
      
      // Переходим к следующей странице
      offset += limit;
      
      // Задержка для соблюдения лимитов API (3 запроса в минуту, интервал 20 секунд)
      if (products.length === limit) {
        Logger.log('Ожидание 20 секунд перед следующим запросом...');
        Utilities.sleep(20000); // 20 секунд
      }
      
    } catch (error) {
      Logger.log('Ошибка при загрузке данных: ' + error.toString());
      throw error;
    }
  }
  
  return allProducts;
}
