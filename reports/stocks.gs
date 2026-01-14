/**
 * Модуль для остатков товаров
 * Выгружает данные с полной перезаписью раз в сутки
 */

/**
 * Заголовки столбцов для листа "Остатки"
 * Первые столбцы фиксированные, затем идут столбцы с названиями складов
 */
function getStocksHeaders() {
  // Базовые столбцы
  var baseHeaders = [
    'Бренд',
    'Предмет',
    'Артикул продавца',
    'Артикул WB',
    'Объем, л',
    'Баркод',
    'Размер вещи',
    'В пути до получателей',
    'В пути возвраты на склад WB',
    'Всего находится на складах'
  ];
  
  // Список складов (будет дополнен динамически при получении данных)
  var warehouseHeaders = [
    'Коледино',
    'Подольск',
    'Казань',
    'Электросталь',
    'Санкт-Петербург Уткина Заводь',
    'Краснодар',
    'Новосибирск',
    'Екатеринбург - Испытателей 14г',
    'Екатеринбург - Перспективный 12',
    'Тула',
    'Атакент',
    'Белая дача',
    'Невинномысск',
    'Рязань (Тюшевское)',
    'Котовск',
    'Самара (Новосемейкино)',
    'Волгоград',
    'СЦ Барнаул',
    'Актобе',
    'Чашниково',
    'Астана Карагандинское шоссе',
    'Владимир',
    'Сарапул',
    'СПБ Шушары',
    'Воронеж',
    'Пенза',
    'Остальные'
  ];
  
  return baseHeaders.concat(warehouseHeaders);
}

/**
 * Получить список всех уникальных складов из данных API
 * @param {Array<Object>} products - Массив товаров из API
 * @return {Array<string>} Массив названий складов
 */
function extractWarehouseNames(products) {
  var warehouses = {};
  
  for (var i = 0; i < products.length; i++) {
    var product = products[i];
    if (product.offices && Array.isArray(product.offices)) {
      for (var j = 0; j < product.offices.length; j++) {
        var office = product.offices[j];
        if (office.officeName) {
          warehouses[office.officeName] = true;
        }
      }
    }
  }
  
  return Object.keys(warehouses).sort();
}

/**
 * Преобразовать товар из API в строку для таблицы
 * @param {Object} product - Товар из API
 * @param {Array<string>} warehouseNames - Список названий складов в порядке столбцов
 * @return {Array} Массив значений для строки таблицы
 */
function convertProductToRow(product, warehouseNames) {
  // Базовые данные товара
  var row = [
    product.brandName || '', // Бренд
    product.subjectName || '', // Предмет
    product.supplierArticle || '', // Артикул продавца
    product.nmId || '', // Артикул WB
    product.volume || '', // Объем, л
    product.barcode || '', // Баркод
    product.techSize || '', // Размер вещи
    product.metrics ? (product.metrics.toClientCount || 0) : 0, // В пути до получателей
    product.metrics ? (product.metrics.fromClientCount || 0) : 0, // В пути возвраты на склад WB
    product.metrics ? (product.metrics.stockCount || 0) : 0 // Всего находится на складах
  ];
  
  // Создаем карту остатков по складам
  var stockByWarehouse = {};
  var otherStock = 0; // Остатки на складах, не входящих в основной список
  
  if (product.offices && Array.isArray(product.offices)) {
    for (var i = 0; i < product.offices.length; i++) {
      var office = product.offices[i];
      if (office.officeName && office.metrics) {
        var stock = office.metrics.stockCount || 0;
        stockByWarehouse[office.officeName] = stock;
      }
    }
  }
  
  // Добавляем остатки по каждому складу в порядке столбцов
  for (var j = 0; j < warehouseNames.length; j++) {
    var warehouseName = warehouseNames[j];
    if (warehouseName === 'Остальные') {
      // Для столбца "Остальные" суммируем остатки со всех складов, не входящих в основной список
      var totalStock = product.metrics ? (product.metrics.stockCount || 0) : 0;
      var knownStock = 0;
      for (var k = 0; k < warehouseNames.length - 1; k++) {
        knownStock += stockByWarehouse[warehouseNames[k]] || 0;
      }
      otherStock = Math.max(0, totalStock - knownStock);
      row.push(otherStock);
    } else {
      row.push(stockByWarehouse[warehouseName] || 0);
    }
  }
  
  return row;
}

/**
 * Основная функция для выгрузки остатков товаров
 * Выгружает данные с полной перезаписью листа
 */
function syncStocks() {
  try {
    Logger.log('=== Начало синхронизации остатков товаров ===');
    
    // Получаем лист
    var sheetName = getStocksSheetName();
    var sheet = getOrCreateSheet(sheetName);
    
    // Загружаем данные из API
    Logger.log('Загрузка остатков из API Wildberries...');
    var response = getStocksReport({
      stockType: '', // Все склады
      limit: 1000
    });
    
    if (!response || !response.data || !response.data.products || response.data.products.length === 0) {
      Logger.log('Нет данных об остатках для загрузки');
      // Очищаем лист и оставляем только заголовки
      var headers = getStocksHeaders();
      clearAndWriteSheet(sheet, headers, []);
      return;
    }
    
    var products = response.data.products;
    Logger.log('Получено товаров из API: ' + products.length);
    
    // Извлекаем список всех складов из данных
    var warehouseNames = extractWarehouseNames(products);
    Logger.log('Найдено складов: ' + warehouseNames.length);
    
    // Формируем заголовки: базовые + склады
    var baseHeaders = [
      'Бренд',
      'Предмет',
      'Артикул продавца',
      'Артикул WB',
      'Объем, л',
      'Баркод',
      'Размер вещи',
      'В пути до получателей',
      'В пути возвраты на склад WB',
      'Всего находится на складах'
    ];
    
    // Добавляем склады из данных, затем стандартные склады, которых нет в данных
    var standardWarehouses = [
      'Коледино',
      'Подольск',
      'Казань',
      'Электросталь',
      'Санкт-Петербург Уткина Заводь',
      'Краснодар',
      'Новосибирск',
      'Екатеринбург - Испытателей 14г',
      'Екатеринбург - Перспективный 12',
      'Тула',
      'Атакент',
      'Белая дача',
      'Невинномысск',
      'Рязань (Тюшевское)',
      'Котовск',
      'Самара (Новосемейкино)',
      'Волгоград',
      'СЦ Барнаул',
      'Актобе',
      'Чашниково',
      'Астана Карагандинское шоссе',
      'Владимир',
      'Сарапул',
      'СПБ Шушары',
      'Воронеж',
      'Пенза',
      'Остальные'
    ];
    
    // Объединяем склады: сначала стандартные, затем остальные из данных
    var allWarehouses = [];
    var warehouseSet = {};
    
    // Добавляем стандартные склады
    for (var i = 0; i < standardWarehouses.length; i++) {
      allWarehouses.push(standardWarehouses[i]);
      warehouseSet[standardWarehouses[i]] = true;
    }
    
    // Добавляем склады из данных, которых нет в стандартном списке
    for (var j = 0; j < warehouseNames.length; j++) {
      if (!warehouseSet[warehouseNames[j]]) {
        allWarehouses.push(warehouseNames[j]);
        warehouseSet[warehouseNames[j]] = true;
      }
    }
    
    var headers = baseHeaders.concat(allWarehouses);
    
    // Преобразуем данные в формат таблицы
    var rows = [];
    for (var k = 0; k < products.length; k++) {
      var row = convertProductToRow(products[k], allWarehouses);
      rows.push(row);
    }
    
    // Полностью перезаписываем лист
    Logger.log('Запись данных в таблицу (полная перезапись)...');
    clearAndWriteSheet(sheet, headers, rows);
    
    Logger.log('=== Синхронизация завершена успешно. Записано товаров: ' + rows.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации остатков: ' + error.toString());
    Logger.log('Стек ошибки: ' + error.stack);
    throw error;
  }
}
