/**
 * Модуль для остатков товаров
 * Выгружает данные с полной перезаписью раз в сутки
 */

/**
 * Заголовки столбцов для листа "Остатки"
 * Первые столбцы фиксированные, затем идут столбцы с названиями складов из API
 */
function getStocksHeaders() {
  // Базовые столбцы (склады будут добавлены динамически при получении данных)
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
  
  return baseHeaders;
}

/**
 * Получить список всех уникальных складов из данных API
 * Исключаем служебные "склады" (В пути..., Всего находится..., Остальные)
 * @param {Array<Object>} products - Массив товаров из API
 * @return {Array<string>} Массив названий складов
 */
function extractWarehouseNames(products) {
  var warehouses = {};
  var excludeNames = [
    'В пути до получателей',
    'В пути возвраты на склад WB',
    'Всего находится на складах',
    'Остальные'
  ];
  
  for (var i = 0; i < products.length; i++) {
    var product = products[i];
    if (product.warehouses && Array.isArray(product.warehouses)) {
      for (var j = 0; j < product.warehouses.length; j++) {
        var warehouse = product.warehouses[j];
        if (warehouse.warehouseName && excludeNames.indexOf(warehouse.warehouseName) === -1) {
          warehouses[warehouse.warehouseName] = true;
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
  // Создаем карту остатков по складам
  var stockByWarehouse = {};
  var toClientsCount = '';
  var fromClientsCount = '';
  var totalStock = '';
  
  if (product.warehouses && Array.isArray(product.warehouses)) {
    for (var i = 0; i < product.warehouses.length; i++) {
      var warehouse = product.warehouses[i];
      if (warehouse.warehouseName && warehouse.quantity !== undefined && warehouse.quantity !== null) {
        // Проверяем служебные "склады"
        if (warehouse.warehouseName === 'В пути до получателей') {
          toClientsCount = warehouse.quantity;
        } else if (warehouse.warehouseName === 'В пути возвраты на склад WB') {
          fromClientsCount = warehouse.quantity;
        } else if (warehouse.warehouseName === 'Всего находится на складах') {
          totalStock = warehouse.quantity;
        } else {
          // Обычный склад
          stockByWarehouse[warehouse.warehouseName] = warehouse.quantity;
        }
      }
    }
  }
  
  // Базовые данные товара
  var row = [
    product.brand || '', // Бренд
    product.subjectName || '', // Предмет
    product.vendorCode || '', // Артикул продавца
    product.nmId || '', // Артикул WB
    product.volume || '', // Объем, л
    product.barcode || '', // Баркод
    product.techSize || '', // Размер вещи
    toClientsCount, // В пути до получателей
    fromClientsCount, // В пути возвраты на склад WB
    totalStock // Всего находится на складах
  ];
  
  // Добавляем остатки по каждому складу в порядке столбцов
  for (var j = 0; j < warehouseNames.length; j++) {
    var warehouseName = warehouseNames[j];
    if (warehouseName === 'Остальные') {
      // Для столбца "Остальные" суммируем остатки со всех складов, не входящих в основной список
      if (totalStock !== '' && typeof totalStock === 'number') {
        var knownStock = 0;
        for (var k = 0; k < warehouseNames.length - 1; k++) {
          var knownWarehouseStock = stockByWarehouse[warehouseNames[k]];
          if (knownWarehouseStock !== undefined && knownWarehouseStock !== null) {
            knownStock += knownWarehouseStock;
          }
        }
        var otherStock = Math.max(0, totalStock - knownStock);
        // Если otherStock > 0, записываем значение, иначе пустое место
        row.push(otherStock > 0 ? otherStock : '');
      } else {
        row.push('');
      }
    } else {
      // Если для склада есть остаток (включая 0), записываем его
      // Если остатка нет, записываем пустое место
      if (stockByWarehouse.hasOwnProperty(warehouseName)) {
        row.push(stockByWarehouse[warehouseName]);
      } else {
        row.push('');
      }
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
    
    // Логируем структуру ответа для отладки
    if (response) {
      Logger.log('Структура ответа: ' + JSON.stringify(Object.keys(response)));
      if (response.data) {
        Logger.log('Ключи в data: ' + JSON.stringify(Object.keys(response.data)));
        if (response.data.products) {
          Logger.log('Количество products: ' + response.data.products.length);
          if (response.data.products.length > 0) {
            Logger.log('Пример первого товара (первые 500 символов): ' + JSON.stringify(response.data.products[0]).substring(0, 500));
          }
        }
      }
    }
    
    if (!response || !response.data) {
      Logger.log('Нет данных в ответе API. Ответ: ' + JSON.stringify(response));
      var headers = getStocksHeaders();
      clearAndWriteSheet(sheet, headers, []);
      return;
    }
    
    // Проверяем разные возможные структуры ответа
    var products = response.data.products || response.data.items || [];
    
    if (!products || products.length === 0) {
      Logger.log('Нет данных об остатках для загрузки. Структура ответа: ' + JSON.stringify(response.data));
      // Очищаем лист и оставляем только заголовки
      var headers = getStocksHeaders();
      clearAndWriteSheet(sheet, headers, []);
      return;
    }
    
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
    
    // Используем склады из API данных + добавляем "Остальные" в конце
    var allWarehouses = warehouseNames.slice(); // Копируем массив
    allWarehouses.push('Остальные');
    
    var headers = baseHeaders.concat(allWarehouses);
    
    // Преобразуем данные в формат таблицы
    var rows = [];
    Logger.log('Начало преобразования ' + products.length + ' товаров в строки таблицы...');
    
    for (var k = 0; k < products.length; k++) {
      try {
        var row = convertProductToRow(products[k], allWarehouses);
        if (row && row.length > 0) {
          rows.push(row);
        } else {
          Logger.log('Предупреждение: товар ' + k + ' не преобразован в строку. Данные: ' + JSON.stringify(products[k]).substring(0, 200));
        }
      } catch (e) {
        Logger.log('Ошибка преобразования товара ' + k + ': ' + e.toString());
        Logger.log('Данные товара: ' + JSON.stringify(products[k]).substring(0, 300));
      }
    }
    
    Logger.log('Преобразовано строк для записи: ' + rows.length);
    
    if (rows.length > 0) {
      Logger.log('Пример первой строки (первые 15 значений): ' + JSON.stringify(rows[0].slice(0, 15)));
      Logger.log('Количество столбцов в первой строке: ' + rows[0].length);
      Logger.log('Количество столбцов в заголовках: ' + headers.length);
      
      // Проверяем соответствие количества столбцов
      if (rows[0].length !== headers.length) {
        Logger.log('ВНИМАНИЕ: Несоответствие количества столбцов! Заголовков: ' + headers.length + ', в строке: ' + rows[0].length);
      }
    }
    
    // Полностью перезаписываем лист
    Logger.log('Запись данных в таблицу (полная перезапись)...');
    if (rows.length > 0) {
      clearAndWriteSheet(sheet, headers, rows);
    } else {
      Logger.log('ВНИМАНИЕ: Нет строк для записи! Проверьте структуру данных API.');
      // Оставляем только заголовки
      clearAndWriteSheet(sheet, headers, []);
    }
    
    Logger.log('=== Синхронизация завершена успешно. Записано товаров: ' + rows.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации остатков: ' + error.toString());
    Logger.log('Стек ошибки: ' + error.stack);
    throw error;
  }
}
