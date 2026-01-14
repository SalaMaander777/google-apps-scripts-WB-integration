/**
 * Модуль для ленты заказов
 * Выгружает данные за предыдущий день с дозаписыванием
 */

/**
 * Заголовки столбцов для листа "Лента заказов"
 */
function getOrdersFeedHeaders() {
  return [
    '№ позиции',
    'Баркод',
    'Наименование товара',
    'Артикул поставщика/цвет',
    'Артикул WB',
    'Размер',
    'Цена',
    'Склад отправки',
    'Регион доставки',
    'Время заказа',
    'Обновление',
    'КОЛ_ВО'
  ];
}

/**
 * Преобразовать заказ из API в строку для таблицы
 * @param {Object} order - Заказ из API
 * @param {number} rowNumber - Номер позиции (начиная с 1)
 * @return {Array} Массив значений для строки таблицы
 */
function convertOrderToRow(order, rowNumber) {
  // Формируем наименование товара из бренда и предмета
  var productName = '';
  if (order.brand) {
    productName = order.brand;
  }
  if (order.subject) {
    if (productName) {
      productName += ' ' + order.subject;
    } else {
      productName = order.subject;
    }
  }
  if (!productName && order.category) {
    productName = order.category;
  }
  
  return [
    rowNumber, // № позиции
    order.barcode || '',
    productName || '', // Наименование товара
    order.supplierArticle || '', // Артикул поставщика/цвет
    order.nmId || '', // Артикул WB
    order.techSize || '', // Размер
    order.finishedPrice || order.priceWithDisc || '', // Цена
    order.warehouseName || '', // Склад отправки
    order.regionName || '', // Регион доставки
    order.date || '', // Время заказа
    order.lastChangeDate || '', // Обновление
    1 // КОЛ_ВО (всегда 1, так как 1 строка = 1 заказ)
  ];
}

/**
 * Основная функция для выгрузки ленты заказов
 * Выгружает данные за предыдущий день с дозаписыванием
 */
function syncOrdersFeed() {
  try {
    Logger.log('=== Начало синхронизации ленты заказов ===');
    
    // Получаем дату предыдущего дня
    var reportDate = getPreviousDay();
    // Для flag=1 используем дату без времени
    var dateFrom = reportDate;
    
    Logger.log('Дата отчета: ' + reportDate);
    
    // Получаем лист
    var sheetName = getOrdersFeedSheetName();
    var sheet = getOrCreateSheet(sheetName);
    
    // Устанавливаем заголовки если лист пуст
    var headers = getOrdersFeedHeaders();
    setSheetHeaders(sheet, headers);
    
    // Загружаем данные из API
    // Используем flag=1 для получения всех заказов за дату
    Logger.log('Загрузка заказов из API Wildberries...');
    var orders = getOrders(dateFrom, 1);
    
    if (!orders || orders.length === 0) {
      Logger.log('Нет заказов для загрузки за ' + reportDate);
      return;
    }
    
    Logger.log('Получено заказов из API: ' + orders.length);
    
    // Получаем текущий номер последней позиции в таблице
    var lastRow = sheet.getLastRow();
    var startRowNumber = 1;
    
    // Если в таблице уже есть данные, находим максимальный номер позиции
    if (lastRow > 1) {
      var positionColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var maxPosition = 0;
      for (var j = 0; j < positionColumn.length; j++) {
        var posValue = positionColumn[j][0];
        if (typeof posValue === 'number' && posValue > maxPosition) {
          maxPosition = posValue;
        }
      }
      startRowNumber = maxPosition + 1;
    }
    
    // Преобразуем данные в формат таблицы
    var rows = [];
    for (var i = 0; i < orders.length; i++) {
      var rowNumber = startRowNumber + i;
      var row = convertOrderToRow(orders[i], rowNumber);
      rows.push(row);
    }
    
    // Записываем данные в таблицу
    Logger.log('Запись данных в таблицу...');
    appendDataToSheet(sheet, rows);
    
    Logger.log('=== Синхронизация завершена успешно. Добавлено заказов: ' + rows.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации ленты заказов: ' + error.toString());
    Logger.log('Стек ошибки: ' + error.stack);
    throw error;
  }
}
