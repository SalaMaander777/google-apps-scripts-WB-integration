/**
 * Модуль для ленты заказов
 * Выгружает данные за предыдущий день с дозаписыванием
 */

/**
 * Заголовки столбцов для листа "Лента заказов"
 */
function getOrdersFeedHeaders() {
  return [
    'Дата выгрузки', // Столбец A - дата выгрузки для перезаписи
    '№ позиции',
    'Баркод',
    'Наименование товара',
    'Артикул поставщика/цвет',
    'Артикул WB',
    'Размер',
    'Цена',
    'Склад отправки',
    'Регион доставки',
    'Дата заказа',
    'Время заказа',
    'Статус',
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
  
  // Разбиваем дату заказа на дату и время
  var orderDate = '';
  var orderTime = '';
  if (order.date) {
    var dateObj = new Date(order.date);
    if (!isNaN(dateObj.getTime())) {
      orderDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      orderTime = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'HH:mm:ss');
    }
  }
  
  // Форматируем дату и время обновления в одну строку
  var updateDateTime = '';
  if (order.lastChangeDate) {
    var updateDateObj = new Date(order.lastChangeDate);
    if (!isNaN(updateDateObj.getTime())) {
      updateDateTime = Utilities.formatDate(updateDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    }
  }
  
  // Определяем статус заказа
  var status = '';
  if (order.isCancel === true || order.cancel_dt) {
    status = 'отменен';
  } else {
    status = 'активен';
  }
  
  return [
    '', // Пустой первый столбец
    rowNumber, // № позиции
    order.barcode || '',
    productName || '', // Наименование товара
    order.supplierArticle || '', // Артикул поставщика/цвет
    order.nmId || '', // Артикул WB
    order.techSize || '', // Размер
    order.priceWithDisc || '', // Цена
    order.warehouseName || '', // Склад отправки
    order.regionName || '', // Регион доставки
    orderDate, // Дата заказа
    orderTime, // Время заказа
    status, // Статус
    updateDateTime, // Обновление (дата и время)
    '' // КОЛ_ВО - будет формулой
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
    
    // Инициализация структуры листа если он пуст
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      // Добавляем пустую строку в строке 1
      sheet.getRange(1, 1).setValue('');
      
      // Устанавливаем заголовки в строке 2
      var headers = getOrdersFeedHeaders();
      sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
      
      // Форматирование заголовков
      var headerRange = sheet.getRange(2, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e0e0e0');
      
      // Форматируем столбец K как дату для всех строк
      sheet.getRange(3, 11, sheet.getMaxRows() - 2, 1).setNumberFormat('M/d/yyyy');
      
      Logger.log('Создана структура листа с пустой строкой и заголовками');
      lastRow = 2;
    }
    
    // Загружаем данные из API
    // Используем flag=1 для получения всех заказов за дату
    Logger.log('Загрузка заказов из API Wildberries...');
    var orders = getOrders(dateFrom, 1);
    
    if (!orders || orders.length === 0) {
      Logger.log('Нет заказов для загрузки за ' + reportDate);
      return;
    }
    
    Logger.log('Получено заказов из API: ' + orders.length);
    
    // Перезаписываем строки за дату отчета по столбцу K (Дата заказа)
    // Данные начинаются со строки 3 (1 - пустая, 2 - заголовки)
    var deletedCount = deleteRowsByDate(sheet, reportDate, 11, 2);
    if (deletedCount > 0) {
      Logger.log('Удалено строк за дату ' + reportDate + ': ' + deletedCount);
    }
    lastRow = sheet.getLastRow();
    
    // Получаем текущий номер последней позиции в таблице
    var startRowNumber = 1;
    
    // Если в таблице уже есть данные, находим максимальный номер позиции
    // Данные начинаются со строки 3 (1 - пустая, 2 - заголовки)
    // Номер позиции находится во 2-м столбце (B)
    if (lastRow > 2) {
      var positionColumn = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
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
    var startRow = lastRow + 1;
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    
    // Форматируем столбец K (Дата заказа) как дату
    sheet.getRange(startRow, 11, rows.length, 1).setNumberFormat('M/d/yyyy');
    
    // Обновляем lastRow после записи
    lastRow = sheet.getLastRow();
    
    // Добавляем формулы в столбец КОЛ_ВО (15-й столбец, O) для ВСЕХ строк данных
    // Данные начинаются со строки 3 (1 - пустая, 2 - заголовки)
    // Статус находится в 13-м столбце (M)
    if (lastRow > 2) {
      Logger.log('Обновление формул в столбце КОЛ_ВО для всех строк данных...');
      var dataRowCount = lastRow - 2; // количество строк с данными
      var formulas = [];
      
      for (var i = 0; i < dataRowCount; i++) {
        var currentRow = 3 + i; // начинаем с 3-й строки
        formulas.push(['=IF(M' + currentRow + '="отменен",0,1)']);
      }
      
      // Устанавливаем все формулы за один раз для лучшей производительности
      sheet.getRange(3, 15, dataRowCount, 1).setFormulas(formulas);
      Logger.log('Формулы обновлены для ' + dataRowCount + ' строк');
    }
    
    Logger.log('=== Синхронизация завершена успешно. Добавлено заказов: ' + rows.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации ленты заказов: ' + error.toString());
    Logger.log('Стек ошибки: ' + error.stack);
    throw error;
  }
}
