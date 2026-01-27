/**
 * Модуль для отчета "Статистика карточек товаров" (Sales Funnel)
 * Выгружает данные по воронке продаж через API за предыдущий день с дозаписыванием
 */

/**
 * Основная функция для синхронизации статистики карточек товаров
 * Получает данные за предыдущий день с сравнением с таким же днем год назад
 */
function syncSalesFunnel() {
  try {
    Logger.log('=== Начало синхронизации статистики карточек товаров ===');
    
    // 1. Получаем дату предыдущего дня
    var reportDate = getPreviousDay();
    Logger.log('Дата отчета: ' + reportDate);
    
    // 2. Получаем дату год назад для сравнения
    var pastDate = getDateYearAgo(reportDate);
    Logger.log('Дата для сравнения: ' + pastDate);
    
    // 3. Получаем или создаем лист
    var sheetName = getSalesFunnelSheetName();
    var sheet = getOrCreateSheet(sheetName);
    
    // Инициализируем заголовки если лист пустой
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      var headers = getSalesFunnelHeaders();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      // Форматирование заголовков
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e0e0e0');
      Logger.log('Заголовки установлены');
    }
    
    // 4. Получаем данные из API с пагинацией
    Logger.log('Получение данных из API...');
    var allProducts = getSalesFunnelData(reportDate, reportDate, pastDate, pastDate);
    
    if (!allProducts || allProducts.length === 0) {
      Logger.log('Нет данных за ' + reportDate);
      return;
    }
    
    Logger.log('Получено товаров: ' + allProducts.length);
    
    // 5. Фильтруем товары для обработки
    var productsToProcess = [];
    
    for (var i = 0; i < allProducts.length; i++) {
      var item = allProducts[i];
      
      // Проверяем структуру данных
      if (!item.product || !item.product.nmId) {
        Logger.log('Пропускаем элемент без product.nmId');
        continue;
      }
      
      productsToProcess.push(item);
    }
    
    if (productsToProcess.length === 0) {
      Logger.log('Нет товаров для обработки');
      return;
    }
    
    Logger.log('Товаров для обработки: ' + productsToProcess.length);
    
    // 7. Формируем данные для записи
    var dataToWrite = [];
    
    for (var i = 0; i < productsToProcess.length; i++) {
      var item = productsToProcess[i];
      var product = item.product || {};
      var statistic = item.statistic || {};
      var selected = statistic.selected || {};
      var past = statistic.past || {};
      var comparison = statistic.comparison || {};
      var stocks = product.stocks || {};
      var selectedWbClub = selected.wbClub || {};
      var pastWbClub = past.wbClub || {};
      var selectedConversions = selected.conversions || {};
      var pastConversions = past.conversions || {};
      var selectedTime = selected.timeToReady || {};
      var pastTime = past.timeToReady || {};
      
      // Формируем строку среднего времени доставки
      var selectedTimeStr = '';
      if (selectedTime.days || selectedTime.hours || selectedTime.mins) {
        selectedTimeStr = (selectedTime.days || 0) + 'д ' + (selectedTime.hours || 0) + 'ч ' + (selectedTime.mins || 0) + 'м';
      }
      
      var pastTimeStr = '';
      if (pastTime.days || pastTime.hours || pastTime.mins) {
        pastTimeStr = (pastTime.days || 0) + 'д ' + (pastTime.hours || 0) + 'ч ' + (pastTime.mins || 0) + 'м';
      }
      
      var row = [
        reportDate,                              // 1. Дата
        product.vendorCode || '',                // 2. Артикул продавца
        product.nmId || '',                      // 3. Номенклатура
        product.title || '',                     // 4. Название
        product.subjectName || '',               // 5. Категория
        product.brandName || '',                 // 6. Бренд
        '',                                      // 7. Удаленный товар (нет в API)
        product.productRating || 0,              // 8. Рейтинг карточки
        selected.openCount || 0,                 // 9. Переходы в карточку
        past.openCount || 0,                     // 10. Переходы в карточку (предыдущий период)
        selected.cartCount || 0,                 // 11. Положили в корзину
        past.cartCount || 0,                     // 12. Положили в корзину (предыдущий период)
        selected.orderCount || 0,                // 13. Заказали, шт
        past.orderCount || 0,                    // 14. Заказали, шт (предыдущий период)
        selected.buyoutCount || 0,               // 15. Выкупили, шт
        past.buyoutCount || 0,                   // 16. Выкупы, шт (предыдущий период)
        selected.cancelCount || 0,               // 17. Отменили, шт
        past.cancelCount || 0,                   // 18. Отменили, шт (предыдущий период)
        selectedConversions.addToCartPercent || 0,   // 19. Конверсия в корзину, %
        pastConversions.addToCartPercent || 0,       // 20. Конверсия в корзину, % (предыдущий период)
        selectedConversions.cartToOrderPercent || 0, // 21. Конверсия в заказ, %
        pastConversions.cartToOrderPercent || 0,     // 22. Конверсия в заказ, % (предыдущий период)
        selectedConversions.buyoutPercent || 0,      // 23. Процент выкупа
        pastConversions.buyoutPercent || 0,          // 24. Процент выкупа (предыдущий период)
        selected.orderSum || 0,                  // 25. Заказали на сумму, руб
        past.orderSum || 0,                      // 26. Заказали на сумму, руб (предыдущий период)
        comparison.orderSumDynamic || 0,         // 27. Динамика суммы заказов, руб
        selected.buyoutSum || 0,                 // 28. Выкупили на сумму, руб
        past.buyoutSum || 0,                     // 29. Выкупили на сумму, руб (предыдущий период)
        selected.cancelSum || 0,                 // 30. Отменили на сумму, руб
        past.cancelSum || 0,                     // 31. Отменили на сумму, руб (предыдущий период)
        selected.avgPrice || 0,                  // 32. Средняя цена, руб
        past.avgPrice || 0,                      // 33. Средняя цена, руб (предыдущий период)
        selected.avgOrdersCountPerDay || 0,      // 34. Среднее количество заказов в день, шт
        past.avgOrdersCountPerDay || 0,          // 35. Среднее количество заказов в день, шт (предыдущий период)
        stocks.wb || 0,                          // 36. Остатки склад ВБ, шт
        stocks.mp || 0,                          // 37. Остатки МП, шт
        stocks.balanceSum || 0,                  // 38. Сумма остатков на складах, руб
        selectedTimeStr,                         // 39. Среднее время доставки
        pastTimeStr,                             // 40. Среднее время доставки (предыдущий период)
        selected.localizationPercent || 0,       // 41. Локальные заказы, %
        past.localizationPercent || 0            // 42. Локальные заказы, % (предыдущий период)
      ];
      
      dataToWrite.push(row);
    }
    
    if (dataToWrite.length === 0) {
      Logger.log('Нет данных для записи');
      return;
    }
    
    // Перезаписываем строки за дату отчета по столбцу A (Дата)
    var deletedCount = deleteRowsByDate(sheet, reportDate, 1, 1);
    if (deletedCount > 0) {
      Logger.log('Удалено строк за дату ' + reportDate + ': ' + deletedCount);
    }
    
    // 8. Дозаписываем данные в таблицу
    appendDataToSheet(sheet, dataToWrite);
    
    Logger.log('=== Синхронизация завершена успешно. Записано строк: ' + dataToWrite.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации статистики карточек товаров: ' + error.toString());
    Logger.log('Стек ошибки: ' + error.stack);
    throw error;
  }
}

/**
 * Получить заголовки для листа статистики карточек товаров
 * @return {Array<string>} Массив заголовков
 */
function getSalesFunnelHeaders() {
  return [
    'Дата',
    'Артикул продавца',
    'Номенклатура',
    'Название',
    'Категория',
    'Бренд',
    'Удаленный товар',
    'Рейтинг карточки',
    'Переходы в карточку',
    'Переходы в карточку (предыдущий период)',
    'Положили в корзину',
    'Положили в корзину (предыдущий период)',
    'Заказали, шт',
    'Заказали, шт (предыдущий период)',
    'Выкупили, шт',
    'Выкупы, шт (предыдущий период)',
    'Отменили, шт',
    'Отменили, шт (предыдущий период)',
    'Конверсия в корзину, %',
    'Конверсия в корзину, % (предыдущий период)',
    'Конверсия в заказ, %',
    'Конверсия в заказ, % (предыдущий период)',
    'Процент выкупа',
    'Процент выкупа (предыдущий период)',
    'Заказали на сумму, руб',
    'Заказали на сумму, руб (предыдущий период)',
    'Динамика суммы заказов, руб',
    'Выкупили на сумму, руб',
    'Выкупили на сумму, руб (предыдущий период)',
    'Отменили на сумму, руб',
    'Отменили на сумму, руб (предыдущий период)',
    'Средняя цена, руб',
    'Средняя цена, руб (предыдущий период)',
    'Среднее количество заказов в день, шт',
    'Среднее количество заказов в день, шт (предыдущий период)',
    'Остатки склад ВБ, шт',
    'Остатки МП, шт',
    'Сумма остатков на складах, руб',
    'Среднее время доставки',
    'Среднее время доставки (предыдущий период)',
    'Локальные заказы, %',
    'Локальные заказы, % (предыдущий период)'
  ];
}

/**
 * Проверить, существует ли запись для товара и даты в листе
 * @param {Sheet} sheet - Лист
 * @param {number} nmId - Артикул WB
 * @param {string} dateStr - Дата в формате YYYY-MM-DD
 * @return {boolean} true если запись уже существует
 */
function productAndDateExistInSheet(sheet, nmId, dateStr) {
  if (isSheetEmpty(sheet)) {
    return false;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { // Только заголовок
    return false;
  }
  
  // Получаем столбцы: Дата отчета (столбец 1) и Артикул WB (столбец 2)
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  
  for (var i = 0; i < data.length; i++) {
    var rowDate = data[i][0];        // Дата отчета в столбце 1
    var rowNmId = data[i][1];        // Артикул WB в столбце 2
    
    // Преобразуем дату в строку для сравнения
    var rowDateStr = '';
    if (rowDate instanceof Date) {
      rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof rowDate === 'string') {
      rowDateStr = rowDate.split('T')[0];
    }
    
    // Сравниваем артикул и дату
    if (String(rowNmId) === String(nmId) && rowDateStr === dateStr) {
      return true;
    }
  }
  
  return false;
}
