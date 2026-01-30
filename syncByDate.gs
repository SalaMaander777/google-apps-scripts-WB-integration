/**
 * Модуль для синхронизации отчетов за конкретную дату
 * Обертки над основными функциями синхронизации с передачей даты
 */

/**
 * Синхронизация остатков товаров за дату
 * @param {string} date - Дата в формате YYYY-MM-DD
 */
function syncStocksByDate(date) {
  Logger.log('syncStocksByDate: остатки не зависят от даты, выполняется обычная синхронизация');
  syncStocks();
}

/**
 * Синхронизация финансовых отчетов за дату
 * @param {string} date - Дата в формате YYYY-MM-DD
 */
function syncFinanceDailyReportByDate(date) {
  try {
    Logger.log('=== Начало синхронизации финансовых отчетов за дату: ' + date + ' ===');
    
    // Получаем или создаем лист
    var sheetName = getFinanceDailySheetName();
    var sheet = getOrCreateSheet(sheetName);
    
    // Устанавливаем заголовки если лист пустой
    if (isSheetEmpty(sheet)) {
      var headers = getFinanceReportHeaders();
      setSheetHeaders(sheet, headers);
    }
    
    // Проверяем, есть ли уже данные за эту дату
    // Получаем данные из API
    var dateRange = {
      dateFrom: date + 'T00:00:00+03:00',
      dateTo: date + 'T23:59:59+03:00'
    };
    
    var records = getReportDetailByPeriod(dateRange.dateFrom, dateRange.dateTo, 'daily');
    
    if (!records || records.length === 0) {
      Logger.log('Нет данных за ' + date);
      return;
    }
    
    // Форматируем данные для таблицы  
    var data = formatFinanceData(records, date);
    
    if (!data || data.length === 0) {
      Logger.log('Нет данных за ' + date);
      return;
    }
    
    // Перезаписываем строки за дату выгрузки по последнему столбцу
    var headers = getFinanceReportHeaders();
    var deletedCount = deleteRowsByDate(sheet, date, headers.length, 1);
    if (deletedCount > 0) {
      Logger.log('Удалено строк за дату ' + date + ': ' + deletedCount);
    }
    
    // Дозаписываем данные
    appendDataToSheet(sheet, data);
    
    Logger.log('=== Синхронизация финансовых отчетов завершена. Записано строк: ' + data.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации финансовых отчетов: ' + error.toString());
    throw error;
  }
}

/**
 * Синхронизация ленты заказов за дату
 * @param {string} date - Дата в формате YYYY-MM-DD
 */
function syncOrdersFeedByDate(date) {
  try {
    Logger.log('=== Начало синхронизации ленты заказов за дату: ' + date + ' ===');
    
    // Получаем или создаем лист
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
    
    // Получаем данные из API
    Logger.log('Загрузка заказов из API Wildberries...');
    var orders = getOrders(date, 1);
    
    if (!orders || orders.length === 0) {
      Logger.log('Нет заказов за ' + date);
      return;
    }
    
    Logger.log('Получено заказов из API: ' + orders.length);
    
    // Перезаписываем строки за указанную дату по столбцу K (Дата заказа)
    // Данные начинаются со строки 3 (1 - пустая, 2 - заголовки)
    var deletedCount = deleteRowsByDate(sheet, date, 11, 2);
    if (deletedCount > 0) {
      Logger.log('Удалено строк за дату ' + date + ': ' + deletedCount);
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
    var data = [];
    for (var i = 0; i < orders.length; i++) {
      var rowNumber = startRowNumber + i;
      var row = convertOrderToRow(orders[i], rowNumber);
      data.push(row);
    }
    
    if (!data || data.length === 0) {
      Logger.log('Нет заказов за ' + date);
      return;
    }
    
    // Записываем данные в таблицу
    Logger.log('Запись данных в таблицу...');
    var startRow = lastRow + 1;
    sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
    
    // Форматируем столбец K (Дата заказа) как дату
    sheet.getRange(startRow, 11, data.length, 1).setNumberFormat('M/d/yyyy');
    
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
    
    Logger.log('=== Синхронизация ленты заказов завершена. Добавлено заказов: ' + data.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации ленты заказов: ' + error.toString());
    Logger.log('Стек ошибки: ' + error.stack);
    throw error;
  }
}

/**
 * Синхронизация аналитики РК за дату
 * @param {string} date - Дата в формате YYYY-MM-DD
 */
function syncAdsAnalyticsByDate(date) {
  try {
    Logger.log('=== Начало синхронизации аналитики РК за дату: ' + date + ' ===');
    
    // Получаем или создаем лист
    var sheetName = getAdsAnalyticsSheetName();
    var sheet = getOrCreateSheet(sheetName);
    
    // Инициализируем заголовки если лист пустой
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      var headers = getAdsAnalyticsHeaders();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e0e0e0');
      Logger.log('Заголовки установлены');
    }
    
    // Получаем список всех рекламных кампаний
    Logger.log('Получение списка рекламных кампаний...');
    var campaigns = getAdverts({
      statuses: '7,9,11'  // 7-завершена, 9-активна, 11-на паузе
    });
    
    if (!campaigns || campaigns.length === 0) {
      Logger.log('Нет активных рекламных кампаний');
      return;
    }
    
    Logger.log('Найдено кампаний: ' + campaigns.length);
    
    // Фильтруем кампании, для которых нужно получить данные
    var campaignsToProcess = [];
    var campaignIds = [];
    
    for (var i = 0; i < campaigns.length; i++) {
      var campaign = campaigns[i];
      
      // Проверяем, что у кампании есть необходимые данные
      if (!campaign.id && !campaign.advertId) {
        Logger.log('Пропускаем кампанию без ID');
        continue;
      }
      
      var campaignId = campaign.id || campaign.advertId;
      
      campaignsToProcess.push(campaign);
      campaignIds.push(campaignId);
    }
    
    if (campaignsToProcess.length === 0) {
      Logger.log('Нет новых кампаний для обработки');
      return;
    }
    
    Logger.log('Кампаний для обработки: ' + campaignsToProcess.length);
    
    // Получаем статистику по кампаниям батчами (по 50 кампаний за раз)
    // API имеет лимит 3 запроса в минуту, поэтому между батчами делаем паузу 20 секунд
    var allStats = [];
    var batchSize = 50;
    
    for (var i = 0; i < campaignIds.length; i += batchSize) {
      var batch = campaignIds.slice(i, Math.min(i + batchSize, campaignIds.length));
      Logger.log('Запрос статистики для батча ' + (Math.floor(i / batchSize) + 1) + ', кампаний: ' + batch.length);
      
      try {
        var stats = getAdvertFullStats(batch, date, date);
        if (stats && stats.length > 0) {
          allStats = allStats.concat(stats);
          Logger.log('Получено статистик в батче: ' + stats.length);
        }
      } catch (error) {
        Logger.log('Ошибка получения статистики для батча: ' + error.toString());
        // Продолжаем обработку следующего батча
      }
      
      // Ждем 20 секунд между запросами для соблюдения лимита API (3 запроса в минуту)
      if (i + batchSize < campaignIds.length) {
        Logger.log('Ожидание 20 секунд перед следующим батчем...');
        Utilities.sleep(20000);
      }
    }
    
    Logger.log('Всего получено статистик: ' + allStats.length);
    
    // Создаем карту статистик по ID кампании для быстрого доступа
    var statsMap = {};
    for (var i = 0; i < allStats.length; i++) {
      var stat = allStats[i];
      if (stat.advertId) {
        statsMap[stat.advertId] = stat;
      }
    }
    
    // Формируем данные для записи
    var dataToWrite = [];
    
    for (var i = 0; i < campaignsToProcess.length; i++) {
      var campaign = campaignsToProcess[i];
      var campaignId = campaign.id || campaign.advertId;
      var stats = statsMap[campaignId] || {};
      
      // Пропускаем кампании с нулевыми показами
      if (!stats.views || stats.views === 0) {
        Logger.log('Пропускаем кампанию ' + campaignId + ' - показы равны 0 или нет статистики');
        continue;
      }
      
      var row = formatAdsAnalyticsRow(campaign, stats, date);
      dataToWrite.push(row);
    }
    
    if (dataToWrite.length === 0) {
      Logger.log('Нет новых данных для записи');
      return;
    }
    
    // Перезаписываем строки за указанную дату по столбцу A (Дата выгрузки)
    var deletedCount = deleteRowsByDate(sheet, date, 1, 1);
    if (deletedCount > 0) {
      Logger.log('Удалено строк за дату ' + date + ': ' + deletedCount);
    }
    
    // Записываем данные
    var startRow = sheet.getLastRow() + 1;
    appendDataToSheet(sheet, dataToWrite);
    
    // Добавляем формулы в столбцы H (Финиш) и L (АРТ)
    var endRow = sheet.getLastRow();
    for (var rowNum = startRow; rowNum <= endRow; rowNum++) {
      // Формула для столбца H (Финиш): =K/J
      sheet.getRange(rowNum, 8).setFormula('=K' + rowNum + '/J' + rowNum);
      
      // Формула для столбца L (АРТ): =ВПР(C,'ID-АРТ'!A:B,2,0)
      sheet.getRange(rowNum, 12).setFormula('=VLOOKUP(C' + rowNum + ',\'ID-АРТ\'!A:B,2,0)');
    }
    
    Logger.log('=== Синхронизация аналитики РК завершена. Записано строк: ' + dataToWrite.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации аналитики РК: ' + error.toString());
    throw error;
  }
}

/**
 * Синхронизация истории рекламных расходов за дату
 * @param {string} date - Дата в формате YYYY-MM-DD
 */
function syncAdsCostsByDate(date) {
  try {
    Logger.log('=== Начало синхронизации истории рекламных расходов за дату: ' + date + ' ===');
    
    // Получаем или создаем лист
    var sheetName = getAdsCostsSheetName();
    var sheet = getOrCreateSheet(sheetName);
    
    // Инициализируем заголовки если лист пустой
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      var headers = getAdsCostsHeaders();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e0e0e0');
      Logger.log('Заголовки установлены');
    }
    
    // Перезаписываем строки за указанную дату по столбцу D (Дата)
    var deletedCount = deleteRowsByDate(sheet, date, 4, 1);
    if (deletedCount > 0) {
      Logger.log('Удалено строк за дату ' + date + ': ' + deletedCount);
    }
    
    // Получаем данные из API
    Logger.log('Получение данных из API...');
    var costsDataRaw = getAdvertCostsHistory(date, date);
    
    if (!costsDataRaw || costsDataRaw.length === 0) {
      Logger.log('Нет данных за ' + date);
      return;
    }
    
    Logger.log('Получено записей затрат: ' + costsDataRaw.length);
    
    // Получаем информацию о кампаниях
    var campaignIds = [];
    var uniqueIds = {};
    for (var i = 0; i < costsDataRaw.length; i++) {
      var advertId = costsDataRaw[i].advertId;
      if (advertId && !uniqueIds[advertId]) {
        uniqueIds[advertId] = true;
        campaignIds.push(advertId);
      }
    }
    
    var campaignInfoMap = getCampaignInfoMap(campaignIds);
    
    // Номер последней строки перед записью (для подстановки в формулы)
    lastRow = sheet.getLastRow();
    
    // Форматируем данные
    var costsData = [];
    for (var i = 0; i < costsDataRaw.length; i++) {
      var cost = costsDataRaw[i];
      
      var updTimeFormatted = '';
      if (cost.updTime) {
        try {
          var updDate = new Date(cost.updTime);
          updTimeFormatted = Utilities.formatDate(updDate, Session.getScriptTimeZone(), 'HH:mm:ss');
        } catch (e) {
          updTimeFormatted = cost.updTime;
        }
      }
      
      var campaignInfo = campaignInfoMap[cost.advertId] || {};
      var bidType = '';
      if (campaignInfo.bid_type === 'unified') {
        bidType = 'Единая ставка';
      } else if (campaignInfo.bid_type === 'manual') {
        bidType = 'Ручная ставка';
      } else {
        bidType = campaignInfo.bid_type || '';
      }
      
      // Формула VLOOKUP по ID кампании (столбец A)
      var formulaRow = lastRow + 1 + i;
      var vlookupFormula = "=VLOOKUP(A" + formulaRow + ",'ID-АРТ'!A:B,2,0)";
      
      var row = [
        cost.advertId || '',
        cost.campName || '',
        bidType,
        date,
        updTimeFormatted,
        cost.paymentType || '',
        (cost.updSum || 0) / 100,
        cost.updNum || 0,              // Номер документа (если пустой — 0)
        vlookupFormula                 // Артикул из листа ID-АРТ
      ];
      
      costsData.push(row);
    }
    
    if (!costsData || costsData.length === 0) {
      Logger.log('Нет данных за ' + date);
      return;
    }
    
    Logger.log('Получено записей затрат: ' + costsData.length);
    
    // Дозаписываем данные
    appendDataToSheet(sheet, costsData);
    
    Logger.log('=== Синхронизация истории рекламных расходов завершена. Записано строк: ' + costsData.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации истории рекламных расходов: ' + error.toString());
    throw error;
  }
}

/**
 * Синхронизация аналитики продавца за дату
 * @param {string} date - Дата в формате YYYY-MM-DD
 */
function syncSalesFunnelByDate(date) {
  try {
    Logger.log('=== Начало синхронизации аналитики продавца за дату: ' + date + ' ===');
    
    // Получаем или создаем лист
    var sheetName = getSalesFunnelSheetName();
    var sheet = getOrCreateSheet(sheetName);
    
    // Инициализируем заголовки если лист пустой
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      var headers = getSalesFunnelHeaders();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e0e0e0');
      Logger.log('Заголовки установлены');
    }
    
    // Получаем данные из API БЕЗ периода сравнения (только selectedPeriod)
    Logger.log('Получение данных из API за ' + date + ' (без периода сравнения)...');
    var allProducts = getSalesFunnelData(date, date, null, null);
    
    if (!allProducts || allProducts.length === 0) {
      Logger.log('Нет данных за ' + date);
      return;
    }
    
    Logger.log('Получено товаров: ' + allProducts.length);
    
    // Перезаписываем строки за указанную дату по столбцу A (Дата)
    var deletedCount = deleteRowsByDate(sheet, date, 1, 1);
    if (deletedCount > 0) {
      Logger.log('Удалено строк за дату ' + date + ': ' + deletedCount);
    }
    
    // Фильтруем товары, для которых уже есть данные
    var productsToProcess = [];
    
    for (var i = 0; i < allProducts.length; i++) {
      var item = allProducts[i];
      
      if (!item.product || !item.product.nmId) {
        Logger.log('Пропускаем элемент без product.nmId');
        continue;
      }
      
      productsToProcess.push(item);
    }
    
    if (productsToProcess.length === 0) {
      Logger.log('Нет новых товаров для обработки');
      return;
    }
    
    Logger.log('Товаров для обработки: ' + productsToProcess.length);
    
    // Формируем данные для записи
    var dataToWrite = [];
    
    for (var i = 0; i < productsToProcess.length; i++) {
      var item = productsToProcess[i];
      var product = item.product || {};
      var statistic = item.statistic || {};
      var selected = statistic.selected || {};
      var past = statistic.past || {};
      var comparison = statistic.comparison || {};
      var stocks = product.stocks || {};
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
        date,                                    // 1. Дата
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
    
    // Дозаписываем данные
    appendDataToSheet(sheet, dataToWrite);
    
    Logger.log('=== Синхронизация аналитики продавца завершена. Записано строк: ' + dataToWrite.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации аналитики продавца: ' + error.toString());
    throw error;
  }
}

/**
 * Синхронизация воронки динамики за неделю
 * Добавляет дневные столбцы за всю неделю и итоговый недельный столбец
 * @param {string} date - Дата (любой день недели или воскресенье для конца недели) в формате YYYY-MM-DD
 */
function syncFunnelDynamicWeekByDate(date) {
  try {
    Logger.log('=== Начало синхронизации недельной воронки динамики ===');
    Logger.log('Входная дата: ' + date);
    
    // Вызываем функцию синхронизации недельной статистики
    // Она сама найдет ближайшее воскресенье и обработает всю неделю
    var result = syncSalesFunnelDynamicWeek(date);
    
    Logger.log('=== Синхронизация недельной воронки динамики завершена ===');
    
    return result;
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации недельной воронки динамики: ' + error.toString());
    throw error;
  }
}
