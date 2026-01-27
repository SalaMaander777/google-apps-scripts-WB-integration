/**
 * Модуль для отчета "Аналитика РК"
 * Выгружает данные по рекламным кампаниям через API за предыдущий день с дозаписыванием
 */

/**
 * Основная функция для синхронизации аналитики рекламных кампаний
 * Получает список активных кампаний, записывает базовую информацию и статистику
 */
function syncAdsAnalytics() {
  try {
    Logger.log('=== Начало синхронизации аналитики РК ===');
    
    // 1. Получаем дату предыдущего дня
    var reportDate = getPreviousDay();
    Logger.log('Дата отчета: ' + reportDate);
    
    // 2. Получаем или создаем лист
    var sheetName = getAdsAnalyticsSheetName();
    var sheet = getOrCreateSheet(sheetName);
    
    // Инициализируем заголовки если лист пустой
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      var headers = getAdsAnalyticsHeaders();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      // Форматирование заголовков
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e0e0e0');
      Logger.log('Заголовки установлены');
    }
    
    // 3. Получаем список рекламных кампаний со статусами 7, 9, 11
    Logger.log('Получение списка рекламных кампаний...');
    var campaigns = getAdverts({
      statuses: '7,9,11'  // 7-завершена, 9-активна, 11-на паузе
    });
    
    if (!campaigns || campaigns.length === 0) {
      Logger.log('Нет кампаний со статусами 7, 9, 11');
      return;
    }
    
    Logger.log('Получено кампаний: ' + campaigns.length);
    
    // 4. Фильтруем кампании и собираем ID для статистики
    var campaignsToProcess = [];
    var campaignIds = [];
    
    for (var i = 0; i < campaigns.length; i++) {
      var campaign = campaigns[i];
      
      // Проверяем, что у кампании есть необходимые данные
      if (!campaign.id) {
        Logger.log('Пропускаем кампанию без ID');
        continue;
      }
      
      campaignsToProcess.push(campaign);
      campaignIds.push(campaign.id);
    }
    
    if (campaignsToProcess.length === 0) {
      Logger.log('Нет новых кампаний для обработки');
      return;
    }
    
    Logger.log('Кампаний для обработки: ' + campaignsToProcess.length);
    
    // 6. Получаем статистику по кампаниям за предыдущий день
    // API имеет лимит 3 запроса в минуту, поэтому обрабатываем батчами по 50
    var allStats = [];
    var batchSize = 50;
    
    for (var i = 0; i < campaignIds.length; i += batchSize) {
      var batch = campaignIds.slice(i, Math.min(i + batchSize, campaignIds.length));
      Logger.log('Запрос статистики для батча ' + (Math.floor(i / batchSize) + 1) + ', кампаний: ' + batch.length);
      
      try {
        var stats = getAdvertFullStats(batch, reportDate, reportDate);
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
    
    // 7. Создаем карту статистик по ID кампании для быстрого доступа
    var statsMap = {};
    for (var i = 0; i < allStats.length; i++) {
      var stat = allStats[i];
      if (stat.advertId) {
        statsMap[stat.advertId] = stat;
      }
    }
    
    // 8. Формируем данные для записи
    var dataToWrite = [];
    
    for (var i = 0; i < campaignsToProcess.length; i++) {
      var campaign = campaignsToProcess[i];
      var stats = statsMap[campaign.id] || {};
      
      // Пропускаем кампании с нулевыми показами
      if (!stats.views || stats.views === 0) {
        Logger.log('Пропускаем кампанию ' + campaign.id + ' - показы равны 0');
        continue;
      }
      
      // Получаем дату выгрузки из timestamps
      var uploadDate = reportDate;
      if (campaign.timestamps && campaign.timestamps.updatedAt) {
        // Извлекаем дату из timestamps.updatedAt
        var updatedAtDate = campaign.timestamps.updatedAt.split('T')[0];
        uploadDate = updatedAtDate;
      }
      
      // Извлекаем данные из вложенной структуры settings
      var campaignName = '';
      var paymentType = '';
      if (campaign.settings) {
        campaignName = campaign.settings.name || '';
        paymentType = campaign.settings.payment_type || '';
      }
      
      // Преобразуем тип ставки в русское название
      var bidTypeRu = '';
      if (campaign.bid_type === 'unified') {
        bidTypeRu = 'Единая ставка';
      } else if (campaign.bid_type === 'manual') {
        bidTypeRu = 'Ручная ставка';
      } else {
        bidTypeRu = campaign.bid_type || '';
      }
      
      // Форматируем период в формате DD.MM.YYYY - DD.MM.YYYY
      var periodFormatted = formatDateRu(reportDate) + ' - ' + formatDateRu(reportDate);
      
      var row = [
        uploadDate,                                      // A. Дата выгрузки
        bidTypeRu,                                       // B. Тип РК (Единая ставка/Ручная ставка)
        campaign.id || '',                               // C. ID РК
        campaignName,                                    // D. Название кампании (из settings.name)
        campaign.status || '',                           // E. Статус
        paymentType,                                     // F. Тип оплаты (из settings.payment_type)
        '',                                              // G. Старт (пустое поле)
        '',                                              // H. Финиш (формула будет добавлена отдельно)
        periodFormatted,                                 // I. Выбранный период
        stats.views || 0,                                // J. Показы
        stats.clicks || 0,                               // K. Клики
        ''                                               // L. АРТ (формула будет добавлена отдельно)
      ];
      
      dataToWrite.push(row);
    }
    
    if (dataToWrite.length === 0) {
      Logger.log('Нет данных для записи');
      return;
    }
    
    // Перезаписываем строки за дату отчета по столбцу A (Дата выгрузки)
    var deletedCount = deleteRowsByDate(sheet, reportDate, 1, 1);
    if (deletedCount > 0) {
      Logger.log('Удалено строк за дату ' + reportDate + ': ' + deletedCount);
    }
    
    // 9. Дозаписываем данные в таблицу
    var startRow = sheet.getLastRow() + 1;
    appendDataToSheet(sheet, dataToWrite);
    
    // 10. Добавляем формулы в столбцы H (Финиш) и L (АРТ)
    var endRow = sheet.getLastRow();
    for (var rowNum = startRow; rowNum <= endRow; rowNum++) {
      // Формула для столбца H (Финиш): =K/J
      sheet.getRange(rowNum, 8).setFormula('=K' + rowNum + '/J' + rowNum);
      
      // Формула для столбца L (АРТ): =ВПР(C,'ID-АРТ'!A:B,2,0)
      sheet.getRange(rowNum, 12).setFormula('=VLOOKUP(C' + rowNum + ',\'ID-АРТ\'!A:B,2,0)');
    }
    
    Logger.log('=== Синхронизация завершена успешно. Записано строк: ' + dataToWrite.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации аналитики РК: ' + error.toString());
    Logger.log('Стек ошибки: ' + error.stack);
    throw error;
  }
}

/**
 * Получить заголовки для листа аналитики РК
 * @return {Array<string>} Массив заголовков
 */
function getAdsAnalyticsHeaders() {
  return [
    'Дата выгрузки',
    'Тип РК',
    'ID РК',
    'Название',
    'Статус',
    'Тип оплаты',
    'Старт',
    'Финиш',
    'Выбранный период',
    'Показы',
    'Клики',
    'АРТ'
  ];
}

/**
 * Форматировать строку для записи в таблицу
 * @param {Object} campaign - Данные кампании
 * @param {Object} stats - Статистика кампании
 * @param {string} reportDate - Дата отчета в формате YYYY-MM-DD
 * @return {Array} Массив значений для строки
 */
function formatAdsAnalyticsRow(campaign, stats, reportDate) {
  // Получаем дату выгрузки
  var uploadDate = reportDate;
  if (campaign.timestamps && campaign.timestamps.updatedAt) {
    var updatedAtDate = campaign.timestamps.updatedAt.split('T')[0];
    uploadDate = updatedAtDate;
  }
  
  // Извлекаем данные из settings
  var campaignName = '';
  var paymentType = '';
  if (campaign.settings) {
    campaignName = campaign.settings.name || '';
    paymentType = campaign.settings.payment_type || '';
  }
  
  // Преобразуем тип ставки
  var bidTypeRu = '';
  if (campaign.bid_type === 'unified') {
    bidTypeRu = 'Единая ставка';
  } else if (campaign.bid_type === 'manual') {
    bidTypeRu = 'Ручная ставка';
  } else {
    bidTypeRu = campaign.bid_type || '';
  }
  
  // Форматируем период
  var periodFormatted = formatDateRu(reportDate) + ' - ' + formatDateRu(reportDate);
  
  var row = [
    uploadDate,
    bidTypeRu,
    campaign.advertId || campaign.id || '',
    campaignName,
    campaign.status || '',
    paymentType,
    '',
    '',
    periodFormatted,
    stats.views || 0,
    stats.clicks || 0,
    ''
  ];
  
  return row;
}

/**
 * Проверить, существует ли запись для кампании и даты в листе
 * @param {Sheet} sheet - Лист
 * @param {number} campaignId - ID кампании
 * @param {string} dateStr - Дата в формате YYYY-MM-DD
 * @return {boolean} true если запись уже существует
 */
function campaignAndDateExistInSheet(sheet, campaignId, dateStr) {
  if (isSheetEmpty(sheet)) {
    return false;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { // Только заголовок
    return false;
  }
  
  // Получаем столбцы: ID РК (столбец 3) и Выбранный период (столбец 9)
  var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  
  for (var i = 0; i < data.length; i++) {
    var rowCampaignId = data[i][2]; // ID РК в столбце 3
    var rowPeriod = data[i][8];      // Выбранный период в столбце 9
    
    // Извлекаем дату из периода (формат: DD.MM.YYYY - DD.MM.YYYY)
    var rowDateStr = '';
    if (typeof rowPeriod === 'string' && rowPeriod.indexOf(' - ') !== -1) {
      var periodStart = rowPeriod.split(' - ')[0];
      // Преобразуем DD.MM.YYYY в YYYY-MM-DD
      var parts = periodStart.split('.');
      if (parts.length === 3) {
        rowDateStr = parts[2] + '-' + parts[1] + '-' + parts[0];
      }
    }
    
    // Сравниваем ID кампании и дату
    if (String(rowCampaignId) === String(campaignId) && rowDateStr === dateStr) {
      return true;
    }
  }
  
  return false;
}
