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
      
      // Проверяем, есть ли уже данные за эту дату и кампанию
      if (campaignAndDateExistInSheet(sheet, campaign.id, reportDate)) {
        Logger.log('Данные для кампании ' + campaign.id + ' за ' + reportDate + ' уже существуют. Пропускаем.');
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
    
    // 5. Получаем статистику по кампаниям за предыдущий день
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
    
    // 6. Создаем карту статистик по ID кампании для быстрого доступа
    var statsMap = {};
    for (var i = 0; i < allStats.length; i++) {
      var stat = allStats[i];
      if (stat.advertId) {
        statsMap[stat.advertId] = stat;
      }
    }
    
    // 7. Формируем данные для записи
    var dataToWrite = [];
    
    for (var i = 0; i < campaignsToProcess.length; i++) {
      var campaign = campaignsToProcess[i];
      var stats = statsMap[campaign.id] || {};
      
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
      
      var row = [
        uploadDate,                                      // 1. Дата выгрузки
        campaign.bid_type || '',                         // 2. Тип РК (unified/manual)
        campaign.id || '',                               // 3. ID РК
        campaignName,                                    // 4. Название кампании (из settings.name)
        campaign.status || '',                           // 5. Статус
        paymentType,                                     // 6. Тип оплаты (из settings.payment_type)
        reportDate,                                      // 7. Дата начала периода статистики
        reportDate,                                      // 8. Дата конца периода статистики
        stats.views || 0,                                // 9. Показы
        stats.clicks || 0,                               // 10. Клики
        stats.ctr || 0,                                  // 11. CTR
        stats.cpc || 0,                                  // 12. CPC
        stats.sum || 0,                                  // 13. Затраты
        stats.atbs || 0,                                 // 14. Добавления в корзину
        stats.orders || 0,                               // 15. Заказы
        stats.cr || 0,                                   // 16. CR (conversion rate)
        stats.shks || 0,                                 // 17. Количество заказанных товаров
        stats.sum_price || 0,                            // 18. Сумма заказов
        stats.canceled || 0                              // 19. Отмены
      ];
      
      dataToWrite.push(row);
    }
    
    if (dataToWrite.length === 0) {
      Logger.log('Нет данных для записи');
      return;
    }
    
    // 8. Дозаписываем данные в таблицу
    appendDataToSheet(sheet, dataToWrite);
    
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
    'Период начало',
    'Период конец',
    'Показы',
    'Клики',
    'CTR',
    'CPC',
    'Затраты',
    'Добавления в корзину',
    'Заказы',
    'CR',
    'Заказано товаров шт',
    'Сумма заказов',
    'Отмены'
  ];
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
  
  // Получаем столбцы: ID РК (столбец 3) и Период начало (столбец 7)
  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  for (var i = 0; i < data.length; i++) {
    var rowCampaignId = data[i][2]; // ID РК в столбце 3
    var rowDate = data[i][6];        // Период начало в столбце 7
    
    // Преобразуем дату в строку для сравнения
    var rowDateStr = '';
    if (rowDate instanceof Date) {
      rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof rowDate === 'string') {
      rowDateStr = rowDate.split('T')[0];
    }
    
    // Сравниваем ID кампании и дату
    if (String(rowCampaignId) === String(campaignId) && rowDateStr === dateStr) {
      return true;
    }
  }
  
  return false;
}
