/**
 * Модуль для отчета "История рекламных расходов"
 * Выгружает данные о фактических затратах на рекламные кампании за предыдущий день
 */

/**
 * Основная функция для синхронизации истории рекламных расходов
 * Получает список затрат за предыдущий день и дозаписывает в таблицу
 */
function syncAdsCosts() {
  try {
    Logger.log('=== Начало синхронизации истории рекламных расходов ===');
    
    // 1. Получаем дату предыдущего дня
    var reportDate = getPreviousDay();
    Logger.log('Дата отчета: ' + reportDate);
    
    // 2. Получаем или создаем лист
    var sheetName = getAdsCostsSheetName();
    var sheet = getOrCreateSheet(sheetName);
    
    // Инициализируем заголовки если лист пустой
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      var headers = getAdsCostsHeaders();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      // Форматирование заголовков
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e0e0e0');
      Logger.log('Заголовки установлены');
    }
    
    // Перезаписываем строки за дату отчета (столбец D — Дата)
    var deletedCount = deleteRowsByDate(sheet, reportDate, 4, 1);
    if (deletedCount > 0) {
      Logger.log('Удалено строк за дату ' + reportDate + ': ' + deletedCount);
    }
    lastRow = sheet.getLastRow();
    
    // 3. Получаем историю затрат за предыдущий день
    // API требует период минимум 1 день, поэтому from = to = reportDate
    Logger.log('Запрос истории затрат за период: ' + reportDate + ' - ' + reportDate);
    var costsData = getAdvertCostsHistory(reportDate, reportDate);
    
    if (!costsData || costsData.length === 0) {
      Logger.log('Нет данных о затратах за ' + reportDate);
      return;
    }
    
    Logger.log('Получено записей затрат: ' + costsData.length);
    
    // 5. Получаем информацию о кампаниях для получения bid_type
    // Собираем уникальные ID кампаний из затрат
    var campaignIds = [];
    var uniqueIds = {};
    for (var i = 0; i < costsData.length; i++) {
      var advertId = costsData[i].advertId;
      if (advertId && !uniqueIds[advertId]) {
        uniqueIds[advertId] = true;
        campaignIds.push(advertId);
      }
    }
    
    Logger.log('Уникальных кампаний: ' + campaignIds.length);
    
    // Получаем информацию о кампаниях
    var campaignInfoMap = getCampaignInfoMap(campaignIds);
    
    // 6. Формируем данные для записи
    var dataToWrite = [];
    
    for (var i = 0; i < costsData.length; i++) {
      var cost = costsData[i];
      
      // Обрабатываем дату и время списания
      var updTimeFormatted = '';
      if (cost.updTime) {
        try {
          // updTime приходит в формате ISO: "2023-08-01T12:34:56Z"
          var updDate = new Date(cost.updTime);
          updTimeFormatted = Utilities.formatDate(updDate, Session.getScriptTimeZone(), 'HH:mm:ss');
        } catch (e) {
          Logger.log('Ошибка парсинга updTime: ' + cost.updTime + ', ошибка: ' + e.toString());
          updTimeFormatted = cost.updTime;
        }
      }
      
      // Получаем bid_type из карты кампаний (перевод на русский как в Аналитике РК)
      var campaignInfo = campaignInfoMap[cost.advertId] || {};
      var bidType = '';
      if (campaignInfo.bid_type === 'unified') {
        bidType = 'Единая ставка';
      } else if (campaignInfo.bid_type === 'manual') {
        bidType = 'Ручная ставка';
      } else {
        bidType = campaignInfo.bid_type || '';
      }
      
      // Номер строки для формулы VLOOKUP (столбец A на листе)
      var formulaRow = lastRow + 1 + i;
      var vlookupFormula = "=VLOOKUP(A" + formulaRow + ",'ID-АРТ'!A:B,2,0)";

      // Порядок столбцов: ID кампании, Название, Раздел (bid_type), Дата, Списания, Источник списания, Сумма, Номер документа, Артикул (VLOOKUP)
      var row = [
        cost.advertId || '',           // 1. ID кампании
        cost.campName || '',           // 2. Название кампании
        bidType,                       // 3. Раздел (bid_type)
        reportDate,                    // 4. Дата
        updTimeFormatted,              // 5. Списания (время списания)
        cost.paymentType || '',        // 6. Источник списания
        (cost.updSum || 0) / 100,      // 7. Сумма (в рублях)
        cost.updNum || 0,              // 8. Номер документа (если пустой — 0)
        vlookupFormula                 // 9. Артикул из листа ID-АРТ
      ];

      dataToWrite.push(row);
    }
    
    if (dataToWrite.length === 0) {
      Logger.log('Нет данных для записи');
      return;
    }
    
    // 7. Дозаписываем данные в таблицу
    appendDataToSheet(sheet, dataToWrite);
    
    Logger.log('=== Синхронизация завершена успешно. Записано строк: ' + dataToWrite.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации истории рекламных расходов: ' + error.toString());
    Logger.log('Стек ошибки: ' + error.stack);
    throw error;
  }
}

/**
 * Получить заголовки для листа истории рекламных расходов
 * @return {Array<string>} Массив заголовков
 */
function getAdsCostsHeaders() {
  return [
    'ID кампании',
    'Название кампании',
    'Раздел (bid_type)',
    'Дата',
    'Списания',
    'Источник списания',
    'Сумма',
    'Номер документа',
    'Артикул'
  ];
}

/**
 * Проверить, существуют ли данные за указанную дату в листе
 * @param {Sheet} sheet - Лист
 * @param {string} dateStr - Дата в формате YYYY-MM-DD
 * @return {boolean} true если данные уже существуют
 */
function dateExistsInAdsCostsSheet(sheet, dateStr) {
  if (isSheetEmpty(sheet)) {
    return false;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { // Только заголовок
    return false;
  }
  
  // Получаем столбец с датой (столбец 4 - "Дата")
  var data = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  
  for (var i = 0; i < data.length; i++) {
    var rowDate = data[i][0];
    
    // Преобразуем дату в строку для сравнения
    var rowDateStr = '';
    if (rowDate instanceof Date) {
      rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof rowDate === 'string') {
      rowDateStr = rowDate.split('T')[0];
    }
    
    // Если найдена хотя бы одна запись за эту дату, возвращаем true
    if (rowDateStr === dateStr) {
      return true;
    }
  }
  
  return false;
}

/**
 * Получить карту информации о кампаниях по их ID
 * @param {Array<number>} campaignIds - Массив ID кампаний
 * @return {Object} Карта campaignId -> {bid_type, name, status}
 */
function getCampaignInfoMap(campaignIds) {
  var campaignInfoMap = {};
  
  if (!campaignIds || campaignIds.length === 0) {
    return campaignInfoMap;
  }
  
  // API лимит - максимум 50 ID за раз
  var batchSize = 50;
  
  for (var i = 0; i < campaignIds.length; i += batchSize) {
    var batch = campaignIds.slice(i, Math.min(i + batchSize, campaignIds.length));
    Logger.log('Запрос информации о кампаниях, батч ' + (Math.floor(i / batchSize) + 1) + ', кампаний: ' + batch.length);
    
    try {
      // Получаем кампании по ID
      var campaigns = getAdverts({
        ids: batch.join(',')
      });
      
      if (campaigns && campaigns.length > 0) {
        for (var j = 0; j < campaigns.length; j++) {
          var campaign = campaigns[j];
          if (campaign.id) {
            campaignInfoMap[campaign.id] = {
              bid_type: campaign.bid_type || '',
              name: (campaign.settings && campaign.settings.name) ? campaign.settings.name : '',
              status: campaign.status || ''
            };
          }
        }
        Logger.log('Получено информации о кампаниях: ' + campaigns.length);
      }
    } catch (error) {
      Logger.log('Ошибка получения информации о кампаниях для батча: ' + error.toString());
      // Продолжаем обработку следующего батча
    }
    
    // Ждем между запросами для соблюдения лимита API (5 запросов в секунду)
    if (i + batchSize < campaignIds.length) {
      Utilities.sleep(1000); // 1 секунда
    }
  }
  
  return campaignInfoMap;
}
