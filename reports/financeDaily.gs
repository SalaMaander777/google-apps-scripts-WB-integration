/**
 * Модуль для ежедневных финансовых отчетов
 * Выгружает данные за предыдущий день с дозаписыванием
 */

/**
 * Заголовки столбцов для листа "Ежедневные фин отчеты"
 */
function getFinanceDailyHeaders() {
  return [
    'Дата',
    'rrd_id',
    'gi_id',
    'subject_name',
    'nm_id',
    'brand_name',
    'sa_name',
    'ts_name',
    'barcode',
    'doc_type_name',
    'quantity',
    'retail_price',
    'retail_amount',
    'sale_percent',
    'commission_percent',
    'office_name',
    'supplier_oper_name',
    'order_dt',
    'sale_dt',
    'rr_dt',
    'shk_id',
    'retail_price_withdisc_rub',
    'delivery_amount',
    'return_amount',
    'delivery_rub',
    'gi_box_type_name',
    'product_discount_for_report',
    'supplier_promo',
    'rid',
    'ppvz_spp_prc',
    'ppvz_kvw_prc_base',
    'ppvz_kvw_prc',
    'ppvz_sales_commission',
    'ppvz_for_pay',
    'ppvz_reward',
    'acquiring_fee',
    'acquiring_bank',
    'ppvz_vw',
    'ppvz_vw_nds',
    'ppvz_office_id',
    'ppvz_office_name',
    'ppvz_supplier_id',
    'ppvz_supplier_name',
    'ppvz_inn',
    'declaration_number',
    'bonus_type_name',
    'sticker_id',
    'site_country',
    'penalty',
    'additional_payment',
    'rebill_logistic_cost',
    'rebill_logistic_org',
    'kiz',
    'storage_fee',
    'deduction',
    'acceptance',
    'srid'
  ];
}

/**
 * Преобразовать запись API в строку для таблицы
 * @param {Object} record - Запись из API
 * @param {string} reportDate - Дата отчета (YYYY-MM-DD)
 * @return {Array} Массив значений для строки таблицы
 */
function convertRecordToRow(record, reportDate) {
  return [
    reportDate, // Дата
    record.rrd_id || '',
    record.gi_id || '',
    record.subject_name || '',
    record.nm_id || '',
    record.brand_name || '',
    record.sa_name || '',
    record.ts_name || '',
    record.barcode || '',
    record.doc_type_name || '',
    record.quantity || '',
    record.retail_price || '',
    record.retail_amount || '',
    record.sale_percent || '',
    record.commission_percent || '',
    record.office_name || '',
    record.supplier_oper_name || '',
    record.order_dt || '',
    record.sale_dt || '',
    record.rr_dt || '',
    record.shk_id || '',
    record.retail_price_withdisc_rub || '',
    record.delivery_amount || '',
    record.return_amount || '',
    record.delivery_rub || '',
    record.gi_box_type_name || '',
    record.product_discount_for_report || '',
    record.supplier_promo || '',
    record.rid || '',
    record.ppvz_spp_prc || '',
    record.ppvz_kvw_prc_base || '',
    record.ppvz_kvw_prc || '',
    record.ppvz_sales_commission || '',
    record.ppvz_for_pay || '',
    record.ppvz_reward || '',
    record.acquiring_fee || '',
    record.acquiring_bank || '',
    record.ppvz_vw || '',
    record.ppvz_vw_nds || '',
    record.ppvz_office_id || '',
    record.ppvz_office_name || '',
    record.ppvz_supplier_id || '',
    record.ppvz_supplier_name || '',
    record.ppvz_inn || '',
    record.declaration_number || '',
    record.bonus_type_name || '',
    record.sticker_id || '',
    record.site_country || '',
    record.penalty || '',
    record.additional_payment || '',
    record.rebill_logistic_cost || '',
    record.rebill_logistic_org || '',
    record.kiz || '',
    record.storage_fee || '',
    record.deduction || '',
    record.acceptance || '',
    record.srid || ''
  ];
}

/**
 * Основная функция для выгрузки ежедневных финансовых отчетов
 * Выгружает данные за предыдущий день с дозаписыванием
 */
function syncFinanceDailyReport() {
  try {
    Logger.log('=== Начало синхронизации ежедневных финансовых отчетов ===');
    
    // Получаем дату предыдущего дня
    var dateRange = getPreviousDayRange();
    var reportDate = getPreviousDay();
    
    Logger.log('Дата отчета: ' + reportDate);
    Logger.log('Период: ' + dateRange.dateFrom + ' - ' + dateRange.dateTo);
    
    // Получаем лист
    var sheetName = getFinanceDailySheetName();
    var sheet = getOrCreateSheet(sheetName);
    
    // Устанавливаем заголовки если лист пуст
    var headers = getFinanceDailyHeaders();
    setSheetHeaders(sheet, headers);
    
    // Проверяем, не загружены ли уже данные за эту дату
    if (dateExistsInSheet(sheet, reportDate, 1)) {
      Logger.log('Данные за ' + reportDate + ' уже существуют в таблице. Пропускаем загрузку.');
      return;
    }
    
    // Загружаем данные из API
    Logger.log('Загрузка данных из API Wildberries...');
    var records = getReportDetailByPeriod(dateRange.dateFrom, dateRange.dateTo, 'daily');
    
    if (!records || records.length === 0) {
      Logger.log('Нет данных для загрузки за ' + reportDate);
      return;
    }
    
    Logger.log('Получено записей из API: ' + records.length);
    
    // Преобразуем данные в формат таблицы
    var rows = [];
    for (var i = 0; i < records.length; i++) {
      var row = convertRecordToRow(records[i], reportDate);
      rows.push(row);
    }
    
    // Записываем данные в таблицу
    Logger.log('Запись данных в таблицу...');
    appendDataToSheet(sheet, rows);
    
    Logger.log('=== Синхронизация завершена успешно. Добавлено строк: ' + rows.length + ' ===');
    
  } catch (error) {
    Logger.log('ОШИБКА при синхронизации ежедневных финансовых отчетов: ' + error.toString());
    Logger.log('Стек ошибки: ' + error.stack);
    throw error;
  }
}
