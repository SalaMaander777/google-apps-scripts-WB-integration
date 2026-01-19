/**
 * Главный файл проекта
 * Содержит функции запуска для триггеров
 */

/**
 * Функция вызывается при открытии таблицы
 * Создает меню с кнопкой "Настройки"
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Настройки скрипта')
    .addItem('Открыть настройки', 'showSettingsDialog')
    .addSeparator()
    .addItem('Синхронизация остатков', 'manualStocksSync')
    .addItem('Синхронизация финансовых отчетов', 'manualFinanceDailySync')
    .addItem('Синхронизация ленты заказов', 'manualOrdersFeedSync')
    .addSeparator()
    .addItem('Синхронизация аналитики РК', 'manualAdsAnalyticsSync')
    .addItem('Синхронизация истории рекламных расходов', 'manualAdsCostsSync')
    .addItem('Синхронизация аналитики продавца', 'manualSalesFunnelSync')
    .addToUi();
}

/**
 * Показать диалоговое окно с настройками
 */
function showSettingsDialog() {
  var html = HtmlService.createHtmlOutputFromFile('settings')
    .setWidth(550)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройки скрипта');
}

/**
 * Получить текущие настройки для отображения в форме
 * @return {Object} Объект с настройками
 */
function getSettings() {
  try {
    var props = PropertiesService.getScriptProperties();
    return {
      wbApiToken: props.getProperty('WB_API_TOKEN') ? '***' : '' // Не показываем токен из соображений безопасности
    };
  } catch (error) {
    Logger.log('Ошибка получения настроек: ' + error.toString());
    return null;
  }
}

/**
 * Сохранить настройки
 * @param {string} wbApiToken - API токен Wildberries
 * @return {Object} Результат сохранения
 */
function saveSettings(wbApiToken) {
  try {
    var props = PropertiesService.getScriptProperties();
    
    // Валидация токена
    if (!wbApiToken || wbApiToken.trim() === '') {
      return {
        success: false,
        error: 'API токен не может быть пустым'
      };
    }
    
    // Сохраняем ID активной таблицы автоматически
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSpreadsheet) {
      props.setProperty('SPREADSHEET_ID', activeSpreadsheet.getId());
    }
    
    // Сохраняем токен
    props.setProperty('WB_API_TOKEN', wbApiToken.trim());
    
    Logger.log('Настройки успешно сохранены');
    
    return {
      success: true
    };
    
  } catch (error) {
    Logger.log('Ошибка сохранения настроек: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Функция для запуска синхронизации ежедневных финансовых отчетов
 * Вызывается по триггеру ежедневно
 */
function runFinanceDailySync() {
  try {
    syncFinanceDailyReport();
  } catch (error) {
    Logger.log('Критическая ошибка в runFinanceDailySync: ' + error.toString());
    // Не пробрасываем ошибку дальше, чтобы не ломать другие триггеры
  }
}

/**
 * Функция для ручного запуска синхронизации ежедневных финансовых отчетов
 * Можно вызвать из меню или вручную
 */
function manualFinanceDailySync() {
  syncFinanceDailyReport();
}

/**
 * Функция для запуска синхронизации ленты заказов
 * Вызывается по триггеру ежедневно
 */
function runOrdersFeedSync() {
  try {
    syncOrdersFeed();
  } catch (error) {
    Logger.log('Критическая ошибка в runOrdersFeedSync: ' + error.toString());
    // Не пробрасываем ошибку дальше, чтобы не ломать другие триггеры
  }
}

/**
 * Функция для ручного запуска синхронизации ленты заказов
 * Можно вызвать из меню или вручную
 */
function manualOrdersFeedSync() {
  syncOrdersFeed();
}

/**
 * Функция для запуска синхронизации остатков товаров
 * Вызывается по триггеру ежедневно
 */
function runStocksSync() {
  try {
    syncStocks();
  } catch (error) {
    Logger.log('Критическая ошибка в runStocksSync: ' + error.toString());
    // Не пробрасываем ошибку дальше, чтобы не ломать другие триггеры
  }
}

/**
 * Функция для ручного запуска синхронизации остатков товаров
 * Можно вызвать из меню или вручную
 */
function manualStocksSync() {
  syncStocks();
}

/**
 * Функция для запуска синхронизации аналитики рекламных кампаний
 * Вызывается по триггеру ежедневно
 */
function runAdsAnalyticsSync() {
  try {
    syncAdsAnalytics();
  } catch (error) {
    Logger.log('Критическая ошибка в runAdsAnalyticsSync: ' + error.toString());
    // Не пробрасываем ошибку дальше, чтобы не ломать другие триггеры
  }
}

/**
 * Функция для ручного запуска синхронизации аналитики РК
 * Можно вызвать из меню или вручную
 */
function manualAdsAnalyticsSync() {
  syncAdsAnalytics();
}

/**
 * Функция для запуска синхронизации истории рекламных расходов
 * Вызывается по триггеру ежедневно
 */
function runAdsCostsSync() {
  try {
    syncAdsCosts();
  } catch (error) {
    Logger.log('Критическая ошибка в runAdsCostsSync: ' + error.toString());
    // Не пробрасываем ошибку дальше, чтобы не ломать другие триггеры
  }
}

/**
 * Функция для ручного запуска синхронизации истории рекламных расходов
 * Можно вызвать из меню или вручную
 */
function manualAdsCostsSync() {
  syncAdsCosts();
}

/**
 * Функция для запуска синхронизации аналитики продавца
 * Вызывается по триггеру ежедневно
 */
function runSalesFunnelSync() {
  try {
    syncSalesFunnel();
  } catch (error) {
    Logger.log('Критическая ошибка в runSalesFunnelSync: ' + error.toString());
    // Не пробрасываем ошибку дальше, чтобы не ломать другие триггеры
  }
}

/**
 * Функция для ручного запуска синхронизации аналитики продавца
 * Можно вызвать из меню или вручную
 */
function manualSalesFunnelSync() {
  syncSalesFunnel();
}
