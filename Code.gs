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
      .addItem('📅 Синхронизация за дату...', 'showDateSelectorDialog')
      .addSeparator()
      .addItem('🗄️ Архивировать и очистить таблицу', 'showArchiveConfirmDialog')
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
  * Показать диалоговое окно выбора даты для синхронизации
  */
  function showDateSelectorDialog() {
    var html = HtmlService.createHtmlOutputFromFile('dateSelector')
      .setWidth(550)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Синхронизация за дату');
  }

  /**
  * Получить текущие настройки для отображения в форме
  * @return {Object} Объект с настройками
  */
  function getSettings() {
    try {
      Logger.log('getSettings вызвана');
      var props = PropertiesService.getScriptProperties();
      var token = props.getProperty('WB_API_TOKEN');
      Logger.log('Токен найден: ' + (token ? 'да' : 'нет'));
      return {
        wbApiToken: token ? '***' : '' // Не показываем токен из соображений безопасности
      };
    } catch (error) {
      Logger.log('Ошибка получения настроек: ' + error.toString());
      throw error; // Пробрасываем ошибку в HTML
    }
  }

  /**
  * Сохранить настройки
  * @param {string} wbApiToken - API токен Wildberries
  * @return {Object} Результат сохранения
  */
  function saveSettings(wbApiToken) {
    try {
      Logger.log('saveSettings вызвана с токеном длиной: ' + (wbApiToken ? wbApiToken.length : 0));
      var props = PropertiesService.getScriptProperties();
      
      // Валидация токена
      if (!wbApiToken || wbApiToken.trim() === '') {
        Logger.log('Токен пустой');
        return {
          success: false,
          error: 'API токен не может быть пустым'
        };
      }
      
      // Сохраняем ID активной таблицы автоматически
      var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (activeSpreadsheet) {
        var spreadsheetId = activeSpreadsheet.getId();
        props.setProperty('SPREADSHEET_ID', spreadsheetId);
        Logger.log('SPREADSHEET_ID сохранен: ' + spreadsheetId);
      }
      
      // Сохраняем токен
      props.setProperty('WB_API_TOKEN', wbApiToken.trim());
      Logger.log('WB_API_TOKEN успешно сохранен');
      
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

  /**
  * Функция для запуска всех ежедневных синхронизаций
  * Вызывается по триггеру ежедневно
  * После синхронизации всех отчетов обновляет лист "Воронка динамика"
  */
  function runAllDailySync() {
    try {
      Logger.log('=== Начало ежедневной синхронизации всех отчетов ===');
      
      // 1. Синхронизация остатков
      try {
        syncStocks();
      } catch (error) {
        Logger.log('Ошибка синхронизации остатков: ' + error.toString());
      }
      
      // 2. Синхронизация финансовых отчетов
      try {
        syncFinanceDailyReport();
      } catch (error) {
        Logger.log('Ошибка синхронизации финансовых отчетов: ' + error.toString());
      }
      
      // 3. Синхронизация ленты заказов
      try {
        syncOrdersFeed();
      } catch (error) {
        Logger.log('Ошибка синхронизации ленты заказов: ' + error.toString());
      }
      
      // 4. Синхронизация аналитики РК
      try {
        syncAdsAnalytics();
      } catch (error) {
        Logger.log('Ошибка синхронизации аналитики РК: ' + error.toString());
      }
      
      // 5. Синхронизация истории рекламных расходов
      try {
        syncAdsCosts();
      } catch (error) {
        Logger.log('Ошибка синхронизации истории рекламных расходов: ' + error.toString());
      }
      
      // 6. Синхронизация аналитики продавца
      try {
        syncSalesFunnel();
      } catch (error) {
        Logger.log('Ошибка синхронизации аналитики продавца: ' + error.toString());
      }
      
      // 7. Обновление воронки динамики (добавление столбцов)
      try {
        updateSalesFunnelDynamic();
      } catch (error) {
        Logger.log('Ошибка обновления воронки динамики: ' + error.toString());
      }
      
      Logger.log('=== Ежедневная синхронизация всех отчетов завершена ===');
      
    } catch (error) {
      Logger.log('Критическая ошибка в runAllDailySync: ' + error.toString());
    }
  }

  /**
  * Функция для ручного запуска обновления листа "Воронка динамика"
  * Можно вызвать из меню или вручную
  */
  function manualUpdateSalesFunnelDynamic() {
    updateSalesFunnelDynamic();
  }

  /**
  * Функция для ручного переформатирования листа "Воронка динамика"
  * Применяет новые правила дизайна ко всем столбцам
  */
  function manualReformatSalesFunnelDynamic() {
    try {
      var result = reformatSalesFunnelDynamicSheet();
      if (result && result.success) {
        SpreadsheetApp.getUi().alert('Успешно: ' + result.message);
      }
    } catch (error) {
      SpreadsheetApp.getUi().alert('Ошибка: ' + error.toString());
    }
  }

  /**
  * Тестовая функция для проверки синхронизации отчетов за дату
  * Можно запустить вручную из Apps Script Editor
  */
  function testSyncReportsByDate() {
    var testDate = '2026-01-17'; // Измените на нужную дату
    var testReports = ['financeDaily']; // Выберите отчеты для тестирования
    
    Logger.log('Запуск тестовой синхронизации за дату: ' + testDate);
    Logger.log('Отчеты: ' + testReports.join(', '));
    
    var result = syncReportsByDate(testDate, testReports);
    
    Logger.log('Результат: ' + JSON.stringify(result));
    return result;
  }

  /**
  * Синхронизация отчетов за выбранную дату
  * @param {string} date - Дата в формате YYYY-MM-DD
  * @param {Array<string>} reports - Массив названий отчетов для синхронизации
  * @return {Object} Результат синхронизации
  */
  function syncReportsByDate(date, reports) {
    try {
      Logger.log('=== syncReportsByDate вызвана ===');
      Logger.log('Дата: ' + date);
      Logger.log('Тип данных date: ' + typeof date);
      Logger.log('Отчеты (JSON): ' + JSON.stringify(reports));
      Logger.log('Тип данных reports: ' + typeof reports);
      Logger.log('reports является массивом: ' + Array.isArray(reports));
      
      // Проверка параметров
      if (!date) {
        Logger.log('ОШИБКА: Параметр date не передан');
        return {
          success: false,
          message: 'Дата не указана'
        };
      }
      
      if (!reports || !Array.isArray(reports) || reports.length === 0) {
        Logger.log('ОШИБКА: Параметр reports не передан или пустой');
        return {
          success: false,
          message: 'Не выбраны отчеты для синхронизации'
        };
      }
      
      Logger.log('=== Начало синхронизации отчетов за дату: ' + date + ' ===');
      Logger.log('Выбранные отчеты: ' + reports.join(', '));
      
      var results = [];
      var errors = [];
      
      // Синхронизация каждого выбранного отчета
      for (var i = 0; i < reports.length; i++) {
        var reportType = reports[i];
        
        try {
          Logger.log('Синхронизация: ' + reportType);
          
          switch(reportType) {
            case 'stocks':
              syncStocksByDate(date);
              results.push('Остатки товаров');
              break;
              
            case 'financeDaily':
              syncFinanceDailyReportByDate(date);
              results.push('Финансовые отчеты');
              break;
              
            case 'ordersFeed':
              syncOrdersFeedByDate(date);
              results.push('Лента заказов');
              break;
              
            case 'adsAnalytics':
              syncAdsAnalyticsByDate(date);
              results.push('Аналитика РК');
              break;
              
            case 'adsCosts':
              syncAdsCostsByDate(date);
              results.push('История рекламных расходов');
              break;
              
            case 'salesFunnel':
              syncSalesFunnelByDate(date);
              results.push('Аналитика продавца');
              break;
              
            case 'funnelDynamic':
              addSalesFunnelDynamicColumn(date);
              results.push('Воронка динамика');
              break;
              
            case 'funnelDynamicWeek':
              syncFunnelDynamicWeekByDate(date);
              results.push('Воронка динамика (неделя)');
              break;
              
            default:
              Logger.log('Неизвестный тип отчета: ' + reportType);
          }
          
        } catch (error) {
          Logger.log('Ошибка синхронизации ' + reportType + ': ' + error.toString());
          errors.push(reportType);
        }
      }
      
      Logger.log('=== Синхронизация завершена ===');
      
      var message = 'Загружено отчетов: ' + results.length;
      if (errors.length > 0) {
        message += '. Ошибки: ' + errors.length;
      }
      
      return {
        success: true,
        message: message,
        results: results,
        errors: errors
      };
      
    } catch (error) {
      Logger.log('Критическая ошибка в syncReportsByDate: ' + error.toString());
      return {
        success: false,
        message: error.toString()
      };
    }
  }
