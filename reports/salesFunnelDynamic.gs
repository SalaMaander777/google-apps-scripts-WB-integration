  /**
  * Модуль для работы с листом "Воронка динамика"
  * Добавляет столбцы с датами и формулами после выгрузки всех отчетов
  */

  /**
  * Обновить лист "Воронка динамика" - добавить столбец с датой
  * Вызывается после выгрузки всех ежедневных отчетов
  */
  function updateSalesFunnelDynamic() {
    try {
      Logger.log('=== Начало обновления листа "Воронка динамика" ===');
      
      // Получаем дату предыдущего дня
      var reportDate = getPreviousDay();
      Logger.log('Дата отчета: ' + reportDate);
      
      // 1. Добавляем дневной столбец за вчера
      addSalesFunnelDynamicColumn(reportDate);
      
      // 2. Если вчера было воскресенье, автоматически добавляем недельный отчет
      // Это происходит при запуске в понедельник утром за вчерашнее воскресенье
      var dateObj = parseDate(reportDate);
      if (dateObj && dateObj.getDay() === 0) { // 0 = воскресенье
        Logger.log('Вчера было воскресенье (' + reportDate + '). Добавляем недельный отчет.');
        try {
          addSalesFunnelDynamicWeekColumn(reportDate);
        } catch (weekError) {
          Logger.log('Ошибка при добавлении недельного отчета: ' + weekError.toString());
          // Не прерываем основное выполнение, если не удалось добавить только недельный отчет
        }
      }
      
      Logger.log('=== Обновление листа "Воронка динамика" завершено успешно ===');
      
    } catch (error) {
      Logger.log('ОШИБКА при обновлении листа "Воронка динамика": ' + error.toString());
      Logger.log('Стек ошибки: ' + error.stack);
      throw error;
    }
  }

  /**
  * Добавить столбец в лист "Воронка динамика" за указанную дату
  * Столбец вставляется в хронологическом порядке
  * @param {string} date - Дата в формате YYYY-MM-DD
  * @return {Object} Результат операции
  */
  function addSalesFunnelDynamicColumn(date) {
    try {
      Logger.log('Добавление столбца для даты: ' + date);
      
      // Получаем или создаем лист
      var sheetName = getSalesFunnelDynamicSheetName();
      var sheet = getOrCreateSheet(sheetName);
      
      // Проверяем, что лист инициализирован (есть заголовки в столбце A)
      if (sheet.getLastRow() < 50) {
        var errorMsg = 'Лист "Воронка динамика" не инициализирован. Должно быть минимум 50 строк. Текущее количество строк: ' + sheet.getLastRow();
        Logger.log('ОШИБКА: ' + errorMsg);
        throw new Error(errorMsg);
      }
      
      // Находим правильную позицию для вставки столбца
      var insertPosition = findColumnPositionByDate(sheet, date);
      
      if (insertPosition === -1) {
        Logger.log('Столбец с датой ' + date + ' уже существует. Пропускаем.');
        return { success: true, message: 'Столбец уже существует', skipped: true };
      }
      
      Logger.log('Позиция для вставки столбца: ' + insertPosition);
      
      // Вставляем новый столбец в нужную позицию
      if (insertPosition <= sheet.getLastColumn()) {
        sheet.insertColumnBefore(insertPosition);
        Logger.log('Вставлен новый столбец перед позицией ' + insertPosition);
      }
      
      // Заполняем столбец данными и формулами
      addDailyColumn(sheet, insertPosition, date);
      
      Logger.log('Столбец успешно добавлен для даты ' + date);
      
      return { success: true, message: 'Столбец успешно добавлен', skipped: false };
      
    } catch (error) {
      Logger.log('ОШИБКА при добавлении столбца: ' + error.toString());
      throw error;
    }
  }

  /**
  * Найти позицию для вставки столбца по дате (хронологический порядок)
  * @param {Sheet} sheet - Лист
  * @param {string} newDate - Дата в формате YYYY-MM-DD
  * @return {number} Позиция для вставки (или -1 если дата уже существует)
  */
  function findColumnPositionByDate(sheet, newDate) {
    var lastCol = sheet.getLastColumn();
    
    // Если только столбец A (заголовки), добавляем в B
    if (lastCol === 1) {
      return 2;
    }
    
    // Читаем даты из строки 7 (Дата с)
    var datesRow = sheet.getRange(7, 2, 1, lastCol - 1).getValues()[0];
    
    // Преобразуем новую дату в объект Date для сравнения
    var newDateObj = parseDateString(newDate);
    
    // Ищем правильную позицию
    for (var i = 0; i < datesRow.length; i++) {
      var cellValue = datesRow[i];
      
      // Пропускаем пустые ячейки
      if (!cellValue || cellValue === '') {
        continue;
      }
      
      // Получаем дату из ячейки
      var existingDateObj = parseCellDate(cellValue);
      
      if (!existingDateObj) {
        continue;
      }
      
      // Сравниваем даты
      var comparison = compareDates(newDateObj, existingDateObj);
      
      if (comparison === 0) {
        // Дата уже существует
        return -1;
      } else if (comparison < 0) {
        // Новая дата раньше существующей - вставляем здесь
        return i + 2; // +2 потому что i начинается с 0, а столбец B = 2
      }
    }
    
    // Новая дата позже всех существующих - добавляем в конец
    return lastCol + 1;
  }

  /**
  * Преобразовать строку даты YYYY-MM-DD в объект Date
  * @param {string} dateStr - Дата в формате YYYY-MM-DD
  * @return {Date} Объект Date (только дата, без времени)
  */
  function parseDateString(dateStr) {
    var parts = dateStr.split('-');
    var date = new Date(parts[0], parts[1] - 1, parts[2]);
    // Обнуляем время, чтобы в ячейке была только дата
    date.setHours(0, 0, 0, 0);
    return date;
  }

  /**
  * Преобразовать значение ячейки в объект Date
  * @param {*} cellValue - Значение ячейки
  * @return {Date|null} Объект Date или null
  */
  function parseCellDate(cellValue) {
    if (cellValue instanceof Date) {
      var date = new Date(cellValue);
      date.setHours(0, 0, 0, 0);
      return date;
    }
    
    if (typeof cellValue === 'string') {
      // Формат DD.MM.YYYY
      var parts = cellValue.split('.');
      if (parts.length === 3) {
        var date = new Date(parts[2], parts[1] - 1, parts[0]);
        date.setHours(0, 0, 0, 0);
        return date;
      }
      
      // Формат M/D/YYYY
      parts = cellValue.split('/');
      if (parts.length === 3) {
        var date = new Date(parts[2], parts[0] - 1, parts[1]);
        date.setHours(0, 0, 0, 0);
        return date;
      }

      // Формат YYYY-MM-DD
      parts = cellValue.split('-');
      if (parts.length === 3) {
        var date = new Date(parts[0], parts[1] - 1, parts[2]);
        date.setHours(0, 0, 0, 0);
        return date;
      }
    }
    
    return null;
  }

  /**
  * Сравнить две даты (только год, месяц, день)
  * @param {Date} date1 - Первая дата
  * @param {Date} date2 - Вторая дата
  * @return {number} -1 если date1 < date2, 0 если равны, 1 если date1 > date2
  */
  function compareDates(date1, date2) {
    if (!date1 || !date2) return 0;
    
    var d1 = new Date(date1);
    var d2 = new Date(date2);
    d1.setHours(0, 0, 0, 0);
    d2.setHours(0, 0, 0, 0);
    
    var t1 = d1.getTime();
    var t2 = d2.getTime();
    
    if (t1 < t2) return -1;
    if (t1 > t2) return 1;
    return 0;
  }

  /**
  * Добавить столбец с ежедневными данными
  * @param {Sheet} sheet - Лист
  * @param {number} col - Номер столбца
  * @param {string} date - Дата в формате YYYY-MM-DD
  */
  function addDailyColumn(sheet, col, date) {
    Logger.log('Добавление ежедневного столбца: колонка=' + col + ', дата=' + date);
    
    // Преобразуем дату в американский формат M/D/YYYY для Google Sheets
    var usDate = formatDateUs(date);
    
    // Получаем букву столбца для использования в формулах
    var colLetter = getColumnLetter(col);
    
    // Массив для хранения всех значений столбца (58 строк)
    var columnData = [];
    
    // Строки 1-6: служебная информация (пустые для нового столбца)
    for (var i = 1; i <= 6; i++) {
      columnData.push(['']);
    }
    
    // Строка 7: Дата с (начальная дата) - в формате M/D/YYYY
    columnData.push([usDate]);
    
    // Строка 8: Дата до (конечная дата) - в формате M/D/YYYY
    columnData.push([usDate]);
    
    // Строка 9: Заказы по артикулу
    columnData.push(['=SUMIFS(\'Аналитика продавца\'!$M:$M,\'Аналитика продавца\'!$B:$B,$A$2,\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
    
    // Строка 10: Динамика заказов
    var prevColLetter = getColumnLetter(col - 1);
    columnData.push(['=IFERROR(' + colLetter + '9/' + prevColLetter + '9-1,0)']);
    
    // Строка 11: Сумма заказов по артикулу в руб
    columnData.push(['=SUMIFS(\'Аналитика продавца\'!$Y:$Y,\'Аналитика продавца\'!$B:$B,$A$2,\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
    
    // Строка 12: Динамика суммы заказов
    columnData.push(['=IFERROR(' + colLetter + '11/' + prevColLetter + '11-1,0)']);
    
    // Строка 13: Заказы по модели
    columnData.push(['=SUMIFS(\'Аналитика продавца\'!$M:$M,\'Аналитика продавца\'!$B:$B,"*"&$A$4&"*",\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
    
    // Строка 14: Динамика заказов по модели
    columnData.push(['=IFERROR(' + colLetter + '13/' + prevColLetter + '13-1,0)']);
    
    // Строка 15: Сумма заказов по модели в руб
    columnData.push(['=SUMIFS(\'Аналитика продавца\'!$Y:$Y,\'Аналитика продавца\'!$B:$B,"*"&$A$4&"*",\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
    
    // Строка 16: Доля заказов артикула в группе по модели
    columnData.push(['=IFERROR(' + colLetter + '11/' + colLetter + '15,0)']);
    
    // Строка 17: Выкупы по артикулу
    columnData.push(['=SUMIFS(\'Ежедневные ф отчеты\'!$N:$N,\'Ежедневные ф отчеты\'!$F:$F,$A$2,\'Ежедневные ф отчеты\'!$M:$M,">="&' + colLetter + '7,\'Ежедневные ф отчеты\'!$M:$M,"<="&' + colLetter + '8,\'Ежедневные ф отчеты\'!$J:$J,"Продажа")']);
    
    // Строка 18: Выкупы по модели
    columnData.push(['=SUMIFS(\'Ежедневные ф отчеты\'!$N:$N,\'Ежедневные ф отчеты\'!$F:$F,"*"&$A$4&"*",\'Ежедневные ф отчеты\'!$M:$M,">="&' + colLetter + '7,\'Ежедневные ф отчеты\'!$M:$M,"<="&' + colLetter + '8,\'Ежедневные ф отчеты\'!$J:$J,"Продажа")']);
    
    // Строка 19: Сумма выкупов по артикулу (фин отчет)
    columnData.push(['=SUMIFS(\'Ежедневные ф отчеты\'!$O:$O,\'Ежедневные ф отчеты\'!$F:$F,$A$2,\'Ежедневные ф отчеты\'!$M:$M,">="&' + colLetter + '7,\'Ежедневные ф отчеты\'!$M:$M,"<="&' + colLetter + '8,\'Ежедневные ф отчеты\'!$J:$J,"Продажа")']);
    
    // Строка 20: Динамика выкупы артикул, руб
    columnData.push(['=IFERROR(' + colLetter + '17/' + prevColLetter + '17-1,0)']);
    
    // Строка 21: Динамика выкупы модели, руб
    columnData.push(['=IFERROR(' + colLetter + '18/' + prevColLetter + '18-1,0)']);
    
    // Строка 22: Возвраты
    columnData.push(['=SUMIFS(\'Ежедневные ф отчеты\'!$N:$N,\'Ежедневные ф отчеты\'!$F:$F,$A$2,\'Ежедневные ф отчеты\'!$M:$M,">="&' + colLetter + '7,\'Ежедневные ф отчеты\'!$M:$M,"<="&' + colLetter + '8,\'Ежедневные ф отчеты\'!$K:$K,"Возврат")']);
    
    // Строка 23: Сумма Возвратов
    columnData.push(['=SUMIFS(\'Ежедневные ф отчеты\'!$T:$T,\'Ежедневные ф отчеты\'!$F:$F,$A$2,\'Ежедневные ф отчеты\'!$M:$M,">="&' + colLetter + '7,\'Ежедневные ф отчеты\'!$M:$M,"<="&' + colLetter + '8,\'Ежедневные ф отчеты\'!$K:$K,"Возврат")']);
    
    // Строка 24: % выкупа по артикулу, факт
    // Для ежедневного столбца используем формулу деления выкупов на заказы за текущий день
    var weekStartCol = getColumnLetter(Math.max(col - 6, 2)); // Начало недели (7 дней назад или столбец B)
    columnData.push(['=SUM(' + weekStartCol + '17:' + colLetter + '17)/SUM(' + weekStartCol + '9:' + colLetter + '9)']);
    
    // Строка 25: % выкупа по модели, факт
    columnData.push(['=SUM(' + weekStartCol + '18:' + colLetter + '18)/SUM(' + weekStartCol + '13:' + colLetter + '13)']);
    
    // Строка 26: CTR артикула
    columnData.push(['=IFERROR(' + colLetter + '30/' + colLetter + '28,0)']);
    
    // Строка 27: Переходы в карточку
    columnData.push(['=SUMIFS(\'Аналитика продавца\'!$I:$I,\'Аналитика продавца\'!$C:$C,$A$1,\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
    
    // Строка 28: Показы реклама по цвету
    columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$J:$J,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$J:$J,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    
    // Строка 29: Динамика показов по цвету
    columnData.push(['=IFERROR(' + colLetter + '28/' + prevColLetter + '28-1,0)']);
    
    // Строка 30: Клики реклама по цвету
    columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$K:$K,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$K:$K,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    
    // Строка 31: Сумма показов РК по артикулу (Аналитика РК M)
    columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$M:$M,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$M:$M,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    // Строка 32: Сумма кликов РК по артикулу (Аналитика РК S)
    columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$S:$S,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$S:$S,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    // Строка 33: (43/58 в новой структуре)
    columnData.push(['=IFERROR(' + colLetter + '46/' + colLetter + '32,0)']);
    // Строка 34: (58/57)
    columnData.push(['=IFERROR(' + colLetter + '32/' + colLetter + '31,0)']);
    // Строка 35: (58/37)
    columnData.push(['=IFERROR(' + colLetter + '31/' + colLetter + '37,0)']);
    
    // Строка 36: Показы реклама по модели (J)
    columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$J:$J,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$J:$J,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    // Строка 37: Клики реклама по модели (K)
    columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$K:$K,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$K:$K,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    
    // Строка 38: Сумма показов РК по модели (M)
    columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$M:$M,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$M:$M,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    // Строка 39: Сумма кликов РК по модели (S)
    columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$S:$S,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$S:$S,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    // Строка 40: (44/66 -> 53/58)
    columnData.push(['=IFERROR(' + colLetter + '47/' + colLetter + '38,0)']);
    
    // Строка 41: CTR модели
    columnData.push(['=IFERROR(' + colLetter + '37/' + colLetter + '36,0)']);
    // Строка 42: Корзина
    columnData.push(['=SUMIFS(\'Аналитика продавца\'!$K:$K,\'Аналитика продавца\'!$C:$C,$A$1,\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
    // Строка 43: Конверсия в корзину
    columnData.push(['=IFERROR(' + colLetter + '42/(' + colLetter + '27+' + colLetter + '30),0)']);
    // Строка 44: Конверсия в заказ
    columnData.push(['=IFERROR(' + colLetter + '9/' + colLetter + '42,0)']);
    // Строка 45: Общий коэффициент
    columnData.push(['=' + colLetter + '26*' + colLetter + '43*' + colLetter + '44*1000']);
    // Строка 46: Затраты на рекламу по артикулу
    columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'История рекламных расходов\'!$G:$G,\'История рекламных расходов\'!$I:$I,$A$2,\'История рекламных расходов\'!$D:$D,">="&' + colLetter + '7,\'История рекламных расходов\'!$D:$D,"<="&' + colLetter + '8)),(SUMIFS(\'История рекламных расходов\'!$G:$G,\'История рекламных расходов\'!$I:$I,$A$2,\'История рекламных расходов\'!$D:$D,">="&' + colLetter + '7,\'История рекламных расходов\'!$D:$D,"<="&' + colLetter + '8,\'История рекламных расходов\'!$B:$B,"*"&$A$6&"*")))']);
    // Строка 47: Затраты на рекламу по модели
    columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'История рекламных расходов\'!$G:$G,\'История рекламных расходов\'!$I:$I,"*"&$A$4&"*",\'История рекламных расходов\'!$D:$D,">="&' + colLetter + '7,\'История рекламных расходов\'!$D:$D,"<="&' + colLetter + '8)),(SUMIFS(\'История рекламных расходов\'!$G:$G,\'История рекламных расходов\'!$I:$I,"*"&$A$4&"*",\'История рекламных расходов\'!$D:$D,">="&' + colLetter + '7,\'История рекламных расходов\'!$D:$D,"<="&' + colLetter + '8,\'История рекламных расходов\'!$B:$B,"*"&$A$6&"*")))']);
    // Строка 48: CPO
    columnData.push(['=IFERROR(' + colLetter + '46/' + colLetter + '9,0)']);
    // Строка 49: CPC по артикулу
    columnData.push(['=IFERROR(' + colLetter + '46/' + colLetter + '30,0)']);
    // Строка 50: CPC по модели
    columnData.push(['=IFERROR(' + colLetter + '47/' + colLetter + '37,0)']);
    // Строка 51: CPS по артикулу
    columnData.push(['=IFERROR(' + colLetter + '46/(' + colLetter + '9*' + colLetter + '24),0)']);
    // Строка 52: CPS по модели
    columnData.push(['=IFERROR(' + colLetter + '47/(' + colLetter + '13*' + colLetter + '25),0)']);
    // Строка 53: ДРР фактическая от заказа цвета на цвет
    columnData.push(['=IFERROR(' + colLetter + '46/' + colLetter + '11,0)']);
    // Строка 54: ДРР вмененная от выкупа цвета на цвет
    columnData.push(['=IFERROR(' + colLetter + '46/(' + colLetter + '11*#REF!),0)']);
    // Строка 55: ДРР фактическая от выкупа цвет на цвет
    columnData.push(['=IFERROR(' + colLetter + '46/#REF!,0)']);
    // Строка 56: ДРР вмененная от выкупа цвета на модель
    columnData.push(['=IFERROR(' + colLetter + '46/(' + colLetter + '15*' + colLetter + '25),0)']);
    // Строка 57: ДРР вмененная всей модели
    columnData.push(['=IFERROR(' + colLetter + '47/(' + colLetter + '15*' + colLetter + '25),0)']);
    // Строка 58: Заказов на 1 клик, руб (по модели)
    columnData.push(['=IFERROR(' + colLetter + '15/' + colLetter + '37,0)']);
    
    // Записываем все данные в столбец одной операцией
    sheet.getRange(1, col, 58, 1).setValues(columnData);
    
    // Применяем формат даты dd.mm.yyyy к строкам 7 и 8
    var dateRange = sheet.getRange(7, col, 2, 1);
    dateRange.setNumberFormat('dd.mm.yyyy');
    
    // Целое число (без дробной части): заказы, показы, клики и т.д.
    var integerRows = [9, 13, 17, 18, 22, 27, 28, 30, 31, 32, 36, 37, 38, 39, 40, 42, 46, 47, 48, 50, 52, 58];
    for (var r = 0; r < integerRows.length; r++) {
      sheet.getRange(integerRows[r], col).setNumberFormat('#,##0');
    }
    // Процентный формат (динамика, доля, % выкупа, CTR, конверсия)
    var percentRows = [10, 12, 14, 16, 20, 21, 24, 25, 26, 29, 33, 34, 35, 41, 43, 44];
    for (var p = 0; p < percentRows.length; p++) {
      sheet.getRange(percentRows[p], col).setNumberFormat('0.00%');
    }
    // Российский рубль: строки с суммами в рублях
    var rubleRows = [11, 15, 19, 23];
    var rubleFormat = '#,##0.00" ₽"';
    for (var r = 0; r < rubleRows.length; r++) {
      sheet.getRange(rubleRows[r], col).setNumberFormat(rubleFormat);
    }
    // Числовой формат с дробной частью для остальных строк (49, 51, 53, 54, 55, 56, 57)
    var numberRows = [49, 51, 53, 54, 55, 56, 57];
    for (var r = 0; r < numberRows.length; r++) {
      sheet.getRange(numberRows[r], col).setNumberFormat('#,##0.00');
    }
    // Строка 45 — общий коэффициент: число
    sheet.getRange(45, col).setNumberFormat('#,##0.00');
    
    // Применяем форматирование для дневного столбца
    // Сначала очищаем фон и УБИРАЕМ ВСЕ РАМКИ со всего столбца (с 1 по 58 строку)
    var fullColumnRange = sheet.getRange(1, col, 58, 1);
    fullColumnRange.setBackground('#ffffff');
    fullColumnRange.setBorder(false, false, false, false, false, false);
    
  
    // 7 и 8 строка: цвет c5e0b3 и ЖИРНАЯ РАМКА
    var range7_8 = sheet.getRange(7, col, 2, 1);
    range7_8.setBackground('#c5e0b3');
    // Устанавливаем рамку только для дат (внешнюю и внутреннюю горизонтальную)
    range7_8.setBorder(true, true, true, true, null, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    // 9 и 13 строка: ярко-зеленый
    sheet.getRange(9, col).setBackground('#00ff00');
    sheet.getRange(13, col).setBackground('#00ff00');
    
    // 11 строка: светло-голубой
    sheet.getRange(11, col).setBackground('#d0e0e3');
    
    // 17 и 18 строка: средне-зеленый
    sheet.getRange(17, col, 2, 1).setBackground('#92d050');
    
    // 24 и 25 строка: серый
    sheet.getRange(24, col, 2, 1).setBackground('#d9d9d9');
    
    // 27-30: светло-голубой(3)
    sheet.getRange(27, col, 4, 1).setBackground('#d0e0e3');
    
    // 31-35: новые строки РК по артикулу (M, S и формулы)
    sheet.getRange(31, col, 5, 1).setBackground('#d9ead3');
    
    // 36-37: показы/клики по модели (J, K)
    sheet.getRange(36, col, 2, 1).setBackground('#d9ead3');
    
    // 38-40: новые строки РК по модели (M, S и формула)
    sheet.getRange(38, col, 3, 1).setBackground('#d9ead3');
    
    // 41: CTR модели
    sheet.getRange(41, col).setBackground('#d9ead3');
    
    // 42: светло-голубой(3)
    sheet.getRange(42, col).setBackground('#d0e0e3');
    
    // 43-44: светло-бирюзовый(3)
    sheet.getRange(43, col, 2, 1).setBackground('#b2dfdb');
    
    // 45: общий коэффициент
    sheet.getRange(45, col).setBackground('#d9d9d9');
    
    // 48-49: светло-бирюзовый(3)
    sheet.getRange(48, col, 2, 1).setBackground('#b2dfdb');
    
    // 50: светло-голубой(3)
    sheet.getRange(50, col).setBackground('#d0e0e3');
    
    // 51-52: светло-бирюзовый(3)
    sheet.getRange(51, col, 2, 1).setBackground('#b2dfdb');
    
    // 53: светло-оранжевый(3)
    sheet.getRange(53, col).setBackground('#ffe0b2');
    
    // 54: светло-оранжевый(2)
    sheet.getRange(54, col).setBackground('#ffcc80');
    
    // 55-56: светло-оранжевый(3)
    sheet.getRange(55, col, 2, 1).setBackground('#ffe0b2');
    
    // 57: светло-зеленый(3)
    sheet.getRange(57, col).setBackground('#b6d7a8');
    // -----------------------------------------------------
    
    Logger.log('Столбец ' + colLetter + ' успешно добавлен');
  }

  /**
  * Получить букву столбца по его номеру
  * @param {number} col - Номер столбца (1 = A, 2 = B, и т.д.)
  * @return {string} Буква столбца
  */
  function getColumnLetter(col) {
    var letter = '';
    while (col > 0) {
      var mod = (col - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      col = Math.floor((col - mod) / 26);
    }
    return letter;
  }

  /**
  * Получить формулы для строк 9–58 дневного столбца (для обновления при переформатировании)
  * @param {number} col - Номер столбца
  * @return {Array.<Array.<string>>} Двумерный массив 50x1 (формулы для строк 9–58)
  */
  function getDailyColumnFormulas(col) {
    var colLetter = getColumnLetter(col);
    var prevColLetter = getColumnLetter(col - 1);
    var weekStartCol = getColumnLetter(Math.max(col - 6, 2));
    var formulas = [];
    formulas.push(['=SUMIFS(\'Аналитика продавца\'!$M:$M,\'Аналитика продавца\'!$B:$B,$A$2,\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
    formulas.push(['=IFERROR(' + colLetter + '9/' + prevColLetter + '9-1,0)']);
    formulas.push(['=SUMIFS(\'Аналитика продавца\'!$Y:$Y,\'Аналитика продавца\'!$B:$B,$A$2,\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
    formulas.push(['=IFERROR(' + colLetter + '11/' + prevColLetter + '11-1,0)']);
    formulas.push(['=SUMIFS(\'Аналитика продавца\'!$M:$M,\'Аналитика продавца\'!$B:$B,"*"&$A$4&"*",\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
    formulas.push(['=IFERROR(' + colLetter + '13/' + prevColLetter + '13-1,0)']);
    formulas.push(['=SUMIFS(\'Аналитика продавца\'!$Y:$Y,\'Аналитика продавца\'!$B:$B,"*"&$A$4&"*",\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
    formulas.push(['=IFERROR(' + colLetter + '11/' + colLetter + '15,0)']);
    formulas.push(['=SUMIFS(\'Ежедневные ф отчеты\'!$N:$N,\'Ежедневные ф отчеты\'!$F:$F,$A$2,\'Ежедневные ф отчеты\'!$M:$M,">="&' + colLetter + '7,\'Ежедневные ф отчеты\'!$M:$M,"<="&' + colLetter + '8,\'Ежедневные ф отчеты\'!$J:$J,"Продажа")']);
    formulas.push(['=SUMIFS(\'Ежедневные ф отчеты\'!$N:$N,\'Ежедневные ф отчеты\'!$F:$F,"*"&$A$4&"*",\'Ежедневные ф отчеты\'!$M:$M,">="&' + colLetter + '7,\'Ежедневные ф отчеты\'!$M:$M,"<="&' + colLetter + '8,\'Ежедневные ф отчеты\'!$J:$J,"Продажа")']);
    formulas.push(['=SUMIFS(\'Ежедневные ф отчеты\'!$O:$O,\'Ежедневные ф отчеты\'!$F:$F,$A$2,\'Ежедневные ф отчеты\'!$M:$M,">="&' + colLetter + '7,\'Ежедневные ф отчеты\'!$M:$M,"<="&' + colLetter + '8,\'Ежедневные ф отчеты\'!$J:$J,"Продажа")']);
    formulas.push(['=IFERROR(' + colLetter + '17/' + prevColLetter + '17-1,0)']);
    formulas.push(['=IFERROR(' + colLetter + '18/' + prevColLetter + '18-1,0)']);
    formulas.push(['=SUMIFS(\'Ежедневные ф отчеты\'!$N:$N,\'Ежедневные ф отчеты\'!$F:$F,$A$2,\'Ежедневные ф отчеты\'!$M:$M,">="&' + colLetter + '7,\'Ежедневные ф отчеты\'!$M:$M,"<="&' + colLetter + '8,\'Ежедневные ф отчеты\'!$K:$K,"Возврат")']);
    formulas.push(['=SUMIFS(\'Ежедневные ф отчеты\'!$T:$T,\'Ежедневные ф отчеты\'!$F:$F,$A$2,\'Ежедневные ф отчеты\'!$M:$M,">="&' + colLetter + '7,\'Ежедневные ф отчеты\'!$M:$M,"<="&' + colLetter + '8,\'Ежедневные ф отчеты\'!$K:$K,"Возврат")']);
    formulas.push(['=SUM(' + weekStartCol + '17:' + colLetter + '17)/SUM(' + weekStartCol + '9:' + colLetter + '9)']);
    formulas.push(['=SUM(' + weekStartCol + '18:' + colLetter + '18)/SUM(' + weekStartCol + '13:' + colLetter + '13)']);
    formulas.push(['=IFERROR(' + colLetter + '30/' + colLetter + '28,0)']);
    formulas.push(['=SUMIFS(\'Аналитика продавца\'!$I:$I,\'Аналитика продавца\'!$C:$C,$A$1,\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
    formulas.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$J:$J,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$J:$J,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    formulas.push(['=IFERROR(' + colLetter + '28/' + prevColLetter + '28-1,0)']);
    formulas.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$K:$K,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$K:$K,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    formulas.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$M:$M,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$M:$M,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    formulas.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$S:$S,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$S:$S,\'Аналитика РК\'!$L:$L,$A$2,\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    formulas.push(['=IFERROR(' + colLetter + '46/' + colLetter + '32,0)']);
    formulas.push(['=IFERROR(' + colLetter + '32/' + colLetter + '31,0)']);
    formulas.push(['=IFERROR(' + colLetter + '31/' + colLetter + '37,0)']);
    formulas.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$J:$J,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$J:$J,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    formulas.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$K:$K,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$K:$K,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    formulas.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$M:$M,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$M:$M,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    formulas.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$S:$S,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$S:$S,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
    formulas.push(['=IFERROR(' + colLetter + '47/' + colLetter + '38,0)']);
    formulas.push(['=IFERROR(' + colLetter + '37/' + colLetter + '36,0)']);
    formulas.push(['=SUMIFS(\'Аналитика продавца\'!$K:$K,\'Аналитика продавца\'!$C:$C,$A$1,\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
    formulas.push(['=IFERROR(' + colLetter + '42/(' + colLetter + '27+' + colLetter + '30),0)']);
    formulas.push(['=IFERROR(' + colLetter + '9/' + colLetter + '42,0)']);
    formulas.push(['=' + colLetter + '26*' + colLetter + '43*' + colLetter + '44*1000']);
    formulas.push(['=IF($A$6="ВСЕ",(SUMIFS(\'История рекламных расходов\'!$G:$G,\'История рекламных расходов\'!$I:$I,$A$2,\'История рекламных расходов\'!$D:$D,">="&' + colLetter + '7,\'История рекламных расходов\'!$D:$D,"<="&' + colLetter + '8)),(SUMIFS(\'История рекламных расходов\'!$G:$G,\'История рекламных расходов\'!$I:$I,$A$2,\'История рекламных расходов\'!$D:$D,">="&' + colLetter + '7,\'История рекламных расходов\'!$D:$D,"<="&' + colLetter + '8,\'История рекламных расходов\'!$B:$B,"*"&$A$6&"*")))']);
    formulas.push(['=IF($A$6="ВСЕ",(SUMIFS(\'История рекламных расходов\'!$G:$G,\'История рекламных расходов\'!$I:$I,"*"&$A$4&"*",\'История рекламных расходов\'!$D:$D,">="&' + colLetter + '7,\'История рекламных расходов\'!$D:$D,"<="&' + colLetter + '8)),(SUMIFS(\'История рекламных расходов\'!$G:$G,\'История рекламных расходов\'!$I:$I,"*"&$A$4&"*",\'История рекламных расходов\'!$D:$D,">="&' + colLetter + '7,\'История рекламных расходов\'!$D:$D,"<="&' + colLetter + '8,\'История рекламных расходов\'!$B:$B,"*"&$A$6&"*")))']);
    formulas.push(['=IFERROR(' + colLetter + '46/' + colLetter + '9,0)']);
    formulas.push(['=IFERROR(' + colLetter + '46/' + colLetter + '30,0)']);
    formulas.push(['=IFERROR(' + colLetter + '47/' + colLetter + '37,0)']);
    formulas.push(['=IFERROR(' + colLetter + '46/(' + colLetter + '9*' + colLetter + '24),0)']);
    formulas.push(['=IFERROR(' + colLetter + '47/(' + colLetter + '13*' + colLetter + '25),0)']);
    formulas.push(['=IFERROR(' + colLetter + '46/' + colLetter + '11,0)']);
    formulas.push(['=IFERROR(' + colLetter + '46/(' + colLetter + '11*#REF!),0)']);
    formulas.push(['=IFERROR(' + colLetter + '46/#REF!,0)']);
    formulas.push(['=IFERROR(' + colLetter + '46/(' + colLetter + '15*' + colLetter + '25),0)']);
    formulas.push(['=IFERROR(' + colLetter + '47/(' + colLetter + '15*' + colLetter + '25),0)']);
    formulas.push(['=IFERROR(' + colLetter + '15/' + colLetter + '37,0)']);
    return formulas;
  }

  /**
  * Получить формулы для строк 9–58 недельного столбца (для обновления при переформатировании)
  * @param {Sheet} sheet - Лист
  * @param {number} col - Номер столбца
  * @return {Array.<Array.<string>>} Двумерный массив 50x1 (формулы для строк 9–58) или null при ошибке
  */
  function getWeeklyColumnFormulas(sheet, col) {
    var colLetter = getColumnLetter(col);
    var startDateCell = sheet.getRange(7, col).getValue();
    var endDateCell = sheet.getRange(8, col).getValue();
    var startDateObj = parseCellDate(startDateCell);
    var endDateObj = parseCellDate(endDateCell);
    if (!startDateObj || !endDateObj) {
      return null;
    }
    var lastCol = sheet.getLastColumn();
    var datesRow = sheet.getRange(8, 2, 1, lastCol - 1).getValues()[0];
    var weekStartCol = null;
    var weekEndCol = null;
    for (var d = 0; d < datesRow.length; d++) {
      var cellValue = datesRow[d];
      if (!cellValue || cellValue === '') continue;
      var cellDateObj = parseCellDate(cellValue);
      if (!cellDateObj) continue;
      if (cellDateObj >= startDateObj && cellDateObj <= endDateObj) {
        if (weekStartCol === null) weekStartCol = d + 2;
        weekEndCol = d + 2;
      }
    }
    if (weekStartCol === null || weekEndCol === null) {
      weekStartCol = col;
      weekEndCol = col;
    }
    var weekStartColLetter = getColumnLetter(weekStartCol);
    var weekEndColLetter = getColumnLetter(weekEndCol);
    var formulas = [];
    formulas.push(['=AVERAGE(' + weekStartColLetter + '9:' + weekEndColLetter + '9)']);
    formulas.push(['']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '11:' + weekEndColLetter + '11)']);
    formulas.push(['']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '13:' + weekEndColLetter + '13)']);
    formulas.push(['']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '15:' + weekEndColLetter + '15)']);
    formulas.push(['=IFERROR(' + colLetter + '11/' + colLetter + '15,0)']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '17:' + weekEndColLetter + '17)']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '18:' + weekEndColLetter + '18)']);
    formulas.push(['']);
    formulas.push(['']);
    formulas.push(['']);
    formulas.push(['']);
    formulas.push(['']);
    formulas.push(['']);
    formulas.push(['=SUM(' + weekStartColLetter + '17:' + weekEndColLetter + '17)/SUM(' + weekStartColLetter + '9:' + weekEndColLetter + '9)']);
    formulas.push(['=' + colLetter + '18/' + colLetter + '13']);
    formulas.push(['=IFERROR(' + colLetter + '30/' + colLetter + '28,0)']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '27:' + weekEndColLetter + '27)']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '28:' + weekEndColLetter + '28)']);
    formulas.push(['']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '30:' + weekEndColLetter + '30)']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '31:' + weekEndColLetter + '31)']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '32:' + weekEndColLetter + '32)']);
    formulas.push(['=IFERROR(' + colLetter + '43/' + colLetter + '58,0)']);
    formulas.push(['=IFERROR(' + colLetter + '58/' + colLetter + '57,0)']);
    formulas.push(['=IFERROR(' + colLetter + '58/' + colLetter + '37,0)']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '36:' + weekEndColLetter + '36)']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '37:' + weekEndColLetter + '37)']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '38:' + weekEndColLetter + '38)']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '39:' + weekEndColLetter + '39)']);
    formulas.push(['=IFERROR(' + colLetter + '53/' + colLetter + '58,0)']);
    formulas.push(['=IFERROR(' + colLetter + '37/' + colLetter + '36,0)']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '42:' + weekEndColLetter + '42)']);
    formulas.push(['=IFERROR(' + colLetter + '42/(' + colLetter + '27+' + colLetter + '30),0)']);
    formulas.push(['=IFERROR(' + colLetter + '9/' + colLetter + '42,0)']);
    formulas.push(['']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '46:' + weekEndColLetter + '46)']);
    formulas.push(['=AVERAGE(' + weekStartColLetter + '47:' + weekEndColLetter + '47)']);
    formulas.push(['=IFERROR(' + colLetter + '46/' + colLetter + '9,0)']);
    formulas.push(['=IFERROR(' + colLetter + '46/' + colLetter + '30,0)']);
    formulas.push(['=IFERROR(' + colLetter + '47/' + colLetter + '37,0)']);
    formulas.push(['=IFERROR(' + colLetter + '46/(' + colLetter + '9*#REF!),0)']);
    formulas.push(['=IFERROR(' + colLetter + '47/(' + colLetter + '13*#REF!),0)']);
    formulas.push(['=IFERROR(' + colLetter + '46/' + colLetter + '11,0)']);
    formulas.push(['=IFERROR(' + colLetter + '46/(' + colLetter + '11*#REF!),0)']);
    formulas.push(['=IFERROR(' + colLetter + '46/#REF!,0)']);
    formulas.push(['=IFERROR(' + colLetter + '46/(' + colLetter + '15*#REF!),0)']);
    formulas.push(['=IFERROR(' + colLetter + '47/(' + colLetter + '15*#REF!),0)']);
    formulas.push(['=IFERROR(' + colLetter + '15/' + colLetter + '37,0)']);
    return formulas;
  }

  /**
  * Найти позицию для вставки недельного столбца по дате окончания
  * @param {Sheet} sheet - Лист
  * @param {string} weekStartDate - Дата начала недели в формате YYYY-MM-DD
  * @param {string} weekEndDate - Дата окончания недели в формате YYYY-MM-DD
  * @return {number} Позиция для вставки (или -1 если такая неделя уже существует)
  */
  function findWeekColumnPosition(sheet, weekStartDate, weekEndDate) {
    var lastCol = sheet.getLastColumn();
    
    // Если только столбец A (заголовки), добавляем в B
    if (lastCol === 1) {
      return 2;
    }
    
    // Читаем даты начала (строка 7) и окончания (строка 8)
    var startDatesRow = sheet.getRange(7, 2, 1, lastCol - 1).getValues()[0];
    var endDatesRow = sheet.getRange(8, 2, 1, lastCol - 1).getValues()[0];
    
    // Преобразуем даты для сравнения
    var newStartDateObj = parseDateString(weekStartDate);
    var newEndDateObj = parseDateString(weekEndDate);
    
    var lastMatchingDayCol = null; // Последний ДНЕВНОЙ столбец с той же датой окончания
    var insertBeforeCol = null; // Столбец, перед которым нужно вставить
    
    // Проходим весь цикл, чтобы найти все совпадающие столбцы
    for (var i = 0; i < endDatesRow.length; i++) {
      var startCellValue = startDatesRow[i];
      var endCellValue = endDatesRow[i];
      
      // Пропускаем пустые ячейки
      if (!endCellValue || endCellValue === '') {
        continue;
      }
      
      // Получаем даты из ячеек
      var existingStartDateObj = parseCellDate(startCellValue);
      var existingEndDateObj = parseCellDate(endCellValue);
      
      if (!existingEndDateObj || !existingStartDateObj) {
        continue;
      }
      
      // Проверяем, не является ли это той же самой неделей (недельный столбец)
      if (compareDates(newStartDateObj, existingStartDateObj) === 0 && 
          compareDates(newEndDateObj, existingEndDateObj) === 0) {
        // Точно такая же неделя уже существует
        Logger.log('Недельный столбец с такими же датами уже существует в позиции ' + (i + 2));
        return -1;
      }
      
      // Сравниваем по дате окончания
      var comparisonEnd = compareDates(newEndDateObj, existingEndDateObj);
      
      Logger.log('Сравнение: Новая_Конец=' + weekEndDate + ' с Сущ_Конец=' + Utilities.formatDate(existingEndDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd') + ' Результат=' + comparisonEnd);

      if (comparisonEnd === 0) {
        // Дата окончания совпадает
        // Проверяем, это дневной столбец (start == end) или недельный (start < end)
        var isDailyColumn = compareDates(existingStartDateObj, existingEndDateObj) === 0;
        
        if (isDailyColumn) {
          // Это дневной столбец (воскресенье) - запоминаем позицию
          lastMatchingDayCol = i + 2;
          Logger.log('!!! НАЙДЕНО ВОСКРЕСЕНЬЕ в позиции ' + (i + 2));
        }
      } else if (comparisonEnd < 0 && insertBeforeCol === null) {
        // Новая дата раньше существующей - запоминаем позицию для вставки
        insertBeforeCol = i + 2;
      }
    }
    
    // Если нашли дневной столбец-воскресенье, вставляем ПОСЛЕ него
    if (lastMatchingDayCol !== null) {
      var position = lastMatchingDayCol + 1;
      Logger.log('Вставка после дневного столбца-воскресенья: позиция ' + position);
      return position;
    }
    
    // Если нашли столбец с более поздней датой, вставляем перед ним
    if (insertBeforeCol !== null) {
      Logger.log('Вставка перед столбцом ' + insertBeforeCol);
      return insertBeforeCol;
    }
    
    // Новая дата позже всех существующих - добавляем в конец
    Logger.log('Вставка в конец, позиция ' + (lastCol + 1));
    return lastCol + 1;
  }

  /**
  * Добавить недельный столбец в лист "Воронка динамика"
  * @param {string} inputDateStr - Любая дата недели в формате YYYY-MM-DD
  * @return {Object} Результат операции
  */
  function addSalesFunnelDynamicWeekColumn(inputDateStr) {
    try {
      Logger.log('Добавление недельного столбца для даты: ' + inputDateStr);
      
      // Определяем понедельник и воскресенье для введённой даты
      var inputDate = parseDateString(inputDateStr);
      var dayOfWeek = inputDate.getDay(); // 0=вс, 1=пн, ..., 6=сб
      
      // Вычисляем понедельник этой недели
      var daysFromMonday = (dayOfWeek === 0) ? 6 : dayOfWeek - 1;
      var startDate = new Date(inputDate);
      startDate.setDate(inputDate.getDate() - daysFromMonday);
      
      // Вычисляем воскресенье = понедельник + 6 дней
      var endDate = new Date(startDate);
      endDate.setDate(startDate.getDate() + 6);
      
      var weekStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var weekEndDate = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      Logger.log('Период недели: ' + weekStartDate + ' (пн) - ' + weekEndDate + ' (вс)');
      
      // Получаем или создаем лист
      var sheetName = getSalesFunnelDynamicSheetName();
      var sheet = getOrCreateSheet(sheetName);
      
      // Проверяем, что лист инициализирован
      if (sheet.getLastRow() < 50) {
        var errorMsg = 'Лист "Воронка динамика" не инициализирован. Должно быть минимум 50 строк.';
        Logger.log('ОШИБКА: ' + errorMsg);
        throw new Error(errorMsg);
      }
      
      // Находим правильную позицию для вставки недельного столбца
      var insertPosition = findWeekColumnPosition(sheet, weekStartDate, weekEndDate);
      
      if (insertPosition === -1) {
        Logger.log('Недельный столбец для недели ' + weekStartDate + ' - ' + weekEndDate + ' уже существует. Пропускаем.');
        return { success: true, message: 'Столбец уже существует', skipped: true };
      }
      
      var currentLastCol = sheet.getLastColumn();
      Logger.log('Текущее количество столбцов: ' + currentLastCol);
      Logger.log('Позиция для вставки недельного столбца: ' + insertPosition);
      
      // Если позиция вставки больше текущего количества столбцов, 
      // значит вставляем в самый конец
      if (insertPosition > currentLastCol) {
        Logger.log('Добавление столбца в конец листа (позиция ' + insertPosition + ')');
        // В Google Sheets при записи в getLastColumn() + 1 столбец добавится сам
      } else {
        // Если позиция в середине — раздвигаем столбцы
        sheet.insertColumnBefore(insertPosition);
        Logger.log('Вставлен новый пустой столбец в позицию ' + insertPosition);
      }
      
      // Заполняем столбец недельными данными
      addWeeklyColumn(sheet, insertPosition, weekStartDate, weekEndDate);
      
      Logger.log('Недельный столбец успешно добавлен для недели ' + weekStartDate + ' - ' + weekEndDate);
      
      return { success: true, message: 'Недельный столбец успешно добавлен', skipped: false };
      
    } catch (error) {
      Logger.log('ОШИБКА при добавлении недельного столбца: ' + error.toString());
      Logger.log('Стек ошибки: ' + error.stack);
      throw error;
    }
  }

  /**
  * Добавить столбец с недельной статистикой
  * @param {Sheet} sheet - Лист
  * @param {number} col - Номер столбца
  * @param {string} startDate - Дата начала недели в формате YYYY-MM-DD
  * @param {string} endDate - Дата окончания недели в формате YYYY-MM-DD
  */
  function addWeeklyColumn(sheet, col, startDate, endDate) {
    Logger.log('Добавление недельного столбца: колонка=' + col + ', период=' + startDate + ' - ' + endDate);
    
    // Преобразуем даты в американский формат M/D/YYYY для Google Sheets
    var usStartDate = formatDateUs(startDate);
    var usEndDate = formatDateUs(endDate);
    
    // Преобразуем даты в объекты Date для сравнения
    var startDateObj = parseDateString(startDate);
    var endDateObj = parseDateString(endDate);
    
    // Получаем букву столбца
    var colLetter = getColumnLetter(col);
    
    // Вычисляем диапазон столбцов для усреднения (7 дней)
    // Ищем столбцы с датами, входящими в неделю
    var weekStartCol = null;
    var weekEndCol = null;
    
    // Читаем все даты из строки 8 (Дата до) для поиска дневных столбцов
    var lastCol = sheet.getLastColumn();
    var datesRow = sheet.getRange(8, 2, 1, lastCol - 1).getValues()[0];
    
    for (var i = 0; i < datesRow.length; i++) {
      var cellValue = datesRow[i];
      if (!cellValue || cellValue === '') continue;
      
      var cellDateObj = parseCellDate(cellValue);
      if (!cellDateObj) continue;
      
      // Проверяем, входит ли дата в диапазон недели
      if (cellDateObj >= startDateObj && cellDateObj <= endDateObj) {
        if (weekStartCol === null) {
          weekStartCol = i + 2; // +2 потому что начинаем с B
        }
        weekEndCol = i + 2;
      }
    }
    
    // Если не нашли дневные столбцы, используем текущий столбец
    if (weekStartCol === null || weekEndCol === null) {
      weekStartCol = col;
      weekEndCol = col;
      Logger.log('Не найдены дневные столбцы для недели, используем текущий столбец');
    }
    
    var weekStartColLetter = getColumnLetter(weekStartCol);
    var weekEndColLetter = getColumnLetter(weekEndCol);
    
    Logger.log('Диапазон для усреднения: ' + weekStartColLetter + ' - ' + weekEndColLetter);
    
    // Массив для хранения всех значений столбца (58 строк — как в дневном)
    var columnData = [];
    
    // Строки 1-6: служебная информация (пустые)
    for (var i = 1; i <= 6; i++) {
      columnData.push(['']);
    }
    
    // Строка 7: Дата с (начальная дата) - в формате M/D/YYYY
    columnData.push([usStartDate]);
    
    // Строка 8: Дата до (конечная дата) - в формате M/D/YYYY
    columnData.push([usEndDate]);
    
    // Строка 9: Среднее заказов по артикулу за неделю
    columnData.push(['=AVERAGE(' + weekStartColLetter + '9:' + weekEndColLetter + '9)']);
    
    // Строка 10: Динамика заказов (пусто для недельных)
    columnData.push(['']);
    
    // Строка 11: Среднее суммы заказов по артикулу за неделю
    columnData.push(['=AVERAGE(' + weekStartColLetter + '11:' + weekEndColLetter + '11)']);
    
    // Строка 12: Динамика суммы заказов (пусто)
    columnData.push(['']);
    
    // Строка 13: Среднее заказов по модели за неделю
    columnData.push(['=AVERAGE(' + weekStartColLetter + '13:' + weekEndColLetter + '13)']);
    
    // Строка 14: Динамика заказов по модели (пусто)
    columnData.push(['']);
    
    // Строка 15: Среднее суммы заказов по модели за неделю
    columnData.push(['=AVERAGE(' + weekStartColLetter + '15:' + weekEndColLetter + '15)']);
    
    // Строка 16: Доля заказов артикула в группе по модели
    columnData.push(['=IFERROR(' + colLetter + '11/' + colLetter + '15,0)']);
    
    // Строка 17: Среднее выкупов по артикулу за неделю
    columnData.push(['=AVERAGE(' + weekStartColLetter + '17:' + weekEndColLetter + '17)']);
    
    // Строка 18: Среднее выкупов по модели за неделю
    columnData.push(['=AVERAGE(' + weekStartColLetter + '18:' + weekEndColLetter + '18)']);
    
    // Строка 19: Сумма выкупов по артикулу (пусто)
    columnData.push(['']);
    
    // Строка 20: Динамика выкупы артикул (пусто)
    columnData.push(['']);
    
    // Строка 21: Динамика выкупы модели (пусто)
    columnData.push(['']);
    
    // Строка 22: Возвраты (пусто)
    columnData.push(['']);
    
    // Строка 23: Сумма Возвратов (пусто)
    columnData.push(['']);
    
    // Строка 24: % выкупа по артикулу за неделю
    columnData.push(['=SUM(' + weekStartColLetter + '17:' + weekEndColLetter + '17)/SUM(' + weekStartColLetter + '9:' + weekEndColLetter + '9)']);
    
    // Строка 25: % выкупа по модели за неделю
    columnData.push(['=' + colLetter + '18/' + colLetter + '13']);
    
    // Строка 26: CTR артикула
    columnData.push(['=IFERROR(' + colLetter + '30/' + colLetter + '28,0)']);
    
    // Строка 27: Среднее переходов в карточку за неделю
    columnData.push(['=AVERAGE(' + weekStartColLetter + '27:' + weekEndColLetter + '27)']);
    
    // Строка 28: Среднее показов рекламы по цвету за неделю
    columnData.push(['=AVERAGE(' + weekStartColLetter + '28:' + weekEndColLetter + '28)']);
    
    // Строка 29: Динамика показов (пусто)
    columnData.push(['']);
    
    // Строка 30: Среднее кликов рекламы по цвету за неделю
    columnData.push(['=AVERAGE(' + weekStartColLetter + '30:' + weekEndColLetter + '30)']);

    // --- Новые строки: среднее из дневных столбцов (соответствуют дневным 31-35) ---
    // Строка 31: Среднее показов РК по артикулу за неделю (из дневных стр. 31)
    columnData.push(['=AVERAGE(' + weekStartColLetter + '31:' + weekEndColLetter + '31)']);
    // Строка 32: Среднее кликов РК по артикулу за неделю (из дневных стр. 32)
    columnData.push(['=AVERAGE(' + weekStartColLetter + '32:' + weekEndColLetter + '32)']);
    // Строка 33: (43/58 в недельном столбце)
    columnData.push(['=IFERROR(' + colLetter + '43/' + colLetter + '58,0)']);
    // Строка 34: (58/57)
    columnData.push(['=IFERROR(' + colLetter + '58/' + colLetter + '57,0)']);
    // Строка 35: (58/37)
    columnData.push(['=IFERROR(' + colLetter + '58/' + colLetter + '37,0)']);

    // Строка 36: Среднее показов рекламы по модели за неделю (из дневных стр. 36)
    columnData.push(['=AVERAGE(' + weekStartColLetter + '36:' + weekEndColLetter + '36)']);
    // Строка 37: Среднее кликов рекламы по модели за неделю (из дневных стр. 37)
    columnData.push(['=AVERAGE(' + weekStartColLetter + '37:' + weekEndColLetter + '37)']);

    // --- Новые строки: среднее из дневных (соответствуют дневным 38-40) ---
    // Строка 38: Среднее показов РК по модели за неделю (из дневных стр. 38)
    columnData.push(['=AVERAGE(' + weekStartColLetter + '38:' + weekEndColLetter + '38)']);
    // Строка 39: Среднее кликов РК по модели за неделю (из дневных стр. 39)
    columnData.push(['=AVERAGE(' + weekStartColLetter + '39:' + weekEndColLetter + '39)']);
    // Строка 40: (53/58 в недельном столбце)
    columnData.push(['=IFERROR(' + colLetter + '53/' + colLetter + '58,0)']);

    // Строка 41: CTR модели
    columnData.push(['=IFERROR(' + colLetter + '37/' + colLetter + '36,0)']);
    // Строка 42: Среднее корзины за неделю (из дневных стр. 42)
    columnData.push(['=AVERAGE(' + weekStartColLetter + '42:' + weekEndColLetter + '42)']);
    // Строка 43: Конверсия в корзину
    columnData.push(['=IFERROR(' + colLetter + '42/(' + colLetter + '27+' + colLetter + '30),0)']);
    // Строка 44: Конверсия в заказ
    columnData.push(['=IFERROR(' + colLetter + '9/' + colLetter + '42,0)']);
    // Строка 45: Общий коэффициент (пусто)
    columnData.push(['']);
    // Строка 46: Среднее затрат на рекламу по артикулу за неделю (из дневных стр. 46)
    columnData.push(['=AVERAGE(' + weekStartColLetter + '46:' + weekEndColLetter + '46)']);
    // Строка 47: Среднее затрат на рекламу по модели за неделю (из дневных стр. 47)
    columnData.push(['=AVERAGE(' + weekStartColLetter + '47:' + weekEndColLetter + '47)']);
    // Строка 48: CPO
    columnData.push(['=IFERROR(' + colLetter + '46/' + colLetter + '9,0)']);
    // Строка 49: CPC по артикулу
    columnData.push(['=IFERROR(' + colLetter + '46/' + colLetter + '30,0)']);
    // Строка 50: CPC по модели
    columnData.push(['=IFERROR(' + colLetter + '47/' + colLetter + '37,0)']);
    // Строка 51: CPS по артикулу
    columnData.push(['=IFERROR(' + colLetter + '46/(' + colLetter + '9*#REF!),0)']);
    // Строка 52: CPS по модели
    columnData.push(['=IFERROR(' + colLetter + '47/(' + colLetter + '13*#REF!),0)']);
    // Строка 53: ДРР фактическая от заказа цвета на цвет
    columnData.push(['=IFERROR(' + colLetter + '46/' + colLetter + '11,0)']);
    // Строка 54: ДРР вмененная от выкупа цвета на цвет
    columnData.push(['=IFERROR(' + colLetter + '46/(' + colLetter + '11*#REF!),0)']);
    // Строка 55: ДРР фактическая от выкупа цвет на цвет
    columnData.push(['=IFERROR(' + colLetter + '46/#REF!,0)']);
    // Строка 56: ДРР вмененная от выкупа цвета на модель
    columnData.push(['=IFERROR(' + colLetter + '46/(' + colLetter + '15*#REF!),0)']);
    // Строка 57: ДРР вмененная всей модели
    columnData.push(['=IFERROR(' + colLetter + '47/(' + colLetter + '15*#REF!),0)']);
    // Строка 58: Заказов на 1 клик, руб (по модели)
    columnData.push(['=IFERROR(' + colLetter + '15/' + colLetter + '37,0)']);

    // Записываем все данные в столбец одной операцией
    sheet.getRange(1, col, 58, 1).setValues(columnData);
    
    // Применяем формат даты dd.mm.yyyy к строкам 7 и 8
    var dateRange = sheet.getRange(7, col, 2, 1);
    dateRange.setNumberFormat('dd.mm.yyyy');
    
    // Числовой формат для недельного столбца (средние заказы, суммы, показы, клики, затраты и т.д.)
    var numberRowsWeekly = [9, 11, 13, 15, 17, 18, 27, 28, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 42, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58];
    for (var r = 0; r < numberRowsWeekly.length; r++) {
      sheet.getRange(numberRowsWeekly[r], col).setNumberFormat('#,##0.00');
    }
    // Процентный формат для строк с процентами (доля, % выкупа, CTR, конверсия)
    var percentRowsWeekly = [16, 24, 25, 26, 41, 43, 44];
    for (var p = 0; p < percentRowsWeekly.length; p++) {
      sheet.getRange(percentRowsWeekly[p], col).setNumberFormat('0.00%');
    }
    
    // Применяем форматирование: недельный отчет теперь имеет такие же цвета, как дневной (58 строк)
    // Сначала очищаем фон и УБИРАЕМ ВСЕ РАМКИ со всего столбца (с 1 по 58 строку)
    var fullColumnRange = sheet.getRange(1, col, 58, 1);
    fullColumnRange.setBackground('#ffffff');
    fullColumnRange.setBorder(false, false, false, false, false, false);
    
    
    var range7_8 = sheet.getRange(7, col, 2, 1);
    range7_8.setBackground('#c5e0b3');
    // Устанавливаем рамку только для дат (внешнюю и внутреннюю горизонтальную)
    range7_8.setBorder(true, true, true, true, null, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    // Для недельного отчета подкрашиваем ключевые строки как в дневном (58 строк)
    // 9 и 13 строка: ярко-зеленый
    sheet.getRange(9, col).setBackground('#00ff00');
    sheet.getRange(13, col).setBackground('#00ff00');
    
    // 11 строка: светло-голубой
    sheet.getRange(11, col).setBackground('#d0e0e3');
    
    // 17 и 18 строка: средне-зеленый
    sheet.getRange(17, col, 2, 1).setBackground('#92d050');
    
    // 24 и 25 строка: серый
    sheet.getRange(24, col, 2, 1).setBackground('#d9d9d9');
    
    // 27-30: светло-голубой(3)
    sheet.getRange(27, col, 4, 1).setBackground('#d0e0e3');
    
    // 31-35: новые строки РК по артикулу (средние из дневных)
    sheet.getRange(31, col, 5, 1).setBackground('#d9ead3');
    
    // 36-37: среднее показов/кликов по модели
    sheet.getRange(36, col, 2, 1).setBackground('#d9ead3');
    
    // 38-40: новые строки РК по модели (средние из дневных)
    sheet.getRange(38, col, 3, 1).setBackground('#d9ead3');
    
    // 41: CTR модели
    sheet.getRange(41, col).setBackground('#d9ead3');
    
    // 42: светло-голубой(3)
    sheet.getRange(42, col).setBackground('#d0e0e3');
    
    // 43-44: светло-бирюзовый(3)
    sheet.getRange(43, col, 2, 1).setBackground('#b2dfdb');
    
    // 45: общий коэффициент
    sheet.getRange(45, col).setBackground('#d9d9d9');
    
    // 48-49: светло-бирюзовый(3)
    sheet.getRange(48, col, 2, 1).setBackground('#b2dfdb');
    
    // 50: светло-голубой(3)
    sheet.getRange(50, col).setBackground('#d0e0e3');
    
    // 51-52: светло-бирюзовый(3)
    sheet.getRange(51, col, 2, 1).setBackground('#b2dfdb');
    
    // 53: светло-оранжевый(3)
    sheet.getRange(53, col).setBackground('#ffe0b2');
    
    // 54: светло-оранжевый(2)
    sheet.getRange(54, col).setBackground('#ffcc80');
    // 55-56: светло-оранжевый(3)
    sheet.getRange(55, col, 2, 1).setBackground('#ffe0b2');
    
    // 57: светло-зеленый(3)
    sheet.getRange(57, col).setBackground('#b6d7a8');
    
    // Недельный отчет: только внешняя рамка (7-58), без внутренних линий между ячейками. Внутренние горизонтали — только под строками 7 и 8.
    if (col < sheet.getMaxColumns()) {
      sheet.getRange(7, col, 52, 1).setBorder(true, true, true, false, false, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
      sheet.getRange(7, col, 1, 1).setBorder(false, false, true, false, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
      sheet.getRange(8, col, 1, 1).setBorder(false, false, true, false, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    } else {
      sheet.getRange(7, col, 52, 1).setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
      sheet.getRange(7, col, 1, 1).setBorder(false, false, true, false, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
      sheet.getRange(8, col, 1, 1).setBorder(false, false, true, false, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    }
    // -----------------------------------------------------
    
    Logger.log('Недельный столбец ' + colLetter + ' успешно добавлен с форматированием');
  }

  /**
  * Синхронизировать недельную статистику для воронки динамики
  * Добавляет данные за все 7 дней недели и итоговый недельный столбец
  * @param {string} inputDateStr - Любая дата недели в формате YYYY-MM-DD
  * @return {Object} Результат операции
  */
  function syncSalesFunnelDynamicWeek(inputDateStr) {
    try {
      Logger.log('=== Начало синхронизации недельной статистики воронки динамики ===');
      Logger.log('Введённая дата: ' + inputDateStr);
      
      // Определяем понедельник и воскресенье для введённой даты
      var inputDate = parseDateString(inputDateStr);
      var dayOfWeek = inputDate.getDay(); // 0=вс, 1=пн, 2=вт, 3=ср, 4=чт, 5=пт, 6=сб
      
      Logger.log('День недели введённой даты: ' + dayOfWeek + ' (0=вс, 1=пн, ..., 6=сб)');
      
      // Вычисляем понедельник этой недели
      // Если воскресенье (0) - отнимаем 6 дней
      // Если понедельник (1) - отнимаем 0 дней
      // Если вторник (2) - отнимаем 1 день и т.д.
      var daysFromMonday = (dayOfWeek === 0) ? 6 : dayOfWeek - 1;
      var startDate = new Date(inputDate);
      startDate.setDate(inputDate.getDate() - daysFromMonday);
      
      // Вычисляем воскресенье = понедельник + 6 дней
      var endDate = new Date(startDate);
      endDate.setDate(startDate.getDate() + 6);
      
      var weekStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var weekEndDate = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      Logger.log('Период недели: ' + weekStartDate + ' (пн) - ' + weekEndDate + ' (вс)');
      
      // Добавляем столбцы для каждого дня недели
      var currentDate = new Date(startDate);
      var daysAdded = 0;
      
      for (var i = 0; i < 7; i++) {
        var dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        Logger.log('Добавление дневного столбца для: ' + dateStr);
        
        try {
          var result = addSalesFunnelDynamicColumn(dateStr);
          if (!result.skipped) {
            daysAdded++;
          }
        } catch (error) {

          Logger.log('Ошибка при добавлении столбца для ' + dateStr + ': ' + error.toString());
        }
        
        currentDate.setDate(currentDate.getDate() + 1);
      }
      
      Logger.log('Добавлено дневных столбцов: ' + daysAdded);
      
      // Добавляем итоговый недельный столбец
      Logger.log('Добавление недельного итогового столбца');
      var weekResult = addSalesFunnelDynamicWeekColumn(weekEndDate);
      
      var message = 'Синхронизация недели завершена. Добавлено дневных столбцов: ' + daysAdded + 
                    ', недельный столбец: ' + (weekResult.skipped ? 'уже существует' : 'добавлен');
      
      Logger.log('=== ' + message + ' ===');
      
      return { 
        success: true, 
        message: message,
        daysAdded: daysAdded,
        weekColumnAdded: !weekResult.skipped
      };
      
    } catch (error) {
      Logger.log('ОШИБКА при синхронизации недельной статистики: ' + error.toString());
      Logger.log('Стек ошибки: ' + error.stack);
      throw error;
    }
  }

  /**
  * Получить имя листа для воронки динамики
  * @return {string} Имя листа
  */
  function getSalesFunnelDynamicSheetName() {
    return 'Воронка динамика';
  }

  /**
  * Отформатировать весь лист "Воронка динамика" согласно новым правилам дизайна
  */
  function reformatSalesFunnelDynamicSheet() {
    try {
      Logger.log('=== Переформатирование листа "Воронка динамика" ===');
      var sheetName = getSalesFunnelDynamicSheetName();
      var sheet = getSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        Logger.log('Лист "' + sheetName + '" не найден');
        return;
      }
      
      var lastCol = sheet.getLastColumn();
      if (lastCol < 2) {
        Logger.log('Нет столбцов с данными для форматирования');
        return;
      }
      
      // Читаем все даты из строк 7 и 8 для определения типа столбцов
      var startDatesRow = sheet.getRange(7, 2, 1, lastCol - 1).getValues()[0];
      var endDatesRow = sheet.getRange(8, 2, 1, lastCol - 1).getValues()[0];
      var weeklyCols = [];
      
      // Обрабатываем каждый столбец
      for (var i = 0; i < startDatesRow.length; i++) {
        var col = i + 2;
        var startDate = startDatesRow[i];
        var endDate = endDatesRow[i];
        
        var isWeekly = false;
        if (startDate && endDate) {
          if (startDate instanceof Date && endDate instanceof Date) {
            isWeekly = startDate.getTime() !== endDate.getTime();
          } else {
            isWeekly = String(startDate) !== String(endDate);
          }
        }
        if (isWeekly) {
          weeklyCols.push(col);
        }
        
        // 1. Применяем базовый фон и рамки (58 строк — как в дневном/недельном столбце)
        var fullColumnRange = sheet.getRange(1, col, 58, 1);
        fullColumnRange.setBackground('#ffffff');
        fullColumnRange.setBorder(false, false, false, false, false, false);
        
        // --- ПРИМЕНЯЕМ ЦВЕТА (58 строк) ---
        // 2. 7 и 8 строка: цвет c5e0b3 и ЖИРНАЯ РАМКА
        var range7_8 = sheet.getRange(7, col, 2, 1);
        range7_8.setBackground('#c5e0b3');
        range7_8.setBorder(true, true, true, true, null, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        
        // 3. 9 и 13 строка: ярко-зеленый
        sheet.getRange(9, col).setBackground('#00ff00');
        sheet.getRange(13, col).setBackground('#00ff00');
        
        // 4. 11 строка: светло-голубой
        sheet.getRange(11, col).setBackground('#d0e0e3');
        
        // 5. 17 и 18 строка: средне-зеленый
        sheet.getRange(17, col, 2, 1).setBackground('#92d050');
        
        // 6. 24 и 25 строка: серый
        sheet.getRange(24, col, 2, 1).setBackground('#d9d9d9');
        
        // 7. 27-30: светло-голубой
        sheet.getRange(27, col, 4, 1).setBackground('#d0e0e3');
        
        // 8. 31-35: новые строки РК по артикулу
        sheet.getRange(31, col, 5, 1).setBackground('#d9ead3');
        
        // 9. 36-37: показы/клики по модели
        sheet.getRange(36, col, 2, 1).setBackground('#d9ead3');
        
        // 10. 38-40: новые строки РК по модели
        sheet.getRange(38, col, 3, 1).setBackground('#d9ead3');
        
        // 11. 41: CTR модели
        sheet.getRange(41, col).setBackground('#d9ead3');
        
        // 12. 42: светло-голубой
        sheet.getRange(42, col).setBackground('#d0e0e3');
        
        // 13. 43-44: светло-бирюзовый
        sheet.getRange(43, col, 2, 1).setBackground('#b2dfdb');
        
        // 14. 45: общий коэффициент
        sheet.getRange(45, col).setBackground('#d9d9d9');
        
        // 15. 48-49: светло-бирюзовый
        sheet.getRange(48, col, 2, 1).setBackground('#b2dfdb');
        
        // 16. 50: светло-голубой
        sheet.getRange(50, col).setBackground('#d0e0e3');
        
        // 17. 51-52: светло-бирюзовый
        sheet.getRange(51, col, 2, 1).setBackground('#b2dfdb');
        
        // 18. 53-56: светло-оранжевый
        sheet.getRange(53, col).setBackground('#ffe0b2');
        sheet.getRange(54, col).setBackground('#ffcc80');
        sheet.getRange(55, col, 2, 1).setBackground('#ffe0b2');
        
        // 19. 57: светло-зеленый
        sheet.getRange(57, col).setBackground('#b6d7a8');
        
        // Форматы чисел: дата в 7–8; целые; проценты; рубли; числа с дробной частью (58 строк)
        sheet.getRange(7, col, 8, col).setNumberFormat('dd.mm.yyyy');
        var integerRowsAll = [9, 13, 17, 18, 22, 27, 28, 30, 31, 32, 36, 37, 38, 39, 40, 42, 46, 47, 48, 50, 52, 58];
        for (var r = 0; r < integerRowsAll.length; r++) {
          sheet.getRange(integerRowsAll[r], col).setNumberFormat('#,##0');
        }
        var percentRowsAll = [10, 12, 14, 16, 20, 21, 24, 25, 26, 29, 33, 34, 35, 41, 43, 44];
        for (var p = 0; p < percentRowsAll.length; p++) {
          sheet.getRange(percentRowsAll[p], col).setNumberFormat('0.00%');
        }
        var rubleRowsAll = [11, 15, 19, 23];
        var rubleFormatAll = '#,##0.00" ₽"';
        for (var r = 0; r < rubleRowsAll.length; r++) {
          sheet.getRange(rubleRowsAll[r], col).setNumberFormat(rubleFormatAll);
        }
        var numberRowsAll = [45, 49, 51, 53, 54, 55, 56, 57];
        for (var r = 0; r < numberRowsAll.length; r++) {
          sheet.getRange(numberRowsAll[r], col).setNumberFormat('#,##0.00');
        }
        
        // Дополнительно для недельных: только внешняя рамка (7-58), внутренние горизонтали только под строками 7 и 8
        if (isWeekly) {
          if (col < sheet.getMaxColumns()) {
            sheet.getRange(7, col, 52, 1).setBorder(
              true, true, true, false,
              false, false,
              '#000000',
              SpreadsheetApp.BorderStyle.SOLID_THICK
            );

            // 7 строка — жирная нижняя НА ДВА СТОЛБЦА
            sheet.getRange(7, col, 1, 2).setBorder(
              false, false, true, false,
              false, false,
              '#000000',
              SpreadsheetApp.BorderStyle.SOLID_THICK
            );

            // 8 строка — жирная нижняя НА ДВА СТОЛБЦА
            sheet.getRange(8, col, 1, 2).setBorder(
              false, false, true, false,
              false, false,
              '#000000',
              SpreadsheetApp.BorderStyle.SOLID_THICK
  );

          } else {
            sheet.getRange(7, col, 52, 1).setBorder(
              true, true, true, true,
              false, false,
              '#000000',
              SpreadsheetApp.BorderStyle.SOLID_THICK
            );

            // Последний столбец — 7 и 8 строки только в нём
            sheet.getRange(7, col, 1, 1).setBorder(
              false, false, true, false,
              false, false,
              '#000000',
              SpreadsheetApp.BorderStyle.SOLID_THICK
            );

            sheet.getRange(8, col, 1, 1).setBorder(
              false, false, true, false,
              false, false,
              '#000000',
              SpreadsheetApp.BorderStyle.SOLID_THICK
            );
          }
        }
        
        // Обновляем формулы для столбца (строки 9–58)
        var formulas = isWeekly ? getWeeklyColumnFormulas(sheet, col) : getDailyColumnFormulas(col);
        if (formulas && formulas.length === 50) {
          sheet.getRange(9, col, 50, 1).setFormulas(formulas);
        }
      }
      
      // После цикла при очистке col+1 могла сброситься правая граница недельного столбца — дорисовываем левую границу у col+1
      for (var w = 0; w < weeklyCols.length; w++) {
        var c = weeklyCols[w];
        if (c < sheet.getMaxColumns()) {
          sheet.getRange(7, c + 1, 52, 1).setBorder(false, true, false, false, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
        }
      }
      
      Logger.log('Переформатирование успешно завершено');
      return { success: true, message: 'Лист успешно отформатирован' };
      
    } catch (error) {
      Logger.log('ОШИБКА при переформатировании: ' + error.toString());
      throw error;
    }
  }
