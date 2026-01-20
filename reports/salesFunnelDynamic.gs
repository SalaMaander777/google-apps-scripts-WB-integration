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
    
    addSalesFunnelDynamicColumn(reportDate);
    
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
 * @return {Date} Объект Date
 */
function parseDateString(dateStr) {
  var parts = dateStr.split('-');
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

/**
 * Преобразовать значение ячейки в объект Date
 * @param {*} cellValue - Значение ячейки
 * @return {Date|null} Объект Date или null
 */
function parseCellDate(cellValue) {
  if (cellValue instanceof Date) {
    return cellValue;
  }
  
  if (typeof cellValue === 'string') {
    // Формат DD.MM.YYYY
    var parts = cellValue.split('.');
    if (parts.length === 3) {
      return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    
    // Формат YYYY-MM-DD
    parts = cellValue.split('-');
    if (parts.length === 3) {
      return new Date(parts[0], parts[1] - 1, parts[2]);
    }
  }
  
  return null;
}

/**
 * Сравнить две даты
 * @param {Date} date1 - Первая дата
 * @param {Date} date2 - Вторая дата
 * @return {number} -1 если date1 < date2, 0 если равны, 1 если date1 > date2
 */
function compareDates(date1, date2) {
  var time1 = date1.getTime();
  var time2 = date2.getTime();
  
  if (time1 < time2) return -1;
  if (time1 > time2) return 1;
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
  
  // Преобразуем дату в формат DD.MM.YYYY для отображения
  var displayDate = formatDateRu(date);
  
  // Получаем букву столбца для использования в формулах
  var colLetter = getColumnLetter(col);
  
  // Массив для хранения всех значений столбца (50 строк)
  var columnData = [];
  
  // Строки 1-6: служебная информация (пустые для нового столбца)
  for (var i = 1; i <= 6; i++) {
    columnData.push(['']);
  }
  
  // Строка 7: Дата с (начальная дата)
  columnData.push([displayDate]);
  
  // Строка 8: Дата до (конечная дата)
  columnData.push([displayDate]);
  
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
  columnData.push(['=SUMIFS(\'Лента заказов\'!$H:$H,\'Лента заказов\'!$E:$E,"*"&$A$4&"*",\'Лента заказов\'!$K:$K,">="&' + colLetter + '7,\'Лента заказов\'!$K:$K,"<="&' + colLetter + '8)']);
  
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
  
  // Строка 31: Показы реклама по модели
  columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$J:$J,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$J:$J,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
  
  // Строка 32: Клики реклама по модели
  columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'Аналитика РК\'!$K:$K,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8)),SUMIFS(\'Аналитика РК\'!$K:$K,\'Аналитика РК\'!$L:$L,"*"&$A$4&"*",\'Аналитика РК\'!$A:$A,">="&' + colLetter + '7,\'Аналитика РК\'!$A:$A,"<="&' + colLetter + '8,\'Аналитика РК\'!$D:$D,"*"&$A$6&"*"))']);
  
  // Строка 33: CTR модели
  columnData.push(['=IFERROR(' + colLetter + '32/' + colLetter + '31,0)']);
  
  // Строка 34: Корзина
  columnData.push(['=SUMIFS(\'Аналитика продавца\'!$K:$K,\'Аналитика продавца\'!$C:$C,$A$1,\'Аналитика продавца\'!$A:$A,">="&' + colLetter + '7,\'Аналитика продавца\'!$A:$A,"<="&' + colLetter + '8)']);
  
  // Строка 35: Конверсия в корзину
  columnData.push(['=IFERROR(' + colLetter + '34/(' + colLetter + '27+' + colLetter + '30),0)']);
  
  // Строка 36: Конверсия в заказ
  columnData.push(['=IFERROR(' + colLetter + '9/' + colLetter + '34,0)']);
  
  // Строка 37: Общий коэффициент
  columnData.push(['=' + colLetter + '26*' + colLetter + '35*' + colLetter + '36*1000']);
  
  // Строка 38: Затраты на рекламу по артикулу
  columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'История рекламных расходов\'!$G:$G,\'История рекламных расходов\'!$I:$I,$A$2,\'История рекламных расходов\'!$D:$D,">="&' + colLetter + '7,\'История рекламных расходов\'!$D:$D,"<="&' + colLetter + '8)),(SUMIFS(\'История рекламных расходов\'!$G:$G,\'История рекламных расходов\'!$I:$I,$A$2,\'История рекламных расходов\'!$D:$D,">="&' + colLetter + '7,\'История рекламных расходов\'!$D:$D,"<="&' + colLetter + '8,\'История рекламных расходов\'!$B:$B,"*"&$A$6&"*")))']);
  
  // Строка 39: Затраты на рекламу по модели
  columnData.push(['=IF($A$6="ВСЕ",(SUMIFS(\'История рекламных расходов\'!$G:$G,\'История рекламных расходов\'!$I:$I,"*"&$A$4&"*",\'История рекламных расходов\'!$D:$D,">="&' + colLetter + '7,\'История рекламных расходов\'!$D:$D,"<="&' + colLetter + '8)),(SUMIFS(\'История рекламных расходов\'!$G:$G,\'История рекламных расходов\'!$I:$I,"*"&$A$4&"*",\'История рекламных расходов\'!$D:$D,">="&' + colLetter + '7,\'История рекламных расходов\'!$D:$D,"<="&' + colLetter + '8,\'История рекламных расходов\'!$B:$B,"*"&$A$6&"*")))']);
  
  // Строка 40: CPO
  columnData.push(['=IFERROR(' + colLetter + '38/' + colLetter + '9,0)']);
  
  // Строка 41: CPC по артикулу
  columnData.push(['=IFERROR(' + colLetter + '38/' + colLetter + '30,0)']);
  
  // Строка 42: CPC по модели
  columnData.push(['=IFERROR(' + colLetter + '39/' + colLetter + '32,0)']);
  
  // Строка 43: CPS по артикулу
  columnData.push(['=IFERROR(' + colLetter + '38/(' + colLetter + '9*' + colLetter + '24),0)']);
  
  // Строка 44: CPS по модели
  columnData.push(['=IFERROR(' + colLetter + '39/(' + colLetter + '13*' + colLetter + '25),0)']);
  
  // Строка 45: ДРР фактическая от заказа цвета на цвет
  columnData.push(['=IFERROR(' + colLetter + '38/' + colLetter + '11,0)']);
  
  // Строка 46: ДРР вмененная от выкупа цвета на цвет
  columnData.push(['=IFERROR(' + colLetter + '38/(' + colLetter + '11*#REF!),0)']);
  
  // Строка 47: ДРР фактическая от выкупа цвет на цвет
  columnData.push(['=IFERROR(' + colLetter + '38/#REF!,0)']);
  
  // Строка 48: ДРР вмененная от выкупа цвета на модель
  columnData.push(['=IFERROR(' + colLetter + '38/(' + colLetter + '15*' + colLetter + '25),0)']);
  
  // Строка 49: ДРР вмененная всей модели
  columnData.push(['=IFERROR(' + colLetter + '39/(' + colLetter + '15*' + colLetter + '25),0)']);
  
  // Строка 50: Заказов на 1 клик, руб (по модели)
  columnData.push(['=IFERROR(' + colLetter + '15/' + colLetter + '32,0)']);
  
  // Записываем все данные в столбец одной операцией
  sheet.getRange(1, col, 50, 1).setValues(columnData);
  
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
 * Получить имя листа для воронки динамики
 * @return {string} Имя листа
 */
function getSalesFunnelDynamicSheetName() {
  return 'Воронка динамика';
}
