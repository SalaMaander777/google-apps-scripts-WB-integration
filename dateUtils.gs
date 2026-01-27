/**
 * Утилиты для работы с датами
 */

/**
 * Получить дату предыдущего дня в формате RFC3339 (Москва UTC+3)
 * @return {string} Дата в формате YYYY-MM-DD
 */
function getPreviousDay() {
  // Используем часовой пояс Москвы для правильного расчета предыдущего дня
  var moscowTimezone = 'Europe/Moscow';
  var now = new Date();
  
  // Получаем текущую дату в часовом поясе Москвы
  var moscowDateStr = Utilities.formatDate(now, moscowTimezone, 'yyyy-MM-dd');
  var moscowDateParts = moscowDateStr.split('-');
  var moscowDate = new Date(moscowDateParts[0], moscowDateParts[1] - 1, moscowDateParts[2]);
  
  // Вычитаем один день
  moscowDate.setDate(moscowDate.getDate() - 1);
  
  var year = moscowDate.getFullYear();
  var month = String(moscowDate.getMonth() + 1).padStart(2, '0');
  var day = String(moscowDate.getDate()).padStart(2, '0');
  
  return year + '-' + month + '-' + day;
}

/**
 * Получить дату начала и конца предыдущего дня в формате RFC3339
 * API Wildberries требует время в часовом поясе Москва (UTC+3)
 * @return {Object} Объект с dateFrom и dateTo
 */
function getPreviousDayRange() {
  var date = getPreviousDay();
  // Формат для API: дата со временем в часовом поясе Москва
  return {
    dateFrom: date + 'T00:00:00+03:00',
    dateTo: date + 'T23:59:59+03:00'
  };
}

/**
 * Преобразовать дату в формат для Google Sheets
 * @param {string} dateStr - Дата в формате YYYY-MM-DD или ISO
 * @return {Date} Объект Date
 */
function parseDate(dateStr) {
  if (!dateStr) return null;
  // Убираем время если есть
  var dateOnly = dateStr.split('T')[0];
  var parts = dateOnly.split('-');
  if (parts.length === 3) {
    return new Date(parts[0], parts[1] - 1, parts[2]);
  }
  return new Date(dateStr);
}

/**
 * Получить дату год назад от указанной даты
 * @param {string} dateStr - Дата в формате YYYY-MM-DD
 * @return {string} Дата год назад в формате YYYY-MM-DD
 */
function getDateYearAgo(dateStr) {
  var parts = dateStr.split('-');
  if (parts.length !== 3) {
    throw new Error('Неверный формат даты. Ожидается YYYY-MM-DD');
  }
  
  var year = parseInt(parts[0], 10);
  var month = parts[1];
  var day = parts[2];
  
  // Вычитаем один год
  var yearAgo = year - 1;
  
  return yearAgo + '-' + month + '-' + day;
}

/**
 * Форматировать дату в русский формат DD.MM.YYYY
 * @param {string} dateStr - Дата в формате YYYY-MM-DD
 * @return {string} Дата в формате DD.MM.YYYY
 */
function formatDateRu(dateStr) {
  if (!dateStr) return '';
  
  var parts = dateStr.split('-');
  if (parts.length !== 3) {
    return dateStr;
  }
  
  var year = parts[0];
  var month = parts[1];
  var day = parts[2];
  
  return day + '.' + month + '.' + year;
}

/**
 * Форматировать дату в американский формат M/D/YYYY (для Google Sheets)
 * @param {string} dateStr - Дата в формате YYYY-MM-DD
 * @return {string} Дата в формате M/D/YYYY
 */
function formatDateUs(dateStr) {
  if (!dateStr) return '';
  
  var parts = dateStr.split('-');
  if (parts.length !== 3) {
    return dateStr;
  }
  
  var year = parts[0];
  var month = parseInt(parts[1], 10); // убираем ведущий ноль
  var day = parseInt(parts[2], 10);   // убираем ведущий ноль
  
  return month + '/' + day + '/' + year;
}

/**
 * Форматировать дату в американский формат MM/DD/YYYY (с ведущими нулями)
 * @param {string} dateStr - Дата в формате YYYY-MM-DD или ISO
 * @return {string} Дата в формате MM/DD/YYYY
 */
function formatDateMMDDYYYY(dateStr) {
  if (!dateStr) return '';
  
  // Убираем время если есть
  var dateOnly = dateStr.split('T')[0];
  var parts = dateOnly.split('-');
  if (parts.length !== 3) {
    return dateStr;
  }
  
  var year = parts[0];
  var month = parts[1]; // сохраняем ведущий ноль
  var day = parts[2];   // сохраняем ведущий ноль
  
  return month + '/' + day + '/' + year;
}
