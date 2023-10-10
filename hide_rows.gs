//Функция сворачивания пустых строк
function hide_rows(name, array, start) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); // получаем вкладку
  const range = sheet.getRange(array); // получаем диапазон ячеек
  const values = range.getValues(); // получаем значения ячеек
  const lastIndex = values.findIndex(row => !row.join('')); // находим последнюю пустую строку
  sheet.showRows(start, lastIndex); // сбрасываем скрытие строк
  if (lastIndex >= 0) {
    const rowsToHide = values.length - lastIndex; // количество строк для скрытия
    sheet.hideRows(lastIndex + start, rowsToHide); // скрываем строки
  }
};