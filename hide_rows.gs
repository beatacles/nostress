//функция сворачивания пустых строк
function hide_rows(name,array,start) {
  let sh = SpreadsheetApp.getActiveSpreadsheet(); // активировать данную таблицу
  let ss = sh.getSheetByName(name);//взять вкладку по названию, где name - это название, например "Р№"
  let r= ss.getRange(array);//по какому массиву будем искать пустые строки, где array - это массив, например "A5:A54"
  let data = r.getValues();//значения массива
  ss.unhideRow(r);
  //цикл для скрытия строк
  let i=0;
  while( data[i].filter(String).length != 0 ){
    i++;
  };
  ss.hideRows(i+start,data.length-i);
};