/**
 * Функции запускаемые при редактировании и открытии
 */
// Функции, запускаемые при редактировании
function OnEdit(e){
  hide_rows('Р№', 'A5:A54',5);
  hide_rows('Р.ВК', 'B10:B259',10);
};

// Функции, запускаемые при открытии
function onOpen(){
  createMenu();
};

/**
 * Front.Создание меню.
 * Создание модального окна для АОСР.
 * Создание модального окна для АВК.
 */

//Создание меню
function createMenu(){
  const ui = SpreadsheetApp.getUi();
  menu = ui.createMenu("Меню");
  menu.addItem("Cоздать АОСР","modalForm_AOSR");
  menu.addItem("Cоздать АВК","modalForm_AVK");
  menu.addToUi();
};

//Создание модального окна АОСР
function modalForm_AOSR(){
  const modalForm = HtmlService.createTemplateFromFile("modal_aosr")
  const htmlOutput = modalForm.evaluate()
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(htmlOutput, "Меню создания АОСР");
};

//Создание модального окна АВК
function modalForm_AVK(){
  const modalForm = HtmlService.createTemplateFromFile("modal_avk")
  const htmlOutput = modalForm.evaluate()
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(htmlOutput, "Меню создания АВК");
};

/**
 * Back.
 * Извлечение данных для выпадающего списка
 * Преобразование данных из формы
 * Создание АОСР
 * Создание реестра к АОСР
 * Создание АВК
 * Создание сопроводительной документации к АВК
 */

//Преобразует данные из форм
function processForm(inputData) {
  console.log("Входные данные:",inputData);
  const name = inputData.name;
  const excep = inputData.excep.length !== 0 ? inputData.excep.split(",").map(item => parseInt(item)) : []; //проверка исключений, перевод из текста в массив чисел
  const acts = Array.from({ length: parseInt(inputData.finish) - parseInt(inputData.start) + 1 }, (_, i) => i + parseInt(inputData.start));//проверка актов, создание массива чисел
  const filteredActs = excep.length !== 0 ? acts.filter(item => !excep.includes(item)) : acts;
  const outputData = { name, acts: filteredActs.reverse() };
  console.log("Выходные данные:",outputData);

  switch (inputData.type) {
    case "aosr":
      multiple_aosr(outputData);
      break;
    case "avk":
      multiple_avk(outputData);
      break;
    default:
      console.log("Ошибка определения типа документа");
  }
};

//Извлекает и подготавливает данные из таблицы по названию именнованного диапазона для выпадающего списка формы
function getDropdownListArray(){
  return SpreadsheetApp.getActiveSpreadsheet()
    .getRangeByName("chapter")
    .getValues()
    .flat()
    .filter(ev => ev !="")
};

//Запуск вывода набора АОСР+реестр
function multiple_aosr(outputData){
  
  //Работа с папками
  nameFolder = outputData.name + "_№"+ outputData.acts[outputData.acts.length - 1]+"-"+outputData.acts[0];//создаем название папки для комлпекта АОСР
  let folderID = "1Hq5o3MsgiSXtEfrhuOBtf0hPtWSZInPs";//папка для складывания папок АОСР
  let folder = DriveApp.getFolderById(folderID);//обращение по id папки
  let newFolderID = folder.createFolder(nameFolder).getId();//создание папки с нужным namefolder
  folder=DriveApp.getFolderById(newFolderID); // взяли ее по id

  //Установка шифра раздела АОСР
  var as = SpreadsheetApp.getActiveSpreadsheet();//активная таблица
  var ss = as.getSheetByName("АОСР");//активный лист АОСР
  var sr = as.getSheetByName("Р№");//активный лист Реестра
  as.getRangeByName("acts.name").setValue(outputData.name);//установили название раздела
  
  //Цикл создания PDF АОСР (4 сек. на один Акт)
  for (i = 0; i<=outputData.acts.length-1; i++ ){
    filename = "АОСР_" + outputData.name + "_№"+ outputData.acts[i];
    as.getRangeByName("acts.number").setValue(outputData.acts[i]);
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    exportPDFtoGDrive(folder, as.getId(),filename,[ [ ss.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,6,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[ss.getSheetId().toString(),[[0,37],[37,82]],[[0,20]]]]
    );
    console.log("Создал АОСР");
  };

  //Цикл создания PDF Реестр к АОСР
  for (i = 0; i<=outputData.acts.length-1; i++ ){
    filename = "АОСР_" + outputData.name + "_№"+ outputData.acts[i]+"_Реестр";
    as.getRangeByName("acts.reg").setValue(outputData.acts[i]);//установили номер реестра
    hide_rows('Р№', 'A5:A54',5);
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    exportPDFtoGDrive(folder, as.getId(),filename,[ [ sr.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,4,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[sr.getSheetId().toString(),null]]
    );
    console.log("Создал АОСР Реестр")
  };
  
  SpreadsheetApp.getUi().alert("Завершено"); 
};

function multiple_avk(outputData){
  //Работа с папками
  let rootID = "1kGeSf7L1PsNVXfIOKujxKCHJWQg4BQ9c";//папка для складывания папок АВК
  let root = DriveApp.getFolderById(rootID);//обращение по id папки
  let nameFolderFirst = "АВК_"+outputData.name + "_№"+ outputData.acts[outputData.acts.length - 1]+"-"+outputData.acts[0];//создаем название папки 1 уровня
  let nameFolderFirstID = root.createFolder(nameFolderFirst).getId();//создание папки с нужным namefolder 1 уровень
  folderFist=DriveApp.getFolderById(nameFolderFirstID); // взяли ее по id папку 1 уровня
  
  
  
  //Установка шифра раздела АВК
  var as = SpreadsheetApp.getActiveSpreadsheet();//активная таблица
  var ss = as.getSheetByName("АВК");//активный лист АВК
  console.log(outputData.acts.length);

  //Цикл создания PDF АВК + сопроводительной документации (17 сек)
  for (i = 0; i<outputData.acts.length; i++ ){
    //Работа с папками
    var folderSecond = DriveApp.getFolderById(nameFolderFirstID);
    var nameFolderSecond = "АВК_" + "№"+ outputData.acts[i] +"_"+ outputData.name;
    var folderSecondID = folderSecond.createFolder(nameFolderSecond).getId();
    folderSecond = DriveApp.getFolderById(folderSecondID);//////////////////////////////////////////////////////////////в эту папку сопроводительные

    //Создание PDF
    filename = "АВК_№"+ outputData.acts[i]+"_"+outputData.name;
    as.getRangeByName("acts.contr").setValue(outputData.acts[i]+"/"+outputData.name);
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    exportPDFtoGDrive(folderSecond, as.getId(),filename,[ [ ss.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,4,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[ss.getSheetId().toString(),null]]
    );
    
    //Создание массива link для ссылок на сопроводительные документы и nameLink для их названий
    var doc = as.getRangeByName("link.doc").getValues().flat().filter(ev => ev !="");
    var url = as.getRangeByName("link.url").getRichTextValues().flat().map(rt => rt.getLinkUrl()).filter(url => url !== null);

    //Цикл сохрания сопроводительных документов
    for (x=0;x<url.length;x++){
      filename = "Приложение_№"+[x]+"_"+doc[x]+".pdf";
      upload(url[x],filename,folderSecond);
    };   
  };
  SpreadsheetApp.getUi().alert("Завершено");
};