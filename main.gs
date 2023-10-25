///Папки
const folderID = "19myKgNT4eWO3fvFADuJ62tFUFcBBhIfe";//id корневой папки
const folder = DriveApp.getFolderById(folderID);//обращение по id папки
//Переменные
const as = SpreadsheetApp.getActiveSpreadsheet();//активная таблица

function inch(value){
  return 0.393701 * value;
};


//Создание модального окна Создать ИД
function modal(){
  const modalForm = HtmlService.createTemplateFromFile("modal")
  const htmlOutput = modalForm.evaluate()
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(htmlOutput, "Меню создания ИД");
};

//Запуск по информации из формы
function processForm(inputData) {
  console.log("Входные данные:",inputData);
  make_aosr(inputData);
  make_vso(inputData);
  make_vik(inputData);
  make_apt(inputData);
  make_ag(inputData);
  make_r1(inputData);
  make_r2(inputData);
  make_r3(inputData);
  make_tp(inputData);
  make_ad(inputData);
  SpreadsheetApp.getUi().alert("Завершено");
}

//Функци создания АОСР и реестров материалов для АОСР
function make_aosr(inputData){
  if (inputData.aosr == false) {return}
  var ss = as.getSheetByName("АОСР");//активный лист АОСР
  var sr = as.getSheetByName("Р№");//активный лист Реестра
  const n_aosr = as.getRangeByName("n.aosr").getValue().split(",");
  
  //Цикл создания PDF АОСР
  for (i = 0; i<=n_aosr.length-1; i++ ){
    filename = n_aosr[i]+"_Акт_АОСР_1";
    as.getRangeByName("aosr").setValue(n_aosr[i]);
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    exportPDFtoGDrive(folder, as.getId(),filename,[ [ ss.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,6,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[ss.getSheetId().toString(),[[0,36],[36,82]],[[0,20]]]]
    );
    console.log("Создал АОСР "+n_aosr[i]);
  };

  //Цикл создания PDF Реестр к АОСР
  for (x = 0; x<=n_aosr.length-1; x++ ){
    name = n_aosr[x]+"_Акт_АОСР_2";
    as.getRangeByName("aosr.r").setValue(n_aosr[x]);
    hide_rows('Р№', 'B4:B88',4);
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    exportPDFtoGDrive(folder, as.getId(),name,[ [ sr.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,4,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[sr.getSheetId().toString(),null]]
    );
    console.log("Создал АОСР Реестр "+n_aosr[x]);
  };
  SpreadsheetApp.getUi().alert("Завершено"); 
};

//Создать Титульный лист
function make_tp(inputData){
  if (inputData.tp == false) {return}
  var ss = as.getSheetByName("Т");//активный лист
  filename = "0_Титульный_лист"
  exportPDFtoGDrive(folder, as.getId(),filename,[ [ ss.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,4,1,[inch(2),0,0,0]],[[ss.getSheetId().toString(),null]]
    );
};

//Создать Реестр№1
function make_r1(inputData){
  if (inputData.r1 == false) {return}
  var ss = as.getSheetByName("Р1");//активный лист
  filename = "0_Реестр_№1"
  exportPDFtoGDrive(folder, as.getId(),filename,[ [ ss.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,4,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[ss.getSheetId().toString(),null]]
    );
};

//Создать Реестр№2
function make_r2(inputData){
  if (inputData.r1 == false) {return}
  var ss = as.getSheetByName("Р2");//активный лист
  filename = "0_Реестр_№2"
  hide_rows('Р2', 'B9:B108',9);
  exportPDFtoGDrive(folder, as.getId(),filename,[ [ ss.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,4,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[ss.getSheetId().toString(),null]]
    );
};

//Создать Реестр№3
function make_r3(inputData){
  if (inputData.r3 == false) {return}
  var ss = as.getSheetByName("Р3");//активный лист
  filename = "0_Реестр_№3";
  hide_rows('Р3', 'B4:B103',4);
  exportPDFtoGDrive(folder, as.getId(),filename,[ [ ss.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,4,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[ss.getSheetId().toString(),null]]
    );
};

//Создать ВИК
function make_vik(inputData){
  if (inputData.vik == false) {return}
  var ss = as.getSheetByName("ВИК");//активный лист
  filename = as.getRangeByName("n.vik").getValue()+"_Акт_ВИК"
  exportPDFtoGDrive(folder, as.getId(),filename,[ [ ss.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,4,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[ss.getSheetId().toString(),null]]
    );
};

//Создать Акт о проведении промывки трубопроводов
function make_apt(inputData){
  if (inputData.apt == false) {return}
  var ss = as.getSheetByName("АПТ");//активный лист
  filename = as.getRangeByName("n.apt").getValue()+"_Акт_АПТ"
  exportPDFtoGDrive(folder, as.getId(),filename,[ [ ss.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,4,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[ss.getSheetId().toString(),null]]
    );
};

//Создать Ведомость смонтированного оборудования
function make_vso(inputData){
  if (inputData.vso == false) {return}
  var ss = as.getSheetByName("ВСО");//активный лист
  filename = "ВСО"
  hide_rows('ВСО', 'B11:B160',11);
  exportPDFtoGDrive(folder, as.getId(),filename,[ [ ss.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,4,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[ss.getSheetId().toString(),null]]
    );
};

//Создать Акт о проведении гидростатического или манометрического испытания на герметичность
function make_ag(inputData){
  if (inputData.ag == false) {return}
  var ss1 = as.getSheetByName("АГ1");//активный лист
  filename = as.getRangeByName("n.ag").getValue().split(",")
  filename1 = filename[0]+"_Акт_АГ1"
  exportPDFtoGDrive(folder, as.getId(),filename1,[ [ ss1.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,4,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[ss1.getSheetId().toString(),null]]
    );
  var ss2 = as.getSheetByName("АГ2");//активный лист
  filename2 = filename[1]+"_Акт_АГ2"
  exportPDFtoGDrive(folder, as.getId(),filename2,[ [ ss2.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,4,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[ss2.getSheetId().toString(),null]]
    );
};

//Создать Акт дезинфекции
function make_ad(inputData){
  if (inputData.ad == false) {return}
  var ss = as.getSheetByName("АД");//активный лист АОСР
  var n_aosr = as.getRangeByName("n.ad");
  filename = n_aosr[0]+"_Акт_АОСР_Акт дезинфекции";
    as.getRangeByName("aosr").setValue(n_aosr[i]);
    exportPDFtoGDrive(folder, as.getId(),filename,[ [ ss.getSheetId().toString() ] ],
      [ 0, null, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, null, null, 2, 1 ],
      ["A4",1,6,1,[0.19685039370078738,0.19685039370078738,0.5905511811023622,0.5905511811023622]],[[ss.getSheetId().toString(),[[0,36],[36,82]],[[0,20]]]]
    );
    console.log("Создал АОСР Акт дезинфекции");
  };

//Создать сопроводительные документы
function make_doc_folder(){
  const objectName = as.getName() + ". СД";
  const target = folder.createFolder(objectName).getId();
  const targetFolder = DriveApp.getFolderById(target);
  //Создание массива link для ссылок и имен
    var doc = as.getRangeByName("link").getValues().flat().filter(ev => ev !="");
    var url = as.getRangeByName("link").getRichTextValues().flat().map(rt => rt.getLinkUrl()).filter(url => url !== null);
    var count_doc = doc.length;
    var count_url = url.length;

    //Цикл сохранения сопроводительных документов
    for (i=0;i<url.length;i++){
      filename = "Документ_№"+[i]+"_"+doc[i]+".pdf"
      upload(url[i],filename,targetFolder);
    };
    SpreadsheetApp.getUi().alert("Загружено "+ count_url + " документов из "+ count_doc);
};