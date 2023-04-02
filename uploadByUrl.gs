//Оркестрация загрузки файлов
function upload(url,filename,folderSecond){
  flag = isDrive(url);
  if (flag === true){
    uploadFilesFromDrive(url,filename,folderSecond);
  } else {
    uploadFilesFromOther(url,filename,folderSecond);
  }
};

//Проверяем является ли домен drive.goggle.com
function isDrive(url) {
  let domain = url.match(/^(?:https?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/\n?]+)/img);
  let result = domain[0].replace("https://", "");
  if (result == "drive.google.com") {
    return true;
    } else { false;}
};

//Загрузка с google drive
function uploadFilesFromDrive(url,filename,folderSecond){
  fileID = getIdFrom(url);
  let getFile = DriveApp.getFileById(fileID);
  blob = getFile.getBlob().setName(filename).getAs('application/pdf');
  file = folderSecond.createFile(blob);
  file.setDescription("Оригинал документа: " + url)
};

//Получение id файла с google drive
function getIdFrom(url) {
  var id = "";
  var parts = url.split(/^(([^:\/?#]+):)?(\/\/([^\/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?/);
  if (url.indexOf('?id=') >= 0){
     id = (parts[6].split("=")[1]).replace("&usp","");
     return id;
   } else {
   id = parts[5].split("/");
   //Using sort to get the id as it is the longest element. 
   var sortArr = id.sort(function(a,b){return b.length - a.length});
   id = sortArr[0];
   return id;
   }
 };

//Загрузка со сторонних сайтов
function uploadFilesFromOther(url,filename,folderSecond) {
  var response =  UrlFetchApp.fetch(url);
  var blob = response.getBlob();
  var file = folderSecond.createFile(blob);
  file.setName(filename);
  file.setDescription("Оригинал документа: " + url);
};




