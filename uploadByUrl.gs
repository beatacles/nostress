//Оркестрация загрузки файлов
function upload(url,filename,folder){
  const flag = isDrive(url);
  if (flag === true){
    uploadFilesFromDrive(url,filename,folder);
  } else {
    uploadFilesFromOther(url,filename,folder);
  }
};

//Проверяем является ли домен drive.goggle.com
function isDrive(url) {
  const domain = url.match(/^(?:https?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/\n?]+)/img);
  const result = domain[0].replace("https://", "");
  return result === "drive.google.com";
};

//Загрузка с google drive
function uploadFilesFromDrive(url,filename,folder){
  const fileID = getIdFrom(url);
  const getFile = DriveApp.getFileById(fileID);
  const blob = getFile.getBlob().setName(filename).getAs('application/pdf');
  const file = folder.createFile(blob);
  file.setDescription(`Оригинал документа: ${url}`);
};

//Получение id файла с google drive
function getIdFrom(url) {
  var id = "";
  const parts = url.split(/^(([^:\/?#]+):)?(\/\/([^\/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?/);
  if (url.indexOf('?id=') >= 0){
     id = (parts[6].split("=")[1]).replace("&usp","");
     return id;
   } else {
   id = parts[5].split("/");
   //Using sort to get the id as it is the longest element. 
   const sortArr = id.sort(function(a,b){return b.length - a.length});
   id = sortArr[0];
   return id;
   }
 };

//Загрузка со сторонних сайтов
function uploadFilesFromOther(url,filename,folder) {
  const response =  UrlFetchApp.fetch(url);
  const blob = response.getBlob();
  const file = folder.createFile(blob);
  file.setName(filename);
  file.setDescription(`Оригинал документа: ${url}`);
};