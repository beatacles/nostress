function encodeDate( yy, mm, dd, hh, ii, ss ){
  var days=[31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if(((yy % 4) == 0) && (((yy % 100) != 0) || ((yy % 400) == 0)))days[1]=29;
  for(var i=0; i<mm; i++) dd += days[ i ];
  yy--;
  return ((((yy * 365 + ((yy-(yy % 4)) / 4) - ((yy-(yy % 100)) / 100) + ((yy-(yy % 400)) / 400) + dd - 693594) * 24 + hh) * 60 + ii) * 60 + ss)/86400.0;
}

function openUrl( url ){
  //Browser.msgBox( '\\\"' );
  var html = HtmlService.createHtmlOutput( '<html><script>' +
    'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};' +
    'var a = document.createElement("a"); a.href="' + url + '"; a.target="_blank";' +
    'if(document.createEvent){' +
    '  var event=document.createEvent("MouseEvents");' +
    '  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}' +
    '  event.initEvent("click",true,true); a.dispatchEvent(event);' +
    '}else{ a.click() }' +
    'close();' +
    '</script>' +
    '<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="' + url + '" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>' +
    '<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>' +
    '</html>').setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
}

function getESID(){
  return ( Math.round( Math.random() * 10000000 ) );
}
function exportPDF( ssID, source, options, format, breaks){
  var dt=new Date();
  var d=encodeDate( dt.getFullYear(), dt.getMonth(), dt.getDate(), dt.getHours(), dt.getMinutes(), dt.getSeconds() );
  var pc=[ null, null, null, null, null, null, null, null, null, 0,
          source,
          10000000, null, null, null, null, null, null, null, null, null, null, null, null, null, null,
          d,
          null, null,
          options,
          format,
          null, 0, breaks, 0 ];
  openUrl( "https://docs.google.com/spreadsheets/d/" + ssID + "/pdf?" +
    "esid=" + getESID() + "&" +
    "id=" + ssID + "&" + 
    "a=true&" +
    "pc=" + encodeURIComponent( JSON.stringify( pc ) ) + "&" +
    "gf=[]&" +
    "lds=[]" );
}

function myExportPDF(){
  var as = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = as.getActiveSheet();
  exportPDF(as.getId(), // Идентификатор таблицы
    [ 
      [ sheet.getSheetId().toString(), 
       0,                             // начальная граница по вертикали
       82,                           // конечная граница по вертикали
       0,                             // начальная граница по горизонтали
       20                              // конечная граница по горизонтали
      ] // 
    ],
    [ // options
      0, /* Не показывать заметки (Hide notes) */
      null,
      1, /* Показывать линии сетки ( Show grid ) */
      0, /* Не показывать номера страниц ( Hide page number ) */
      0, /* Не показывать название книги ( Hide spreadsheet title ) */
      0, /* Не показывать название листа ( Hide sheet title ) */
      0, /* Не показывать текущую дату ( Hide current date ) */
      0, /* Не показывать текущее время ( Hide current time ) */
      1, /* Повторять закрепленные строки ( Repeat fixed rows ) */
      1, /* Повторять закрепленные столбцы ( Repeat fixed columns ) */
      1, /* Порядок страниц вниз, затем вверх  ( Page order down then up ) */
      1,
      null,
      null,
      1, /* Выравнивание по левому краю ( Left align ) */
      1 /* Выравнивание по верхнему краю (Top align ) */
    ],
    [ // format
      "A4", /* Формат листа A4 ( A4 sheet format ) */
      1, /* Ориентация страницы вертикальная ( Page orientation vertical ) */ 
      6, /* Выровнять по разрывам страниц ( Align to page breaks ) */
      1,
      [
        0.275591,    // Отступ сверху
        0.275591,    // Отступ снизу
        0.551181,     // Отступ слева
        0.551181      // Отступ справа
      ]
    ],
    [ // breaks
      [
        sheet.getSheetId().toString(),
        [ [ 0, 36 ] ] // 36 строка разрыва
      ]

    ]
  );
};

function getFolder( path ){
  var folder = null;
  var parents = DriveApp.getFileById( SpreadsheetApp.getActiveSpreadsheet().getId() ).getParents();
  if( parents.hasNext() ) folder = parents.next();
  else folder = DriveApp.getRootFolder();

  if( path.charAt(0) == '/' ) path = path.substr( 1, path.length - 1 );
  if( path.charAt( path.length - 1 ) == '/' ) path = path.substr( 0, path.length - 1 );
  var folders = path.split( '/' );

  if(!folders)return folder;

  var fldrs = null, id;
  for( var i = 0; i < folders.length; i++ ){
    id = folder.getId();
    fldrs = folder.getFoldersByName( folders[ i ] );
    while( fldrs.hasNext() ){
      folder = fldrs.next();
      if( folder.getParents().next().getId() == id )break;
    }
    if( id == folder.getId() ) return null;
  }
  return folder;
};

function exportPDFtoGDrive( folder, ssID, filename, source, options, format, breaks ){
  var dt = new Date();
  var d = encodeDate( dt.getFullYear(), dt.getMonth(), dt.getDate(), dt.getHours(), dt.getMinutes(), dt.getSeconds() );
  var pc=[ null, null, null, null, null, null, null, null, null, 0,
          source,
          10000000, null, null, null, null, null, null, null, null, null, null, null, null, null, null,
          d,
          null, null,
          options,
          format,
          null, 0, breaks ,0 ];

  var options = {
    'method': 'post',
    'payload': "a=true&pc=" + JSON.stringify(pc) + "&gf=[]&lds=[]",
    'headers': {Authorization: "Bearer " + ScriptApp.getOAuthToken()},
    'muteHttpExceptions': true
  };
  var theBlob = UrlFetchApp.fetch( "https://docs.google.com/spreadsheets/d/" + ssID + "/pdf?" +
    "esid=" + getESID() + "&" +
    "id=" + ssID, options ).getBlob();

  return folder.createFile( theBlob ).setName( filename + ".pdf" );
};