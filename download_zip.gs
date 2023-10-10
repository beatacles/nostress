///Выдает прямую ссылку на скачивание
function g2pc() {
  var folderId = "19myKgNT4eWO3fvFADuJ62tFUFcBBhIfe"; // Please set the folder ID here.

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var blobs = [];
  while (files.hasNext()) {
    blobs.push(files.next().getBlob());
  }
  var zipBlob = Utilities.zip(blobs, folder.getName() + ".zip");
  var fileId = DriveApp.createFile(zipBlob).getId();
  var url = "https://drive.google.com/uc?export=download&id=" + fileId;
  openUrl(url);
  Logger.log(url);
  openUrl(url);
}

/**
 * Open a URL in a new tab.
 * https://gist.github.com/smhmic/e7f9a8188f59bb1d9f992395c866a047
 */
function openUrl(url){
  var html = HtmlService.createHtmlOutput('<!DOCTYPE html><html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically.  Click below:<br/><a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(55);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
}