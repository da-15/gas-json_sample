var NAME_SHEET_LOG2 = 'ログ（取得）';
// スクリプトを公開 → ウェブアプリケーションとして導入...する。
// ※ アクセスできるユーザは匿名含む全員とする

function doGet(e) {

}

function doPost(e) {
  var params = JSON.parse(e.postData.getDataAsString());
  var value = params.value; // => JsonのValueに対する値が取れる
  
  setLog２(value)
  
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  output.setContent(JSON.stringify({ message: "success" }));
  
  return output;
}


/* --------------------------------------------------
ログシートに実行結果を出力する
*/
function setLog２(jsonData){
  var thisBook = SpreadsheetApp.getActiveSpreadsheet();
  var shLog = thisBook.getSheetByName(NAME_SHEET_LOG2);
  var newRow = shLog.getLastRow() + 1;
  var date = new Date();
  
  // 日時
  shLog.getRange(newRow, 1).setValue(Utilities.formatDate(date,'JST','yyyy/MM/dd HH:mm:ss'));
  // Jsonの内容
  shLog.getRange(newRow, 2).setValue(jsonData);
}


