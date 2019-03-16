var ROW_START_DATA = 3; // データ開始行
var COL_1 = 2; // データが定義してある列
var COL_2 = 3;

var NAME_SHEET_DATA = 'データ設定';
var NAME_SHEET_LOG = 'ログ（送信）';
var API_URL = 'B1'; // Json送付先のURLを指定


/* --------------------------------------------------
メニュ-の表示
*/
function onOpen(){
  //メニュー配列
  var myMenu=[
      {name: 'JSONをPOSTする', functionName: 'main'}
    ];
  //メニューを追加
  SpreadsheetApp.getActiveSpreadsheet().addMenu('★マクロ実行',myMenu);
}


/* --------------------------------------------------
main
*/
function main(){
  var arrSKU = [];
  var arrData = [];
  var thisBook = SpreadsheetApp.getActiveSpreadsheet();
  var shData = thisBook.getSheetByName(NAME_SHEET_DATA);
  var strURL = shData.getRange(API_URL).getValue();
  
  // SKUと在庫数のリストを取得
  arrSKU = getActiveRowValues(shData, COL_1, COL_2, ROW_START_DATA);
  Logger.log(arrSKU.length);
  
  
  for(var i = 0; i < arrSKU.length; i++){
    // シート名を付与
    Logger.log((i + 1) + '_' + arrSKU[i]);
    
    arrData = arrSKU[i].split(',');
    // 在庫反映
    sendJsonPost(arrData[0], arrData[1], strURL);
  }
  
  Browser.msgBox('反映が完了しました。\\n結果はシート「ログ」を参照してください。');
}

/* --------------------------------------------------
指定した列を配列に取込(カンマ区切り)
*/
function getActiveRowValues(sheet, numCol1, numCol2, numStartRow){
  var arrResult = [];
  var strData;
  
  if(numStartRow > 0 && numCol1 > 0){
    //指定された列を配列にする
    for(var i = numStartRow; i <= sheet.getLastRow(); i++){
      strData = '';
      strData += sheet.getRange(i, numCol1).getValue() + ',';
      strData += sheet.getRange(i, numCol2).getValue();
      
      arrResult.push(strData);
    }
  }

  return arrResult;
}

/* --------------------------------------------------
ログシートに実行結果を出力する
*/
function setLog(jsonData, response){
  var thisBook = SpreadsheetApp.getActiveSpreadsheet();
  var shLog = thisBook.getSheetByName(NAME_SHEET_LOG);
  var newRow = shLog.getLastRow() + 1;
  var date = new Date();
  
  // 日時
  shLog.getRange(newRow, 1).setValue(Utilities.formatDate(date,'JST','yyyy/MM/dd HH:mm:ss'));
  // Jsonの内容
  shLog.getRange(newRow, 2).setValue(jsonData);
  // 実行結果
  shLog.getRange(newRow, 3).setValue(response);

}

/* --------------------------------------------------
在庫をセット -> JSONデータの送信
*/
function sendJsonPost(goods_sku, stock, apiUrl){
  // Basic認証のとき有効化
  // var authData = Utilities.base64Encode(basicUsr + ':' + basicPass);
  
  // 実行日時
  var now = new Date();
  
  // JSONデータを生成
  var jsonData = {
    'value': goods_sku
    };
  
  // Postデータを生成
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(jsonData),
    // Basic認証のとき有効化
    // 'headers': {'Authorization' : 'Basic ' + authData},
    'muteHttpExceptions': true
    };
  
  Logger.log(JSON.stringify(jsonData).trim());
  
  // JSONをPostする。
  var response = UrlFetchApp.fetch(apiUrl, options);
  Logger.log(response.getContentText().trim());
  
  // ログを出力
  setLog(JSON.stringify(jsonData).trim(), response.getContentText().trim());
}