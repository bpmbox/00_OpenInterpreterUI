// Spreadsheetが開かれた時に自動的に実行されます.
function onOpen(e) {

  // 現在開いている、スプレッドシートを取得します.
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // メニュー項目を定義します.
  var entries = [
    {name : "宝飾 成約"  , functionName : "housyoku_createEVA"},
    {name : "服飾 成約"  , functionName : "brand_createEVA"},
    {name : "宝飾 全合わず"  , functionName : "awazu_housyoku_createEVA"},
    {name : "服飾 全合わず"  , functionName : "awazu_brand_createEVA"},
    //{name : "査物反映　※開発中"  , functionName : "sabutsu_createEVA"},
    {name : "宝飾 成約【ラベルあり】"  , functionName : "housyoku_createEVA_addLabel"},
    {name : "服飾 成約【ラベルあり】"  , functionName : "brand_createEVA_addLabel"},
  ];

  // 「追加メニュー」という名前でメニューに追加します.
  spreadsheet.addMenu("追加メニュー", entries);

    
    
  // シート生成のメニュー項目
  var entries = [
    {name : "宝飾 20行"  , functionName : "copy_kin_20"},
    {name : "宝飾 40行"  , functionName : "copy_kin_40"},
    {name : "宝飾 100行"  , functionName : "copy_kin_100"},
    {name : "服飾 20行"  , functionName : "copy_bra_20"},
    {name : "服飾 40行"  , functionName : "copy_bra_40"},
  ];
  spreadsheet.addMenu("シート生成", entries);


  // シート生成のメニュー項目
  var entries = [
    //{name : "金価格更新 old"  , functionName : "gold_price"},
    {name : "金価格更新"  , functionName : "gold_price_new"},
    {name : "※手動で本日のレートを更新する。（4時に自動反映してます。）"  , functionName : "rate_save"},
  ];
  spreadsheet.addMenu("リスト更新", entries);

  // 予約一覧のメニュー項目
  var entries = [
    {name : "予約一覧"  , functionName : "getReserves"},
  ];
  spreadsheet.addMenu("予約チェック", entries);

 /**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */

  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('音声録音', 'showSidebar')
      .addToUi();

} // END onOpen


function copy_kin_20(){
  copy_name_ss('宝20行',20);
}
function copy_kin_40(){
  copy_name_ss('宝40行',40);
}
function copy_kin_100(){
  copy_name_ss('宝100行',100);
}
function copy_bra_20(){
  copy_name_ss('服20行',20);
}
function copy_bra_40(){
  copy_name_ss('服40行',40);
}
/****************************************************************
スプレッドシートのシートのコピーと名前変更と保護
****************************************************************/
function copy_name_ss(name,num) {
  var ss_active_all = SpreadsheetApp.getActiveSpreadsheet();

  var ss_sheet_temp = ss_active_all.getSheetByName(name);
  var ss_sheet_copy = ss_sheet_temp.copyTo(ss_active_all);

  // コピーしたシートの名前変更
  //まず存在確認
  var ss_confirm = '';
  for(var i=1; i<=100; i++){
    ss_confirm = ss_active_all.getSheetByName(i);
    Logger.log(ss_confirm);
    if(ss_confirm == null){
      break;
    }
  }
  ss_sheet_copy.setName(i);
  SpreadsheetApp.setActiveSheet(ss_sheet_copy);
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(1);

  var last = parseInt(num) + 3;
  var range = ss_sheet_copy.getRange('A3:A'+last);
  var protection = range.protect().setDescription('A列の保護');

  protection.setWarningOnly(true);

}



   
    
//宝飾用成約処理
function housyoku_createEVA() {
    createEVA('seiyaku_housyoku');
}
    
//服飾用成約処理
function brand_createEVA() {
    createEVA('seiyaku_brand');
}
//宝飾用成約処理　ラベルあり
function housyoku_createEVA_addLabel(){
    createEVA('seiyaku_housyoku_addLabel');
}    
//服飾用成約処理　ラベルあり
function brand_createEVA_addLabel() {
    createEVA('seiyaku_brand_addLabel');
}

//全合わずのときの宝飾処理
function awazu_housyoku_createEVA() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetdata = spreadsheet.getSheetValues(1, 1, 1, spreadsheet.getLastColumn());
    var ecc_id = parseInt(sheetdata[0][15]); //顧客SEQの取得 P1

    //合わず顧客SEQ109175か、空欄出ないときはアラート出す
    if((ecc_id != 109175) && (ecc_id > 0)){
      var confirm = Browser.msgBox("顧客SEQが入ってます。全返却でよろしいですか？", Browser.Buttons.OK_CANCEL);
      if(confirm == 'ok'){
        //Browser.msgBox(confirm);
        createEVA('awazu_housyoku');
      }
    }else{
      createEVA('awazu_housyoku');
    }
}

//全合わずのときのブランド処理
function awazu_brand_createEVA() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetdata = spreadsheet.getSheetValues(1, 1, 1, spreadsheet.getLastColumn());
    var ecc_id = parseInt(sheetdata[0][14]); //顧客SEQの取得 O1

    //合わず顧客SEQ109175か、空欄出ないときはアラート出す
    if((ecc_id != 109175) && (ecc_id > 0)){
      var confirm = Browser.msgBox("顧客SEQが入ってます。全返却でよろしいですか？", Browser.Buttons.OK_CANCEL);
      if(confirm == 'ok'){
        //Browser.msgBox(confirm);
        createEVA('awazu_brand');
      }
    }else{
      createEVA('awazu_brand');
    }
}



    
function createEVA(type) {

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sht_title = spreadsheet.getName();
    var sht_url = spreadsheet.getUrl();
    var url_id = spreadsheet.getId();
    var sht = spreadsheet.getActiveSheet();
    var sht_id = sht.getSheetId();
    var sht_name = sht.getSheetName();

    // POSTデータ
    var data = {
    "url_id" : url_id,
    "sht_id" : sht_id,
    "sht_name" : sht_name,
    "sht_title" : sht_title,
    "sht_url" : sht_url,
    "type" : type,
    "tokentoken" : "djfkal;jfjkdaslfj;sdljvslf;dkjvfsdlk;jfo;sirfjer;wodfja;lkfjer;eoiwjfa;dosjv;odlfjair;oerwjfn;lksdnvlkscnv;lzcxknvo;ifsnh;igfsnjfg;iasdjhfoi;weahgo;rihjgo;ihejrg;osfadj;lasdjfgaoi"
    }
    // POSTオプション
    var options = {
    "method" : "POST",
    "payload" : data,
    "muteHttpExceptions" : true,
    }

    var url = "http://rifa.life/lounge_API/buy_card_lounge.php";
    var response = UrlFetchApp.fetch(url, options);
    var content = response.getContentText("UTF-8");

    var html = HtmlService
    .createHtmlOutput(content)
    .setSandboxMode(HtmlService.SandboxMode.EMULATED)
    .setWidth(1200)
    .setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, '結果通知');
}


/*

金価格更新

*/
function gold_price(type) {

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var sheet_1 = spreadsheet.getSheetByName('価格表');
  var sheet_2 = spreadsheet.getSheetByName('価格マスタ');
  var formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy/MM/dd");
  var formattedDate_2 = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  
  var objDate = new Date();
  objDate.setDate(objDate.getDate() - 1);
  var yesterday = Utilities.formatDate(objDate, "GMT", "yyyy-MM-dd");
  
  var sheet_1_a1_val = sheet_1.getRange("A1").getValue();
  var sheet_2_b1_val = sheet_2.getRange("B1").getValue();
  
  if(sheet_1_a1_val == ''){
    sheet_1.getRange("A1").setValue(formattedDate);
  }
  if(sheet_2_b1_val == ''){
    sheet_2.getRange("B1").setValue(formattedDate_2);
    sheet_2.getRange("B2").setValue(yesterday);
  }
  
  
  var sht_title = spreadsheet.getName();
  var sht_url = spreadsheet.getUrl();
  var url_id = spreadsheet.getId();
  var sht = spreadsheet.getActiveSheet();
  var sht_id = sht.getSheetId();
  var sht_name = sht.getSheetName();
  
  // POSTデータ
  var data = {
    "url_id" : url_id,
    "sht_id" : sht_id,
    "sht_name" : sht_name,
    "sht_title" : sht_title,
    "sht_url" : sht_url,
    "type" : type,
    "tokentoken" : "djfkal;jfjkdaslfj;sdljvslf;dkjvfsdlk;jfo;sirfjer;wodfja;lkfjer;eoiwjfa;dosjv;odlfjair;oerwjfn;lksdnvlkscnv;lzcxknvo;ifsnh;igfsnjfg;iasdjhfoi;weahgo;rihjgo;ihejrg;osfadj;lasdjfgaoi"
  }
  // POSTオプション
  var options = {
    "method" : "POST",
    "payload" : data,
    "muteHttpExceptions" : true,
  }
  
  var url = "http://rifa.life/lounge_API/gold_price.php";
  var response = UrlFetchApp.fetch(url, options);
  var content = response.getContentText("UTF-8");
}



/*

レート更新

*/
function rate_save(type) {

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sht_title = spreadsheet.getName();
    var sht_url = spreadsheet.getUrl();
    var url_id = spreadsheet.getId();
    var sht = spreadsheet.getActiveSheet();
    var sht_id = sht.getSheetId();
    var sht_name = sht.getSheetName();

    // POSTデータ
    var data = {
    "url_id" : url_id,
    "sht_id" : sht_id,
    "sht_name" : sht_name,
    "sht_title" : sht_title,
    "sht_url" : sht_url,
    "type" : type,
    "tokentoken" : "djfkal;jfjkdaslfj;sdljvslf;dkjvfsdlk;jfo;sirfjer;wodfja;lkfjer;eoiwjfa;dosjv;odlfjair;oerwjfn;lksdnvlkscnv;lzcxknvo;ifsnh;igfsnjfg;iasdjhfoi;weahgo;rihjgo;ihejrg;osfadj;lasdjfgaoi"
    }
    // POSTオプション
    var options = {
    "method" : "POST",
    "payload" : data,
    "muteHttpExceptions" : true,
    }

    var url = "http://rifa.life/lounge_API/rate_save.php";
    var response = UrlFetchApp.fetch(url, options);
    var content = response.getContentText("UTF-8");

    var html = HtmlService
    .createHtmlOutput("「金価格更新」ボタンを押したらGSSへ料率反映されます。")
    .setSandboxMode(HtmlService.SandboxMode.EMULATED)
    .setWidth(1200)
    .setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, '結果通知');

}

function gold_price_new(type) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var sheet_1 = spreadsheet.getSheetByName('価格表');
  var sheet_2 = spreadsheet.getSheetByName('価格マスタ');

  
  var sht_title = spreadsheet.getName();
  var sht_url = spreadsheet.getUrl();
  var url_id = spreadsheet.getId();
  var sht = spreadsheet.getActiveSheet();
  var sht_id = sht.getSheetId();
  var sht_name = sht.getSheetName();
  
  // POSTデータ
  var data = {
    "url_id" : url_id,
    "sht_id" : sht_id,
    "sht_name" : sht_name,
    "sht_title" : sht_title,
    "sht_url" : sht_url,
    "type" : type,
    "tokentoken" : "djfkal;jfjkdaslfj;sdljvslf;dkjvfsdlk;jfo;sirfjer;wodfja;lkfjer;eoiwjfa;dosjv;odlfjair;oerwjfn;lksdnvlkscnv;lzcxknvo;ifsnh;igfsnjfg;iasdjhfoi;weahgo;rihjgo;ihejrg;osfadj;lasdjfgaoi"
  }
  // POSTオプション
  var options = {
    "method" : "POST",
    "payload" : data,
    "muteHttpExceptions" : true,
  }
  
  var url = "http://rifa.life/lounge_API/gold_price_new.php";
  var response = UrlFetchApp.fetch(url, options);
  var content = response.getContentText("UTF-8");
  
  var html = HtmlService
    .createHtmlOutput(content)
    .setSandboxMode(HtmlService.SandboxMode.EMULATED)
    .setWidth(1200)
    .setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, '結果通知');

}


function getReserves() {

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var spread_sheet_id = spreadsheet.getId();
    var sht = spreadsheet.getActiveSheet();
    var sheet_name = sht.getSheetName();


    var url = "https://rifa.life/evaProject/ycbm/shop_purchase_reserves/" + spread_sheet_id + "/" + sheet_name;
    var response = UrlFetchApp.fetch(url);
    var content = response.getContentText("UTF-8");

    var html = HtmlService
    .createHtmlOutput(content)
    .setSandboxMode(HtmlService.SandboxMode.EMULATED)
    .setWidth(1200)
    .setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, '結果通知');
}
