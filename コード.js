// シートの自動作成スクリプト
function create_tentou_Sheet() {
  // テンプレートファイル
  var templateFile = DriveApp.getFileById('1HFIpn6wxwVfdpqMD6zJh2zusZO6xSnRvV1m_4NQZNvQ');
  // 出力フォルダ

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("シートマスタ");
  var yyyy = sheet.getRange("A2").setValue(Utilities.formatDate(new Date(),"JST","yyyy"));
  var mm = sheet.getRange("A3").setValue(Utilities.formatDate(new Date(),"JST","MM"));
  var dd = sheet.getRange("A4").setValue(Utilities.formatDate(new Date(),"JST","dd"));

  var now_line = sheet.getRange("A11").getValue(); // 今年の行
  var column = [
    "D","E","F","G","H","I","J","K","L","M","N","O"
  ]; 
  var now_column = column[(Utilities.formatDate(new Date(),"JST","MM") - 1)];
  
  var folder_id_cell = now_column+now_line;
  var folder_id = sheet.getRange(folder_id_cell).getValue();
  
  Logger.log(folder_id);
  
  
  var OutputFolder = DriveApp.getFolderById(folder_id);
  // 出力ファイル名
  var OutputFileName = Utilities.formatDate(new Date(),"JST","yyMMdd")+'-店頭買取';
//  var OutputFileName = 'testests-店頭買取';
  
  var copy_file = templateFile.makeCopy(OutputFileName, OutputFolder);





  //　コピーしたファイルのURL
    var sht_url = copy_file.getUrl();

    // POSTデータ
    var data = {
    "sht_url" : sht_url,
    "tokentoken" : "djfkal;jfjkdaslfj;sdljvslf;dkjvfsdlk;jfo;sirfjer;wodfja;lkfjer;eoiwjfa;dosjv;odlfjair;oerwjfn;lksdnvlkscnv;lzcxknvo;ifsnh;igfsnjfg;iasdjhfoi;weahgo;rihjgo;ihejrg;osfadj;lasdjfgaoi"
    }
    // POSTオプション
    var options = {
    "method" : "POST",
    "payload" : data,
    "muteHttpExceptions" : true,
    }

    var url = "http://rifa.life/lounge_API/create_tentou_sheet.php";
    var response = UrlFetchApp.fetch(url, options);
    var content = response.getContentText("UTF-8");

}


function datetest(){
  //comment 
  var a = Utilities.formatDate(new Date(),"JST","yyMMddHH");
  Logger.log(a);
}