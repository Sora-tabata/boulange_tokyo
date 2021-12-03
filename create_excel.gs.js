function myFunction4() {
    // 選択中のスプレッドシート
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var spreadSheetId = ss.getId();
    var sheetId = ss.getSheetId();
    var url = 'https://docs.google.com/spreadsheets/d/' + spreadSheetId + '/export?format=xlsx&gid=' + sheetId;
  
    // ダイアログ表示
    var html = HtmlService.createHtmlOutput('<p>リンクをクリックしてダウンロードしてください。</p>')
       .append('<p style="text-align:center;"><a href="' + url +'" style="text-align: center;">Dwonload<a/></p>')
       .append('<div style="text-align:right;"><input type="button" value="Close" onclick="google.script.host.close()" /></div>')
       .setWidth(300).setHeight(200);
     SpreadsheetApp.getUi().showModalDialog(html, 'ダウンロード');
  }
  