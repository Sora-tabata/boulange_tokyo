/** @OnlyCurrentDoc */
function myFunction(today, shop, flight){
  var ash = SpreadsheetApp.getActiveSpreadsheet();
  var shtn = ash.getSheetByName("(名前変更不可)納品書");
//引数の定義
  today = shtn.getRange("K1").getValue();
  shop = shtn.getRange("A1").getValue();

  todayformatted = Utilities.formatDate(today, "JST", "yyyy/MM/dd");
  const data_day = ash.getRange("admin!D2:E33").getValues();//日付のリスト取り込み

  for (var i=0;i<data_day.length;i++){
    if (Utilities.formatDate(data_day[i][1], "JST", "yyyy/MM/dd") == todayformatted){
      var today_column = (i+1)*7-1
      break
    }
  }

  if (shop == "BA立川"){
    var sht = ash.getSheetByName("(名前変更不可)オーダーシート立川");
    var lastRow = sht.getRange('B:B').getValues().filter(String).length;
    var data_all = sht.getRange(6, 1, lastRow,229).getValues();
    var memo = sht.getRange(1, today_column+1, 1, 1).getValue();
  }
  else if (shop == "BA東急渋谷"){
    var sht = ash.getSheetByName("(名前変更不可)オーダーシート東急渋谷");
    var lastRow = sht.getRange('B:B').getValues().filter(String).length;
    var data_all = sht.getRange(6,1,lastRow,229).getValues();
    var memo = sht.getRange(1, today_column+1, 1, 1).getValue();
  }
  else if (shop == "BAドーム"){
    var sht = ash.getSheetByName("(名前変更不可)オーダーシートドーム");
    var lastRow = sht.getRange('B:B').getValues().filter(String).length;
    var data_all = sht.getRange(6, 1, lastRow,229).getValues();
    var memo = sht.getRange(1, today_column+1, 1, 1).getValue();
  }
  else if (shop == "POP-UP"){
    var sht = ash.getSheetByName("(名前変更不可)オーダーシートPOP-UP");
    var lastRow = sht.getRange('B:B').getValues().filter(String).length;
    var data_all = sht.getRange(6, 1, lastRow,229).getValues();
    var memo = sht.getRange(1, today_column+1, 1, 1).getValue();
  }
  flight = shtn.getRange("B2").getValue();
  



//納品書の日付を取得



//data_allから合計が0になってる行を削除

  var all_datan = []
  
  for (var j=0;j<lastRow;j++){
    if (data_all[j][today_column+5] != 0){
      all_datan.push(data_all[j])
    }
  }
//列を削除
 
  var deleteArray = []
  for (var k=5;k<230;k++){
    if (k == today_column-1){
      break
    }
    deleteArray.push(k)    
  }
  for (var l=today_column+7;l<230;l++){
    deleteArray.push(l)
  }
  
  for (var m=0; m<all_datan.length; m++){
    for (var n=0; n<deleteArray.length;n++){
      all_datan[m].splice(deleteArray[n]-n, 1);
    }
  }

  //Logger.log(all_datan)
  //shtn.appendRow(all_datan)
  if (!all_datan.length){
    Browser.msgBox("記入されたデータがありません。オーダーシートを確認してください。", Browser.Buttons.OK)
  }
  var lastColumn = all_datan[0].length; //カラムの数を取得する
  Logger.log(lastColumn)
  var lastRow = all_datan.length;   //行の数を取得する

//全日ver
  if (flight == "全日"){
    shtn.getRange(5, 1, 1000, lastColumn).clear();
    shtn.getRange(5,1,lastRow,lastColumn).setValues(all_datan)
    shtn.getRange(1, 3, 1, 1).setValue(memo);
  }

//1便ver
  
  else if (flight == "1便"){
    for (var i=0; i<all_datan.length; i++){
      
      all_datan[i][7] = ''
      all_datan[i][8] = ''
      all_datan[i][10] = all_datan[i][6]
      all_datan[i][11] = all_datan[i][3]*all_datan[i][10]
      all_datan[i][5] = all_datan[i][2]*all_datan[i][10]
    }
    shtn.getRange(5, 1, 1000, lastColumn).clear();
    shtn.getRange(5,1,lastRow,lastColumn).setValues(all_datan);
    shtn.getRange(1, 3, 1, 1).setValue(memo);
  }

//2便ver
  else if (flight == "2便"){
    for (var i=0; i<all_datan.length; i++){
      
      all_datan[i][6] = ''
      all_datan[i][8] = ''
      all_datan[i][10] = all_datan[i][7]
      all_datan[i][11] = all_datan[i][3]*all_datan[i][10]
      all_datan[i][5] = all_datan[i][2]*all_datan[i][10]
    }
    shtn.getRange(5, 1, 1000, lastColumn).clear();
    shtn.getRange(5,1,lastRow,lastColumn).setValues(all_datan);
    shtn.getRange(1, 3, 1, 1).setValue(memo);
  }

//3便ver
  else if (flight == "3便"){
    for (var i=0; i<all_datan.length; i++){
      
      all_datan[i][6] = ''
      all_datan[i][7] = ''
      all_datan[i][10] = all_datan[i][8]
      all_datan[i][11] = all_datan[i][3]*all_datan[i][10]
      all_datan[i][5] = all_datan[i][2]*all_datan[i][10]
    }
    shtn.getRange(5, 1, 1000, lastColumn).clear();
    shtn.getRange(5,1,lastRow,lastColumn).setValues(all_datan);
    shtn.getRange(1, 3, 1, 1).setValue(memo);

  }
  shtn.getRange(5,13,lastRow,1).clear();
  //Logger.log(today_column)
  //Logger.log(all_datan)
  //Logger.log(lastRow)

  //交互の背景色指定
    for (var i=1;i<=all_datan.length;i++){
      if(i%2 == 0){
        shtn.getRange(i+4, 1, 1, lastColumn).setBackgroundColor('#D3D3D3');
      }
      else{
        shtn.getRange(i+4, 1, 1, lastColumn).setBackgroundColor('#FFFFFF');
      }
    }

  var d = new Date();
  var y = d.getFullYear();
  var mon = d.getMonth()+1
  var d2 = d.getDate();
  var h = d.getHours();
  var min = d.getMinutes();
  var s = d.getSeconds();
  var now = y+"/"+mon+"/"+d2+" "+h+":"+min+":"+s;
  shtn.getRange("K3").setValue(now);

    Browser.msgBox("更新が完了しました。OKを押した後、数秒で反映されます。", Browser.Buttons.OK)
  
}
