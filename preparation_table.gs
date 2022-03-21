function myFunction6() {
  var ash2 = SpreadsheetApp.getActiveSpreadsheet();
  var shtn3 = ash2.getSheetByName("(名前変更不可)仕込み表")

  today2 = shtn3.getRange("M1").getValue();

  var sht_tachikawa = ash2.getSheetByName("(名前変更不可)オーダーシート立川");
  var sht_shibuya = ash2.getSheetByName("(名前変更不可)オーダーシート東急渋谷");
  var sht_dome = ash2.getSheetByName("(名前変更不可)オーダーシートドーム");
  var sht_popup = ash2.getSheetByName("(名前変更不可)オーダーシートPOP-UP");
  var sht_mobile = ash2.getSheetByName("(名前変更不可)オーダーシート移動販売");

  //日付のリスト取り込み
  todayformatted2 = Utilities.formatDate(today2, "JST", "yyyy/MM/dd");
  const data_day2 = ash2.getRange("admin!D2:E33").getValues();


  //納品書の日付を取得
  for (var i=0;i<data_day2.length;i++){
    if (Utilities.formatDate(data_day2[i][1], "JST", "yyyy/MM/dd") == todayformatted2){
      var today_column2 = (i+1)*7-2
      break
    }
  }
  var lastRow2 = sht_shibuya.getRange('B:B').getValues().filter(String).length;
  
  //Logger.log(today_column2)
  
  function getsht2data(datasht){
    var data_all2 = datasht.getRange(6, 2, lastRow2,229).getValues();

    //列を削除
    
    var deleteArray2 = []
    for (var k=4;k<230;k++){
      if (k == today_column2+18){
        break
      }
      deleteArray2.push(k)    
    }

    for (var l=today_column2+19;l<230;l++){
      deleteArray2.push(l)
    }
      
    for (var m=0; m<data_all2.length; m++){
      for (var n=0; n<deleteArray2.length;n++){
        data_all2[m].splice(deleteArray2[n]-n, 1);
      }
    }

    return data_all2
  }

  function getsht2datapre(datasht){
    var data_all2 = datasht.getRange(6, 2, lastRow2,229).getValues();

    //列を削除
    
    var deleteArray2 = []
    for (var k=4;k<230;k++){
      if (k == today_column2+11){
        break
      }
      deleteArray2.push(k)    
    }

    for (var l=today_column2+12;l<230;l++){
      deleteArray2.push(l)
    }
      
    for (var m=0; m<data_all2.length; m++){
      for (var n=0; n<deleteArray2.length;n++){
        data_all2[m].splice(deleteArray2[n]-n, 1);
      }
    }

    return data_all2
  }

  function getsht2datasub(datasht2) {
    var data_all5 = datasht2.getRange(6, 2, lastRow2,229).getValues();
    var deleteArray2 = []
    for (var k=0;k<230;k++){
      if (k == today_column2+18){
        break
      }
      deleteArray2.push(k)    
    }

    for (var l=today_column2+19;l<230;l++){
      deleteArray2.push(l)
    }
      
    for (var m=0; m<data_all5.length; m++){
      for (var n=0; n<deleteArray2.length;n++){
        data_all5[m].splice(deleteArray2[n]-n, 1);
      }
    }
    return data_all5
  }

  function getsht2datasubpre(datasht2) {
    var data_all5 = datasht2.getRange(6, 2, lastRow2,229).getValues();
    var deleteArray2 = []
    for (var k=0;k<230;k++){
      if (k == today_column2+11){
        break
      }
      deleteArray2.push(k)    
    }

    for (var l=today_column2+12;l<230;l++){
      deleteArray2.push(l)
    }
      
    for (var m=0; m<data_all5.length; m++){
      for (var n=0; n<deleteArray2.length;n++){
        data_all5[m].splice(deleteArray2[n]-n, 1);
      }
    }
    return data_all5
  }




  data_dome = getsht2data(sht_dome);
  data_tachikawa = getsht2datasub(sht_tachikawa);
  data_shibuya = getsht2datasub(sht_shibuya);
  data_popup = getsht2datasub(sht_popup);
  data_mobile = getsht2datasub(sht_mobile);
  data_domepre = getsht2datapre(sht_dome);
  data_tachikawapre = getsht2datasubpre(sht_tachikawa);
  data_shibuyapre = getsht2datasubpre(sht_shibuya);
  data_popuppre = getsht2datasubpre(sht_popup);
  data_mobilepre = getsht2datasubpre(sht_mobile);
  
  var data_all3 = []
  
  for (var h=0;h<lastRow2;h++){
    data_all3[h] = data_dome[h].concat(data_tachikawa[h]).concat(data_shibuya[h]).concat(data_popup[h]).concat(data_mobile[h])
  }

  var data_all3pre = []
  for (var h=0;h<lastRow2;h++){
    data_all3pre[h] = data_domepre[h].concat(data_tachikawapre[h]).concat(data_shibuyapre[h]).concat(data_popuppre[h]).concat(data_mobilepre[h])
  }

  var all_datan3 = [];
  var all_datan3pre = [];

  for (var x=0;x<data_all3.length;x++) {
      all_datan3[x] = [data_all3[x][4],
                      data_all3[x][5],
                      data_all3[x][6],
                      data_all3[x][7],
                      data_all3[x][8]]
      all_datan3pre[x] = [data_all3pre[x][4],
                      data_all3pre[x][5],
                      data_all3pre[x][6],
                      data_all3pre[x][7],
                      data_all3pre[x][8]]
  }

  //Logger.log(all_datan3)
  if (!all_datan3.length){
    Browser.msgBox("記入されたデータがありません。オーダーシートを確認してください。", Browser.Buttons.OK)
  }
  var data_prepared = ash2.getRange("(名前変更不可)仕込み表!B3:C999").getValues();
  all_datan5 = []
  for (var x=0;x<data_prepared.length;x++) {
    num = data_prepared[x][0]-1
    Logger.log(num)
    if (num == -1){
      all_datan5[x] = ["","","","",""]
      continue
    }
    else if (num == 167 || num == 45){
      all_datan5[x] = [all_datan3[num][0],
                      all_datan3[num][1],
                      all_datan3[num][2],
                      all_datan3[num][3],
                      all_datan3[num][4]]
      continue

    }
    else if (num == 179){
      all_datan5[x] = [all_datan3pre[num][0],
                      all_datan3[num][1],
                      all_datan3pre[num][2],
                      all_datan3[num][3],
                      all_datan3[num][4]]
      continue
    }
    else{
      all_datan5[x] = [all_datan3pre[num][0],
                      all_datan3pre[num][1],
                      all_datan3pre[num][2],
                      all_datan3pre[num][3],
                      all_datan3pre[num][4]]
      continue 
    }
  }
  Logger.log(all_datan5)
  var lastColumn3 = all_datan5[0].length; //カラムの数を取得する
  var lastRow3 = all_datan5.length;   //行の数を取得する
  shtn3.getRange(3, 8, 1000, lastColumn3).clear();
  shtn3.getRange(3, 8, lastRow3, lastColumn3).setValues(all_datan5);
  //交互の背景色指定


var d = new Date();
var y = d.getFullYear();
var mon = d.getMonth()+1
var d2 = d.getDate();
var h = d.getHours();
var min = d.getMinutes();
var s = d.getSeconds();
var now = y+"/"+mon+"/"+d2+" "+h+":"+min+":"+s;
shtn3.getRange("M3").setValue(now);



//Logger.log(memo)

Browser.msgBox("更新が完了しました。OKを押した後、数秒で反映されます。", Browser.Buttons.OK)


  

}
