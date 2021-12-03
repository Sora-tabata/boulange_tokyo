function myFunction2() {
  var ash2 = SpreadsheetApp.getActiveSpreadsheet();
  var shtn3 = ash2.getSheetByName("(名前変更不可)オーダーシート3店舗分")

  today2 = shtn3.getRange("A1").getValue();

  var sht_tachikawa = ash2.getSheetByName("(名前変更不可)オーダーシート立川");
  var sht_shibuya = ash2.getSheetByName("(名前変更不可)オーダーシート東急渋谷");
  var sht_dome = ash2.getSheetByName("(名前変更不可)オーダーシートドーム");

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
  

  
  function getsht2data(datasht){
    var data_all2 = datasht.getRange(6, 2, lastRow2,229).getValues();

    //列を削除
    
    var deleteArray2 = []
    for (var k=4;k<230;k++){
      if (k == today_column2+4){
        break
      }
      deleteArray2.push(k)    
    }

    for (var l=today_column2+5;l<230;l++){
      deleteArray2.push(l)
    }
      
    for (var m=0; m<data_all2.length; m++){
      for (var n=0; n<deleteArray2.length;n++){
        data_all2[m].splice(deleteArray2[n]-n, 1);
      }
    }

    return data_all2
  }

  data_dome = getsht2data(sht_dome);

  function getsht2datasub(datasht2) {
    var data_all5 = datasht2.getRange(6, 2, lastRow2,229).getValues();
    var deleteArray2 = []
    for (var k=0;k<230;k++){
      if (k == today_column2+4){
        break
      }
      deleteArray2.push(k)    
    }

    for (var l=today_column2+5;l<230;l++){
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


  var data_all3 = []
  
  for (var h=0;h<lastRow2;h++){
    data_all3[h] = data_dome[h].concat(data_tachikawa[h]).concat(data_shibuya[h])
  }
  Logger.log(data_all3)
  var all_datan3 = [];

  for (var x=0;x<data_all3.length;x++) {
      all_datan3[x] = [data_all3[x][0],
                      data_all3[x][1],
                      data_all3[x][2],
                      data_all3[x][3],
                      data_all3[x][1]*data_all3[x][4],
                      data_all3[x][4],
                      data_all3[x][2]*data_all3[x][4],
                      data_all3[x][1]*data_all3[x][5],
                      data_all3[x][5],
                      data_all3[x][2]*data_all3[x][5],
                      data_all3[x][1]*data_all3[x][6],
                      data_all3[x][6],
                      data_all3[x][2]*data_all3[x][6]]
  }
  for (var y=0;y<all_datan3.length;y++){
    if (all_datan3[y][5] == 0 && all_datan3[y][8] == 0 && all_datan3[y][11] == 0){
      all_datan3[y].splice(0, 13)
    }
  }

  var all_datan4 = []
  for (var z=0;z<all_datan3.length;z++){
    if (all_datan3[z].length != 0){
      all_datan4.push(all_datan3[z])
    }
  }
  var lastColumn3 = all_datan4[0].length; //カラムの数を取得する
  var lastRow3 = all_datan4.length;   //行の数を取得する
  shtn3.getRange(6, 1, 1000, lastColumn3).clear();
  shtn3.getRange(6, 1, lastRow3, lastColumn3).setValues(all_datan4);
  //交互の背景色指定
  for (var i=1;i<=lastRow3;i++){
    if(i%2 == 0){
      shtn3.getRange(i+5, 1, 1, lastColumn3).setBackgroundColor('#D3D3D3');
      }
    else{
      shtn3.getRange(i+5, 1, 1, lastColumn3).setBackgroundColor('#FFFFFF');
      }
    shtn3.getRange(i+5, 6, 1, 1).setBackgroundColor('#ffff00')
    shtn3.getRange(i+5, 9, 1, 1).setBackgroundColor('#ffff00')
    shtn3.getRange(i+5, 12, 1, 1).setBackgroundColor('#ffff00')
 }

var d = new Date();
var y = d.getFullYear();
var mon = d.getMonth()+1
var d2 = d.getDate();
var h = d.getHours();
var min = d.getMinutes();
var s = d.getSeconds();
var now = y+"/"+mon+"/"+d2+" "+h+":"+min+":"+s;
shtn3.getRange("C4").setValue(now);
 Browser.msgBox("更新が完了しました。OKを押した後、数秒で反映されます。", Browser.Buttons.OK)


  

}