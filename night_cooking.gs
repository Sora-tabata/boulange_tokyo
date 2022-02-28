function myFunction5() {
  var ash2 = SpreadsheetApp.getActiveSpreadsheet();
  var shtn3 = ash2.getSheetByName("(名前変更不可)オーダーシート3店舗分")
  var shtn5 = ash2.getSheetByName("(名前変更不可)夜勤製造表")

  today2 = shtn5.getRange("B1").getValue();

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

  function getsht2datapre(datasht){
    var data_all2 = datasht.getRange(6, 2, lastRow2,229).getValues();

    //列を削除
    
    var deleteArray2 = []
    for (var k=4;k<230;k++){
      if (k == today_column2-3){
        break
      }
      deleteArray2.push(k)    
    }

    for (var l=today_column2-2;l<230;l++){
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

  function getsht2datasubpre(datasht2) {
    var data_all5 = datasht2.getRange(6, 2, lastRow2,229).getValues();
    var deleteArray2 = []
    for (var k=0;k<230;k++){
      if (k == today_column2-3){
        break
      }
      deleteArray2.push(k)    
    }

    for (var l=today_column2-2;l<230;l++){
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
  if(today_column2 == 5){
    data_domepre = data_dome
    data_tachikawapre = data_tachikawa
    data_shibuyapre = data_shibuya
  }
  else{
    data_domepre = getsht2datapre(sht_dome);
    data_tachikawapre = getsht2datasubpre(sht_tachikawa);
    data_shibuyapre = getsht2datasubpre(sht_shibuya);
  }
  var data_all3 = []
  
  for (var h=0;h<lastRow2;h++){
    data_all3[h] = data_dome[h].concat(data_tachikawa[h]).concat(data_shibuya[h])
  }

  var data_all3pre = []
  for (var h=0;h<lastRow2;h++){
    data_all3pre[h] = data_domepre[h].concat(data_tachikawapre[h]).concat(data_shibuyapre[h])
  }

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
    if (all_datan3[y][5] == 0 && all_datan3[y][8] == 0 && all_datan3[y][11] == 0 && data_all3pre[y][4] == 0 && data_all3pre[y][5] == 0 && data_all3pre[y][6] == 0){
      all_datan3[y].splice(0, 13)
      data_all3pre[y].splice(0, 6)
    }
  }

  var all_datan4 = []
  var all_datan4pre = []
  for (var z=0;z<all_datan3.length;z++){
    if (all_datan3[z].length != 0){
      all_datan4.push(all_datan3[z])
      all_datan4pre.push(data_all3pre[z])
    }
  }
  if (!all_datan4.length){
    Browser.msgBox("記入されたデータがありません。オーダーシートを確認してください。", Browser.Buttons.OK)
  }



  var lastRow5 = shtn5.getRange('K:K').getValues().filter(String).length-1;
  var boolean = ash2.getRange("(名前変更不可)夜勤製造表!J2:L999").getValues();
  night_products = [];
  night_products2 = [];
  Logger.log(all_datan3[0][5])
  for (var i=0;i<lastRow5;i++){
    if (boolean[i][2] == true){
      night_products.push([boolean[i][0], 
                           boolean[i][1],
                           ])
    }
    else{
      continue;
    }
  }
  Logger.log(night_products)
    for (var h=0;h<night_products.length;h++){
      num = night_products[h][0]-1
      Logger.log(num)
      if(num == 13 || num == 14 || num == 15 || num == 27 || num == 142 || num == 147){
        night_products2.push([night_products[h][0],
                            night_products[h][1],
                            data_all3[num][4],
                            data_all3[num][5],
                            data_all3[num][6],
                            data_all3[num][4]+data_all3[num][5]+data_all3[num][6],
                            Math.ceil((data_all3[num][4]+data_all3[num][5]+data_all3[num][6])/12),
                            "12の倍数(切り上げ)"])
      }
      else if(num == 120 || num == 145){
        night_products2.push([night_products[h][0],
                            night_products[h][1],
                            data_all3[num][4],
                            data_all3[num][5],
                            data_all3[num][6],
                            data_all3[num][4]+data_all3[num][5]+data_all3[num][6],
                            Math.ceil((data_all3[num][4]+data_all3[num][5]+data_all3[num][6])/11),
                            "11の倍数(切り上げ)"])
      }
      else{
        night_products2.push([night_products[h][0],
                              night_products[h][1],
                              data_all3[num][4],
                              data_all3[num][5],
                              data_all3[num][6],
                              data_all3[num][4]+data_all3[num][5]+data_all3[num][6],
                              "N/A",
                              "N/A"])
      }
  }

  Logger.log(night_products2)
  var lastColumn6 = night_products2[0].length; //カラムの数を取得する
  var lastRow6 = night_products2.length;
  shtn5.getRange(3,1,999,lastColumn6).clear();
  //Logger.log(today_column)
  //Logger.log(all_datan)
  //Logger.log(lastRow)

  //交互の背景色指定
    for (var i=1;i<=night_products.length;i++){
      if(i%2 == 0){
        shtn5.getRange(i+2, 1, 1, lastColumn6).setBackgroundColor('#D3D3D3');
      }
      else{
        shtn5.getRange(i+2, 1, 1, lastColumn6).setBackgroundColor('#FFFFFF');
      }
    }

  shtn5.getRange(3, 1, lastRow6, lastColumn6).setValues(night_products2)
  //Logger.log(night_products2)
  Browser.msgBox("更新が完了しました。OKを押した後、数秒で反映されます。", Browser.Buttons.OK)
}
