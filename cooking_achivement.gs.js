function myFunction3(){
  var ash_3 = SpreadsheetApp.getActiveSpreadsheet();
  var shtn_3 = ash_3.getSheetByName("(名前変更不可)納品書");

  today_3 = shtn_3.getRange("K1").getValue();
  shop_3 = shtn_3.getRange("A1").getValue();

  if (shop_3 == "BA立川"){
    achiv_sht = ash_3.getSheetByName("(名前変更不可)製造実績立川")
  }
  else if(shop_3 == "BA東急渋谷"){
    achiv_sht = ash_3.getSheetByName("(名前変更不可)製造実績東急渋谷")
  }
  

  todayformatted_3 = Utilities.formatDate(today_3, "JST", "yyyy/MM/dd");

  const data_day_4 = ash_3.getRange("admin!E2:F33").getValues();




  for (var i=0;i<data_day_4.length;i++){
    if (Utilities.formatDate(data_day_4[i][0], "JST", "yyyy/MM/dd") == todayformatted_3){
      var today_column_4 = data_day_4[i][1]
      break
    }
  }
  //achiv_sht.getRange(5, today_column_4+5, 950, 2).clear();
  var lastRow_3 = shtn_3.getRange('B:B').getValues().filter(String).length;
  var all_datan_3 = shtn_3.getRange(5, 1, lastRow_3, 13).getValues();
  var lastRow_4 = achiv_sht.getRange('B:B').getValues().filter(String).length;
  var all_datan_ach = achiv_sht.getRange(5, 1, lastRow_4, 2).getValues();
  //Logger.log(all_datan_3)
  if (!all_datan_3.length){
    Browser.msgBox("記入されたデータがありません。オーダーシートを確認してください。", Browser.Buttons.OK)
  }
  var lastColumn_3 = all_datan_3[0].length; //カラムの数を取得する
  //Logger.log(all_datan_3)
  var lastRow_3 = all_datan_3.length;   //行の数を取得する
  //achiv_sht.getRange(5, today_column_4+5, 950, 2).clear();
  for (var i=0; i<all_datan_3.length;i++){
    if (all_datan_3[i][12] != ''){
      achiv_sht.getRange(all_datan_3[i][0]+4, today_column_4+5, 1, 1).setValue(all_datan_3[i][12])
      achiv_sht.getRange(all_datan_3[i][0]+4, today_column_4+6, 1, 1).setValue(all_datan_3[i][12]*all_datan_3[i][3])
    }
    if (all_datan_3[i][12] == 0){
      achiv_sht.getRange(all_datan_3[i][0]+4, today_column_4+5, 1, 1).setValue('')
      achiv_sht.getRange(all_datan_3[i][0]+4, today_column_4+6, 1, 1).setValue('')
    }
    if (all_datan_3[i][12] === '') {
      achiv_sht.getRange(all_datan_3[i][0]+4, today_column_4+5, 1, 1).setValue(all_datan_3[i][10])
      achiv_sht.getRange(all_datan_3[i][0]+4, today_column_4+6, 1, 1).setValue(all_datan_3[i][10]*all_datan_3[i][3])
    }
    if (all_datan_3[i][12] === 0){
      achiv_sht.getRange(all_datan_3[i][0]+4, today_column_4+5, 1, 1).setValue(0)
      achiv_sht.getRange(all_datan_3[i][0]+4, today_column_4+6, 1, 1).setValue(0)
    }
  }
  achiv_sht.getRange(4, today_column_4+5, 1, 1).setValue("製造実績数")
  achiv_sht.getRange(4, today_column_4+6, 1, 1).setValue("仕入金額")
  //交互の背景色指定
  for (var i=1;i<=990;i++){
    if(i%2 == 0){
      achiv_sht.getRange(i+4, 1, 1, 69).setBackgroundColor('#D3D3D3');
    }
    else{
      achiv_sht.getRange(i+4, 1, 1, 69).setBackgroundColor('#FFFFFF');
    }
  }
  Browser.msgBox("製造実績シートへの入力が完了しました。", Browser.Buttons.OK)

}
