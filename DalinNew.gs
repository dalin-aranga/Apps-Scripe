// you can any variable nam change you want
//blue color words are variable



function myFunctio(){
  
  // create work book object
  var wrkBk = SpreadsheetApp.getActiveSpreadsheet();
  
  // define sheet name
  var wrkSht = wrkBk.getSheetByName("Sheet1");
  
  //json file url first part
  var url_first = "https://widget3.zacks.com/data/chart/json/"
  
  //json file url second part
  var url_last = "/pe_ratio/www.zacks.com"
  
  //get company name
  var company_name = wrkSht.getRange('G'+1).getValue();
  Utilities.sleep(1000);
  
  // full url
  var url = url_first + company_name + url_last
  
  //fetch the url
  var res = UrlFetchApp.fetch(url);
  
  //get text content 
  var content = res.getContentText();
  
  //loads json format
  var json = JSON.parse(content);
  
  //get want keys
  var daily_pe_ratio = json["daily_pe_ratio"];
  var monthly_pe_ratio = json["monthly_pe_ratio"];
  
  
  // create variable for runing loop
  var i = 0;
  
  // create arry to store in data
  var arrdate = [];
  var arrprice = [];
  
  for (var key in daily_pe_ratio) {
    // get tha date in daily pe ratio
    arrdate.push(key);
    
    if (i == 6) { break; }
    
    i = i+1;
    
}  
  
   var j=0;
   for (j = 0; j < 7; j++) {
     //get the value in daily pe ratio
     var daily_pe_ratio_new = json["daily_pe_ratio"][arrdate[j]];
     arrprice.push(daily_pe_ratio_new);

} 
  
  
  
  
  var k = 0;
  
  // create arry to store in data
  var arrdatem = [];
  var arrpricem = [];
  
  for (var keym in monthly_pe_ratio) {
    //get the date in monthly pe ratio
    arrdatem.push(keym);
    
    if (k == 13) { break; }
    
    k = k+1;
    
}  
   var n=0;
   for (n = 0; n < 14; n++) {
     //get the value in monthly pe ratio
     var daily_pe_ratio_newm = json["monthly_pe_ratio"][arrdatem[n]];
     arrpricem.push(daily_pe_ratio_newm);

} 
  
 
  wrkSht.getRange("B"+7).activate();
  // first sheet location clear
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  // set all date and values into sheet loacation
  wrkSht.getRange('B' +7).setValue(arrdate[0])
  
  wrkSht.getRange("B"+8).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('B' +8).setValue(arrdate[1])
  
  wrkSht.getRange("B"+9).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('B' +9).setValue(arrdate[2])
  
  wrkSht.getRange("B"+10).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('B' +10).setValue(arrdate[3])
  
  wrkSht.getRange("B"+11).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('B' +11).setValue(arrdate[4])
  
  wrkSht.getRange("B"+12).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('B' +12).setValue(arrdate[5])
  
  wrkSht.getRange("B"+13).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('B' +13).setValue(arrdate[6])
  
  
  wrkSht.getRange("C"+7).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('C' +7).setValue(arrprice[0])
  
  wrkSht.getRange("C"+8).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('C' +8).setValue(arrprice[1])
  
  wrkSht.getRange("C"+9).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('C' +9).setValue(arrprice[2])
  
  wrkSht.getRange("C"+10).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('C' +10).setValue(arrprice[3])
  
  wrkSht.getRange("C"+11).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('C' +11).setValue(arrprice[4])
  
  wrkSht.getRange("C"+12).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('C' +12).setValue(arrprice[5])
  
  wrkSht.getRange("C"+13).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('C' +13).setValue(arrprice[6])
  
  
  
  wrkSht.getRange("G"+7).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +7).setValue(arrdatem[0])
  
  wrkSht.getRange("G"+8).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +8).setValue(arrdatem[1])
  
  wrkSht.getRange("G"+9).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +9).setValue(arrdatem[2])
  
  wrkSht.getRange("G"+10).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +10).setValue(arrdatem[3])
  
  wrkSht.getRange("G"+11).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +11).setValue(arrdatem[4])
  
  wrkSht.getRange("G"+12).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +12).setValue(arrdatem[5])
  
  wrkSht.getRange("G"+13).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +13).setValue(arrdatem[6])
  
  wrkSht.getRange("G"+14).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +14).setValue(arrdatem[7])
  
  wrkSht.getRange("G"+15).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +15).setValue(arrdatem[8])
  
  wrkSht.getRange("G"+16).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +16).setValue(arrdatem[9])
  
  wrkSht.getRange("G"+17).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +17).setValue(arrdatem[10])
  
  wrkSht.getRange("G"+18).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +18).setValue(arrdatem[11])
  
  wrkSht.getRange("G"+19).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +19).setValue(arrdatem[12])
  
  wrkSht.getRange("G"+20).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('G' +20).setValue(arrdatem[13])
  
  
  
  
  wrkSht.getRange("H"+7).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +7).setValue(arrpricem[0])
  
  wrkSht.getRange("H"+8).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +8).setValue(arrpricem[1])
  
  wrkSht.getRange("H"+9).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +9).setValue(arrpricem[2])
  
  wrkSht.getRange("H"+10).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +10).setValue(arrpricem[3])
  
  wrkSht.getRange("H"+11).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +11).setValue(arrpricem[4])
  
  wrkSht.getRange("H"+12).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +12).setValue(arrpricem[5])
  
  wrkSht.getRange("H"+13).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +13).setValue(arrpricem[6])
  
  wrkSht.getRange("H"+14).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +14).setValue(arrpricem[7])
  
  wrkSht.getRange("H"+15).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +15).setValue(arrpricem[8])
  
  wrkSht.getRange("H"+16).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +16).setValue(arrpricem[9])
  
  wrkSht.getRange("H"+17).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +17).setValue(arrpricem[10])
  
  wrkSht.getRange("H"+18).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +18).setValue(arrpricem[11])
  
  wrkSht.getRange("H"+19).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +19).setValue(arrpricem[12])
  
  wrkSht.getRange("H"+20).activate();
  wrkSht.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
  wrkSht.getRange('H' +20).setValue(arrpricem[13])
  
  
  
  
  
 
  
    
}
  
  
    

  
  
  

  
  
  


