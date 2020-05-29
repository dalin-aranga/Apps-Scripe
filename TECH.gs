
function myFunction2(){
   // create work book object
  var wrkBk = SpreadsheetApp.getActiveSpreadsheet();
  
  
  
  // define sheet name
  var wrkSht1 = wrkBk.getSheetByName("Sheet1");
  
  var wrkSht2 = wrkBk.getSheetByName("Sheet2");
  
  
  //get company name
  var company_name1 = wrkSht1.getRange('G'+1).getValue();
  
 
  
  var a;
  
  for (a=2;a<107;a++){
    var company_name2 = wrkSht2.getRange(3,a).getValue();
    
    if (company_name2==company_name1){
      var d1 = wrkSht1.getRange('C'+7).getValue();
      var d2 = wrkSht1.getRange('C'+8).getValue();
      var d3 = wrkSht1.getRange('C'+9).getValue();
      var d4 = wrkSht1.getRange('C'+10).getValue();
      var d5 = wrkSht1.getRange('C'+11).getValue();
      var d6 = wrkSht1.getRange('C'+12).getValue();
      var d7 = wrkSht1.getRange('C'+13).getValue();
    
      var m1 = wrkSht1.getRange('H'+7).getValue();
      var m2 = wrkSht1.getRange('H'+8).getValue();
      var m3 = wrkSht1.getRange('H'+9).getValue();
      var m4 = wrkSht1.getRange('H'+10).getValue();
      var m5 = wrkSht1.getRange('H'+11).getValue();
      var m6 = wrkSht1.getRange('H'+12).getValue();
      var m7 = wrkSht1.getRange('H'+13).getValue();
      var m8 = wrkSht1.getRange('H'+14).getValue();
      var m9 = wrkSht1.getRange('H'+15).getValue();
      var m10 = wrkSht1.getRange('H'+16).getValue();
      var m11 = wrkSht1.getRange('H'+17).getValue();
      var m12 = wrkSht1.getRange('H'+18).getValue();
      var m13 = wrkSht1.getRange('H'+19).getValue();
      var m14 = wrkSht1.getRange('H'+20).getValue();
    
      
    
      wrkSht2.getRange(26,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(26,a).setValue(d1)
      
      wrkSht2.getRange(27,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(27,a).setValue(d2)
      
      wrkSht2.getRange(28,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(28,a).setValue(d3)
      
      wrkSht2.getRange(29,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(29,a).setValue(d4)
      
      wrkSht2.getRange(30,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(30,a).setValue(d5)
      
      wrkSht2.getRange(31,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(31,a).setValue(d6)
      
      wrkSht2.getRange(32,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(32,a).setValue(d7)
      
      wrkSht2.getRange(34,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(34,a).setValue(m1)
      
      
      wrkSht2.getRange(35,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(35,a).setValue(m2)
      
      wrkSht2.getRange(36,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(36,a).setValue(m3)
      
      wrkSht2.getRange(37,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(37,a).setValue(m4)
      
      wrkSht2.getRange(38,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(38,a).setValue(m5)
      
      wrkSht2.getRange(39,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(39,a).setValue(m6)
      
      wrkSht2.getRange(40,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(40,a).setValue(m7)
      
      wrkSht2.getRange(41,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(41,a).setValue(m8)
      
      wrkSht2.getRange(42,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(42,a).setValue(m9)
      
      wrkSht2.getRange(43,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(43,a).setValue(m10)
      
      
      wrkSht2.getRange(44,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(44,a).setValue(m11)
      
      wrkSht2.getRange(45,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(45,a).setValue(m12)
      
      wrkSht2.getRange(46,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(46,a).setValue(m13)
      
      wrkSht2.getRange(47,a).activate();
      wrkSht2.getActiveRangeList().clear({contentsOnly : true, skipFilteredRow : true});
      wrkSht2.getRange(47,a).setValue(m14)
      
      break;
    }
    else{
      continue;
     
    }
      
    
  
  }
  


}
