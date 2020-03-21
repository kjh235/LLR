function test1223(){
  var unreads = GmailApp.search("is:unread from:(noreply@lularoebless.com) subject:(Purchase Receipt from LuLaRoe - Order Number)");
  var message = unreads[0].getMessages();
  Logger.log(message[0].getBody());
  
  var unreads = GmailApp.search("is:unread from:(noreply@lularoe.com) subject:(LuLaRoe Web Order Confirmation)");
  var message = unreads[0].getMessages();
  Logger.log(message[0].getBody());
}




function RemoveFile(filetoremove) {
  var files = DriveApp.getFilesByName(filetoremove);
  //var filesearch = DriveApp.searchFiles("name = 'invoices' and (mimeType = 'application/vnd.google-apps.spreadsheet'");
  while (files.hasNext()) {
    var file = files.next();
    file.setTrashed(true);    
  }
};

function downloadSheetAsXlsx(sheetID, fileNm, fileType) {
  
  var spreadSheet = SpreadsheetApp.openById(sheetID); 
  var ssID = spreadSheet.getId();
  var fullname = fileNm + "." + fileType;
  if(fileType == "csv")
  {
    var sheet1 = spreadSheet.getSheetByName(fileNm);
    var sheet2 = spreadSheet.setActiveSheet(sheet1);
    var shtID = sheet2.getSheetId();
    var url = "https://docs.google.com/spreadsheets/d/"+ssID+"/export?format="+fileType+"&gid="+shtID;
  }
  else
  {var url = "https://docs.google.com/spreadsheets/d/"+ssID+"/export?format="+fileType}  
  var params = {method:"GET", headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  var response = UrlFetchApp.fetch(url, params);
  
  // remove current version from drive
  RemoveFile(fullname);
  // add new version to drive
  DriveApp.createFile(response).setName(fullname);
  
};


function CSVToArray( strData, strDelimiter ) {
  // Check to see if the delimiter is defined. If not,
  // then default to COMMA.
  strDelimiter = (strDelimiter || ",");
  
  // Create a regular expression to parse the CSV values.
  var objPattern = new RegExp(
    (
      // Delimiters.
      "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +
      
      // Quoted fields.
      "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +
      
      // Standard fields.
      "([^\"\\" + strDelimiter + "\\r\\n]*))"
    ),
    "gi"
  );
  
  // Create an array to hold our data. Give the array
  // a default empty first row.
  var arrData = [[]];
  
  // Create an array to hold our individual pattern
  // matching groups.
  var arrMatches = null;
  
  // Keep looping over the regular expression matches
  // until we can no longer find a match.
  while (arrMatches = objPattern.exec( strData )){
    
    // Get the delimiter that was found.
    var strMatchedDelimiter = arrMatches[ 1 ];
    
    // Check to see if the given delimiter has a length
    // (is not the start of string) and if it matches
    // field delimiter. If id does not, then we know
    // that this delimiter is a row delimiter.
    if (
      strMatchedDelimiter.length &&
      (strMatchedDelimiter != strDelimiter)
      ){
        
        // Since we have reached a new row of data,
        // add an empty row to our data array.
        arrData.push( [] );
        
      }
    
    // Now that we have our delimiter out of the way,
    // let's check to see which kind of value we
    // captured (quoted or unquoted).
    if (arrMatches[ 2 ]){
      
      // We found a quoted value. When we capture
      // this value, unescape any double quotes.
      var strMatchedValue = arrMatches[ 2 ].replace(
        new RegExp( "\"\"", "g" ),
        "\""
      );
      
    } else {
      
      // We found a non-quoted value.
      var strMatchedValue = arrMatches[ 3 ];
      
    }
    
    // Now that we have our value string, let's add
    // it to the data array.
    arrData[ arrData.length - 1 ].push( strMatchedValue );
  }
  
  // Return the parsed data.
  return( arrData );
};
// http://www.bennadel.com/blog/1504-Ask-Ben-Parsing-CSV-Strings-With-Javascript-Exec-Regular-Expression-Command.htm
// This will parse a delimited string into an array of
// arrays. The default delimiter is the comma, but this
// can be overriden in the second argument.


function importDatafromRoot(fileNm, destWorkbookID, destSheetNm) {
  var fSource = DriveApp.getRootFolder(); 
  // reports_folder_id = id of folder where csv reports are saved
  var fi = fSource.getFilesByName(fileNm); 
  // latest report file
  
  var ss = SpreadsheetApp.openById(destWorkbookID);
  // data_sheet_id = id of spreadsheet that holds the data to be updated with new report data
  //ss.deleteSheet(newsheet);
  if ( fi.hasNext() ) { // proceed if "report.csv" file exists in the reports folder
    var file = fi.next();
    var csv = file.getBlob().getDataAsString();
    var csvData = CSVToArray(csv); // see below for CSVToArray function
    var centralStore = ss.getSheetByName (destSheetNm);
    centralStore.clear();
    
    var calPaste = centralStore.getRange("A1");
    if(ss.getSheetByName ("NEWDATA") != null){
      var newsheet = ss.setActiveSheet(ss.getSheetByName("NEWDATA"));  
    }
    else{
      var newsheet = ss.insertSheet('NEWDATA');
    }
    var newData = ss.getSheetByName ("NEWDATA");
    var calRange = newData.getRange("A1:N1000");
    //var newsheet = ss.insertSheet('NEWDATA'); // create a 'NEWDATA' sheet to store imported data
    // loop through csv data array and insert (append) as rows into 'NEWDATA' sheet
    for ( var i=0, lenCsv=csvData.length; i<lenCsv; i++ ) {
      newsheet.getRange(i+1, 1, 1, csvData[i].length).setValues(new Array(csvData[i]));
    }
    
    /*
    ** report data is now in 'NEWDATA' sheet in the spreadsheet - process it as needed,
    ** then delete 'NEWDATA' sheet using ss.deleteSheet(newsheet)
    */
    
    
    
    var calPaste = centralStore.getRange("A1");  
    
    calRange.copyTo(calPaste, {contentsOnly: true})
    
    ss.setActiveSheet(ss.getSheetByName("NEWDATA"));
    ss.deleteActiveSheet();
    
    ss.setActiveSheet(ss.getSheetByName(destSheetNm));
    
  }
};





function appendRowtoSheet(workbookID, sheetNm, arraytoAppend){
  var spreadsheet = SpreadsheetApp.openById(workbookID);
  var a1sheet = spreadsheet.getSheetByName(sheetNm);
  var sheet = spreadsheet.setActiveSheet(a1sheet);
  
  //sheet.appendRow(arraytoAppend);
  /*
  var form_range = "Invoices!S2:U2" ;
  
  //var cell = sheet.getRange(form_range);
  //cell.setFormulasR1C1(cell.getValues());
  sheet.insertRowAfter(1);
  //sheet.setActiveSelection("A2:U2");
  var range = sheet.getRange("Invoices!A2:U2").setValue(arraytoAppend);
  //range.setValues(arraytoAppend);
  
  //var range = sheet.getRange("S2:U"+(arraytoAppend.length+1));
  //range.setValues(arraytoAppend);
  var range2 = sheet.getRange(form_range);
  //range.setValues(arraytoAppend);
  var formulas1 = range2.getValues();
  range2.setFormulasR1C1(formulas1);*/
  sheet.insertRowAfter(1);
  for( var x = 0; x < arraytoAppend.length;x++){
    //var cell = sheet.getRange(2, x + 1);
    var range = sheet.getRange(2, x + 1).setValue(arraytoAppend[x]);
  }
  
  
};


function appendRowstoSheet(workbookID, sheetNm, arraytoAppend){
  var spreadsheet = SpreadsheetApp.openById(workbookID);
  var a1sheet = spreadsheet.getSheetByName(sheetNm);
  var sheet = spreadsheet.setActiveSheet(a1sheet);
  
  for( var x = 0; x < arraytoAppend.length;x++){
    sheet.appendRow(arraytoAppend[x]); 
  }
  /*
  var form_range = "S2:U"+(arraytoAppend.length+1);
  sheet.insertRowsAfter(1,arraytoAppend.length);
  
  var range = sheet.getRange(form_range);
  range.setValues(arraytoAppend);
  range.setFormulasR1C1(range.getValues());
  */
  
};

function clearSheet(workbookID, sheetNm){
  var spreadsheet = SpreadsheetApp.openById(workbookID);
  var a1sheet = spreadsheet.getSheetByName(sheetNm);
  var sheet = spreadsheet.setActiveSheet(a1sheet);
  sheet.clear();
  
};




function getWorkbookColumnData(workbookID, sheetNm, colNum){
  
  var ss = SpreadsheetApp.openById(workbookID);
  var sheet = ss.getSheetByName(sheetNm);
  var last_row = ss.getLastRow();
  if(last_row == null || last_row == 0)
  {
    return "";
  }
  else{
    var colB = [];
    colB = sheet.getRange(2, colNum, last_row).getValues();
    return colB;
  }
  
};


function getOrderSummaryFromMessage(emailmsg){
  var ordr = emailmsg.getBody();
  var emlbody = emailmsg.getPlainBody();
  Logger.log(ordr);
  Logger.log(emlbody);
  var order_dt = emailmsg.getDate();
  var customer1 = "<h3>";
  var customer2 = "</h3>";
  var cust_pos1 = ordr.search(customer1);
  var cust_pos2 = ordr.search(customer2);
  var cust_nm = ordr.substring(cust_pos1+4,cust_pos2);
  
  var order1 = "<h4>";
  var order2 = "</h4>";
  var ordr_pos1 = ordr.search(order1);
  var ordr_pos2 = ordr.search(order2);
  var ordr_nmbr = ordr.substring(ordr_pos1+4,ordr_pos2);   
  var ordr_num = ordr_nmbr.substr(ordr_nmbr.search("#")+1,8);
  var shpname = "";
  var addr1 = "";
  var addr2 = "";
  var city = "";
  var state = "";
  var zip = "";
  var cityst = "";
  var billname = "";
  var billaddr1 = "";
  var billaddr2 = "";
  var billcity = "";
  var billcityst = "";
  var billstate = "";
  var billzip = "";
  
  var amt_indx = ordr.search("Amount");
  var mdl_indx = ordr.search("Model");
  var tp_indx = ordr.search("Total Paid");
  var ship_indx = ordr.search("Shipping");
  var bill_indx = ordr.search("Billing");
  var shipinfo = false;
  
  // if shipping information is in the email, parse the shipping info
  if((ship_indx > 0 ) && (bill_indx > 0 ) && (bill_indx > ship_indx))
  {
    
    shipinfo = true;
    var shp_tbl = ordr.substring(ship_indx,bill_indx);
    var shp = shp_tbl.split("/td></tr>");
    shpname = shp[0].substring(shp[0].lastIndexOf(">")+1,shp[0].lastIndexOf("<"));
    addr1 = shp[1].substring(shp[1].lastIndexOf(">")+1,shp[1].lastIndexOf("<"));
    addr2 = shp[2].substring(shp[2].lastIndexOf(">")+1,shp[2].lastIndexOf("<"));
    city = shp[3].substring(shp[3].lastIndexOf(">")+1,shp[3].lastIndexOf("<"));
    cityst = city.split(",");
    city = cityst[0];
    state = cityst[1];
    zip = shp[4].substring(shp[4].lastIndexOf(">")+1,shp[4].lastIndexOf("<"));
  }
  
  // if the billing information is in the email, parse the billing info
  if((ordr.search("Billing")) > 0){
    
    var bill_tbl = ordr.substring(bill_indx,amt_indx);
    var bill = bill_tbl.split("/td></tr>");
    billaddr1 = bill[0].substring(bill[0].lastIndexOf(">")+1,bill[0].lastIndexOf("<"));
    billaddr2 = bill[1].substring(bill[1].lastIndexOf(">")+1,bill[1].lastIndexOf("<"));
    billcity = bill[2].substring(bill[2].lastIndexOf(">")+1,bill[2].lastIndexOf("<"));
    billcityst = billcity.split(",");
    billcity = billcityst[0];
    billstate = billcityst[1];
    billzip = bill[3].substring(bill[3].lastIndexOf(">")+1,bill[3].lastIndexOf("<"));         
  }
  
  if(!shipinfo){
    shpname = cust_nm;
    addr1 = billaddr1;
    addr2 = billaddr2;
    city = billcity;
    state = billstate;
    zip = billzip;
  }
  
  
  if(ordr.search("Discount") >= 0)
  {
    //Logger.log("There is a discount");
    var disc_indx = ordr.search("Discount");
    var tp_indx = ordr.search("Total Paid");
    
    var amnt_tbl = ordr.substring(amt_indx,disc_indx);
    var amt = amnt_tbl.split("/td>");
    var amt_nb = amt[0].substring(amt[0].lastIndexOf(";")+1,amt[0].lastIndexOf("<"));
    Logger.log(amt[0]);
    var disc_tbl = ordr.substring(disc_indx,tp_indx);
    var disc = disc_tbl.split("/td>");
    var disc_nb = disc[0].substring(disc[0].lastIndexOf(";")+1,disc[0].lastIndexOf("<"));
    
    var tp_tbl = ordr.substring(tp_indx,mdl_indx);
    var tp = tp_tbl.split("/td>");
    var tp_nb = tp[0].substring(tp[0].lastIndexOf(";")+1,tp[0].lastIndexOf("<"));
  }// ending if 
  else
  {
    // Logger.log("There is no Discount");
    disc_nb = "0";
    var tp_indx = ordr.search("Total Paid");
    var tp_tbl = ordr.substring(tp_indx,mdl_indx);
    var tp = tp_tbl.split("/td>");
    var tp_nb = tp[0].substring(tp[0].lastIndexOf(";")+1,tp[0].lastIndexOf("<"));
    amt_nb = tp_nb;
  }// ending else
  
  var items_end = ordr.search("<!-- End of wrapper table -->");
  var items_tbl = ordr.substring(mdl_indx,items_end);
  var itms = items_tbl.split("/tr>");
  var hdr = itms[0];
  var itmz = itms[1];
  var footr = itms[2];
  
  
  
  
  
  /*************** get footer summary row */
  var footr1 = footr.split("/b");
  for ( var h = 0; h < footr1.length - 1; h++){
    if((footr1[h].lastIndexOf(">")+1) == footr1[h].lastIndexOf("<")){
      // Do nothing 
    } // ending if 
    else{
      var footr_1 = footr1[h].substring(footr1[h].lastIndexOf(">") + 1,footr1[h].lastIndexOf("<"));
    }// ending else 
    
  }// ending h for loop
  
  
  var order_summary = [];
  order_summary[0] = cust_nm;
  order_summary[1] = ordr_num;
  order_summary[2] = shpname;
  order_summary[3] = addr1;
  order_summary[4] = addr2;
  order_summary[5] = city;
  order_summary[6] = state; 
  order_summary[7] = zip;
  order_summary[8] = order_dt;
  order_summary[9] = footr_1;
  order_summary[10] = amt_nb;
  order_summary[11] = disc_nb;
  order_summary[12] = tp_nb;
  
  return order_summary;
  
  
  
};


function getOrderDetailsFromMessage(emailmsg){
  var ordr = emailmsg.getBody();
  Logger.log(ordr);
  var amt_indx = ordr.search("Amount");
  var mdl_indx = ordr.search("Model");
  var tp_indx = ordr.search("Total Paid");
  var items_end = ordr.search("<!-- End of wrapper table -->");
  var items_tbl = ordr.substring(mdl_indx,items_end);
  var itms = items_tbl.split("/tr>");
  var hdr = itms[0];
  var itmz = itms[1];
  var order1 = "<h4>";
  var order2 = "</h4>";
  var ordr_pos1 = ordr.search(order1);
  var ordr_pos2 = ordr.search(order2);
  var ordr_nmbr = ordr.substring(ordr_pos1+4,ordr_pos2);   
  var ordr_num = ordr_nmbr.substr(ordr_nmbr.search("#")+1,8);
  /********** get items table header */
  var hdr1 = hdr.split("/t");
  var hdr_2 = [];
  hdr1[0] = hdr1[0] + "/b";
  var hdr2 = hdr1[0] + hdr1[1] + hdr1[2] + "/b";
  hdr2 = hdr2.split("/b");
  
  for( var j = 0; j < hdr1.length - 1 ; j++){
    
    if((hdr1[j].lastIndexOf("<b>")+3) == hdr1[j].lastIndexOf("</b")){
    }// ending if
    else{
      var hdr_1 = hdr1[j].substring(hdr1[j].lastIndexOf("<b>") + 3,hdr1[j].lastIndexOf("</b><"));
      
    }// ending else
    
    hdr_2[j] = (hdr_1);
    hdr_2[0] = "Model";
  }// ending J for loop 
  
  hdr_2.splice(0,0,ordr_num);
  
  var itmz1 = itmz.split("/td><tr>");
  var itmz_1;
  var itmz_2;
  var x;
  var itmz_array1 = [];
  var itmz_array2 = [];
  
  itmz_array2.push(hdr_2);
  for( var k = 0; k < itmz1.length; k++){
    
    itmz_1 = itmz1[k].split("/td");
    var itmz_array1 = [];
    for( var j = 0; j < hdr_2.length-1; j++){
      if(k==0){
        x=j;
      }
      if((itmz_1[j].lastIndexOf(">")+1) == itmz_1[j].lastIndexOf("<") ){
        
        if((itmz_1[j].length>=2) && (j<=x)){
          itmz_2 = "0";
        }  
      }
      else{
        itmz_2 = itmz_1[j].substring(itmz_1[j].lastIndexOf(">") + 1,itmz_1[j].lastIndexOf("<"));
        
      }
      
      itmz_array1.push(itmz_2);
    }//ending j for loop 
    
    itmz_array1.splice(0,0,ordr_num);
    
    itmz_array2.push(itmz_array1);
    
  } //ending k for loop 
  var order_itms = [];
  var all_ordr_itms = [];
  var len = itmz_array2.length;
  
  for (var p =1;p<len; p++){
    for( var q =2;q<itmz_array2[p].length;q++){
      if(itmz_array2[p][q]!=0){
        order_itms.push(ordr_num);
        order_itms.push(itmz_array2[p][q]);
        order_itms.push(itmz_array2[p][1]);
        order_itms.push(itmz_array2[0][q]);
        all_ordr_itms.push(order_itms);
      }// ending if statement
    }// ending Q for loop
    order_itms = [];
  } // ending P for loop
  return all_ordr_itms;
};

function removeDuplicates(activesheet) {
  //var sheet = SpreadsheetApp.getActiveSheet();
  var data = activesheet.getDataRange().getValues();
  var newData = new Array();
  for(i in data){
    var row = data[i];
    var duplicate = false;
    for(j in newData){
      if(row.join() == newData[j].join()){
        duplicate = true;
      }
    }
    if(!duplicate){
      newData.push(row);
    }
  }
  activesheet.clearContents();
  activesheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
};

function initialSetup(){
  var invoiceexists = false;
  var workbooks = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
  
  while (workbooks.hasNext()) {
    var sheet = workbooks.next();
    if(sheet.getName() == "Invoices"){
      invoiceexists = true;
      //var invoiceSheet = sheet.getId();
      //Logger.log(invoiceSheet);
      return sheet.getId();
    } //  end if
  } // end while 
  if(!invoiceexists){
    var newInvoice = DriveApp.createFile('Invoices','',MimeType.GOOGLE_SHEETS);
    return newInvoice.getId();
  }
  
  
};




function LLRPurchaseOrder(label){
  //var label = GmailApp.search("to:( lularoekristenhickman@gmail.com) from:(noreply@lularoe.com) subject:(LuLaRoe Web Order Confirmation) newer_than:5d");
  
  var thread = label;
  var orders = [];
  var all_itms_array = [];
  var orders2 = [];
  var all_itms_array2 = [];
  var t=0;
  
  var itms_hdr = [];
  for (var i = 0; i < thread.length; i++) {
    
    var message = thread[i].getMessages();
    for (var z = 0; z < message.length; z++) {
      t++;
      var all_hdr_info2 = [];
      var all_hdr_info = [];
      var ordr_itms_randc = [];
      var order_summary = []; 
      var itms = [6];
      var all_itms = [];
      
      
      var ordr = message[z].getBody();
      var order_dt =  message[z].getDate();
      var ship2_pos = ordr.search("SHIP TO");
      var pymt_pos = ordr.search("Payment");
      var qnty_pos = ordr.search("Quantity");
      var subtotal_pos = ordr.search("Subtotal");
      var endofmsg_pos = ordr.search("</P>");
      
      var ordr_hdr = ordr.substring(ship2_pos,pymt_pos);
      var ordr_itms = ordr.substring(qnty_pos-1,subtotal_pos);
      var ordr_summ = ordr.substring(subtotal_pos-1, ordr.length);
      
      var endofmsg_pos = ordr_summ.search("</table>");
      
      ordr_summ = ordr_summ.substring(0,endofmsg_pos);
      
      // --- here starts the order header parsing
      var ordr_hdr_split = ordr_hdr.split("</td");
      var ordr_hdr_split2 = ordr_hdr_split[1].split("<BR>");
      var ordr_hdr_split3 = ordr_hdr_split[3].split("<BR>");
      var ordr_hdr_split4 = ordr_hdr_split[4].split("<br>");
      
      var ship_nm = ordr_hdr_split2[1];
      var ship_addr = ordr_hdr_split2[2];
      var ship_cityst = ordr_hdr_split2[3];
      
      
      
      all_hdr_info.push(ordr_hdr_split3[0].substring(ordr_hdr_split3[0].lastIndexOf(">") + 1 , ordr_hdr_split3[0].length));
      all_hdr_info.push(ordr_hdr_split3[1]);
      all_hdr_info.push(ordr_hdr_split3[2]);
      all_hdr_info.push(ordr_hdr_split3[3]);
      all_hdr_info.push(ordr_hdr_split3[4]);
      all_hdr_info.push(ordr_hdr_split3[5]);
      all_hdr_info.push(ordr_hdr_split3[6]);
      all_hdr_info.push(ordr_hdr_split3[7]);
      all_hdr_info.push(ordr_hdr_split3[8]);
      all_hdr_info.push(ordr_hdr_split3[9]);
      //Logger.log(all_hdr_info);
      
      
      all_hdr_info2.push(ordr_hdr_split4[0].substring(ordr_hdr_split4[0].lastIndexOf(">") + 1 , ordr_hdr_split4[0].length));
      var ordr_nmbr = ordr_hdr_split4[0].substring(ordr_hdr_split4[0].lastIndexOf(">") + 1 , ordr_hdr_split4[0].length);
      all_hdr_info2.push(ordr_hdr_split4[1]);
      all_hdr_info2.push(ordr_hdr_split4[2]);
      all_hdr_info2.push(ordr_hdr_split4[3]);
      all_hdr_info2.push(ordr_hdr_split4[4]);
      all_hdr_info2.push(ordr_hdr_split4[5]);
      all_hdr_info2.push(ordr_hdr_split4[6]);
      all_hdr_info2.push(ordr_hdr_split4[7]);
      all_hdr_info2.push(ordr_hdr_split4[8]);
      all_hdr_info2.push(ordr_hdr_split4[9]);
      //Logger.log(all_hdr_info2);
      
      // --- here starts the individual items parsing
      //get items table header
      
      
      
      var itms_hdr = [];
      var ordr_itms_hdr = ordr_itms.substring(0,ordr_itms.search("</tr"));
      var ordr_itms_hdr_array = ordr_itms_hdr.split("/th");
      for(var d = 0; d < ordr_itms_hdr_array.length - 1; d++){
        itms_hdr[d] = ordr_itms_hdr_array[d].substring(ordr_itms_hdr_array[d].lastIndexOf(">") + 1 , ordr_itms_hdr_array[d].lastIndexOf("<"));}
      itms_hdr.splice(0,0,"Order Dt");
      itms_hdr.splice(0,0,"Order#");
      itms_hdr.push("Actual Pieces");
      
      
      var ordr_itms_data = ordr_itms.substring(ordr_itms.search("</tr"),ordr_itms.length)
      var ordr_itms_rows = ordr_itms_data.split("/td></tr");
      Logger.log(ordr_itms_rows);
      var ordr_itms_randc = [];
      for(var y = 0; y < ordr_itms_rows.length - 1; y++){
        ordr_itms_randc = ordr_itms_rows[y].split("/td");
        Logger.log(ordr_itms_randc);
        for(var x = 0; x < ordr_itms_randc.length;x++){
          
          ordr_itms_randc[x]= ordr_itms_randc[x].substring(ordr_itms_randc[x].lastIndexOf(">") + 1 , ordr_itms_randc[x].lastIndexOf("<"));
          
          
          
        }// end x for loop
        if(ordr_itms_randc[0]== "")
        {
          //Logger.log("ZERO");
          ordr_itms_randc.splice(0,1);
        }
        
        ordr_itms_randc.splice(0,0,order_dt);
        ordr_itms_randc.splice(0,0,ordr_nmbr);
        all_itms.push(ordr_itms_randc);
        //Logger.log(ordr_itms_randc);
      } // end y for loop
      
      
      
      // } // end z for loop
      
      // --- here starts the order summary parsing
      var ordr_summ_rows = ordr_summ.split("/td></tr");
      //Logger.log(ordr_summ_rows);
      //var order_summary = [];
      for(var r = 0; r < ordr_summ_rows.length - 1; r++){
        var ordr_summ_randc = [];
        var ordr_summ_rows2 = ordr_summ_rows[r].split("/font");
        //Logger.log(ordr_summ_rows2);
        for(var j = 0; j < ordr_summ_rows2.length;j++){
          ordr_summ_randc[j]= ordr_summ_rows2[j].substring(ordr_summ_rows2[j].lastIndexOf(">") + 1 , ordr_summ_rows2[j].lastIndexOf("<"));
          if (ordr_summ_randc[j] == ""){
            ordr_summ_randc[j]= ordr_summ_rows2[j].substring(ordr_summ_rows2[j].lastIndexOf("<b>") + 3 , ordr_summ_rows2[j].lastIndexOf("</b>"));
          }
          //Logger.log(ordr_summ_randc[x]);
        }// end x for loop
        //Logger.log(ordr_summ_randc);
        
        order_summary.push(ordr_summ_randc);
      } // end y for loop
      for( var w = 0; w < order_summary.length; w++){
        if(order_summary[w][0] == "Subtotal:"){
          all_hdr_info.push(order_summary[w][0]);
          all_hdr_info2.push(order_summary[w][1]);}
        if(order_summary[w][0] == "Shipping:"){
          all_hdr_info.push(order_summary[w][0]);
          all_hdr_info2.push(order_summary[w][1]);}
        if(order_summary[w][0] == "Taxes:"){
          all_hdr_info.push(order_summary[w][0]);
          all_hdr_info2.push(order_summary[w][1]);}
        if(order_summary[w][0] == "Total:"){
          all_hdr_info.push(order_summary[w][0]);
          all_hdr_info2.push(order_summary[w][1]);}
        if(order_summary[w][0] == "Amount Paid:"){
          all_hdr_info.push(order_summary[w][0]);
          all_hdr_info2.push(order_summary[w][1]);}
        if(order_summary[w][0] == "Total Pieces:"){
          all_hdr_info.push(order_summary[w][0]);
          all_hdr_info2.push(order_summary[w][1]);}   
      }
      
      orders.push(all_hdr_info2);
      all_itms_array.push(all_itms);
    }// end z for loop
    thread[i].markRead();
    
  } // end i for loop
  
  
  
  //appendRowtoSheet("14VzdU_0cZZ_x4Ab3zTF1SCUS5mTYZEdkZxhLFae2T0w","PurchaseOrders",all_hdr_info);
  //for(var x = 0; x < orders2.length; x++){
  appendRowstoSheet("14VzdU_0cZZ_x4Ab3zTF1SCUS5mTYZEdkZxhLFae2T0w","PurchaseOrders",orders);//}
  
  //appendRowtoSheet("14VzdU_0cZZ_x4Ab3zTF1SCUS5mTYZEdkZxhLFae2T0w","POitems",itms_hdr);
  for(var k = 0; k < all_itms_array.length; k++){
    //appendRowstoSheet("14VzdU_0cZZ_x4Ab3zTF1SCUS5mTYZEdkZxhLFae2T0w","POitems",all_itms_array[k]);
  }
  /* 
  var spreadsheet = SpreadsheetApp.openById("14VzdU_0cZZ_x4Ab3zTF1SCUS5mTYZEdkZxhLFae2T0w");
  var a1sheet = spreadsheet.getSheetByName("POitems");
  var sheet = spreadsheet.setActiveSheet(a1sheet);
  
  // set the formula array to be duplicated as many times as there are rows
  var formulas = '=if(iserror(search("2 Pack",R[0]C[-4])),R[0]C[-3],R[0]C[-3]*2)';
  var last_row = sheet.getLastRow();
  var form_range = "I2:" + "I" + last_row;
  var cell = sheet.getRange(form_range);
  cell.setFormulaR1C1(formulas);
  removeDuplicates(sheet);
  */
  //var a2sheet = spreadsheet.getSheetByName("PurchaseOrders");
  //var sheet2 = spreadsheet.setActiveSheet(a2sheet);
  //removeDuplicates(sheet2);
};





function lastValue(column, sheet) {
  var lastRow = sheet.getMaxRows();
  var values = sheet.getRange(column + "1:" + column + lastRow).getValues();
  
  for (; values[lastRow - 1] == "" && lastRow > 0; lastRow--) {}
  return lastRow;
};



function getOrderSummaryFromMessageNewBless(emailmsg){
  
  var method = 3;
  
  var ordr = emailmsg.getBody();
  var emlbody = emailmsg.getPlainBody();
  Logger.log(emlbody);
  var order_dt = emailmsg.getDate();
  var ThreadNm = emailmsg.getSubject();
  var recipient = emailmsg.getTo();
  if(recipient.search("<") != -1){recipient = recipient.substring(recipient.lastIndexOf("<") + 1 , recipient.lastIndexOf(">"));}
  var order_summary = [];
  var shpname = "";
  var cust_nm = "";
  var addr1 = "";
  var addr2 = "";
  var city = "";
  var state = "";
  var zip = "";
  var cityst = "";
  var billname = "";
  var billaddr1 = "";
  var billaddr2 = "";
  var billcity = "";
  var billcityst = "";
  var billstate = "";
  var billzip = "";
  var tmp;
  //New Email Style as of 5.15.2018
  if(method == 3){
    var ty_indx = emlbody.search("= Thank You! =");
    var cus_indx = emlbody.search("== Bill To\: ==");
    var sh_indx = emlbody.search("== Ship To\: ==");
    var ordrsmm = emlbody.search("== Order Summary ==");
    
    var dt_indx = emlbody.search("Date:");
    var ordrnm = emlbody.search("Order #:");
    var ppup_indx = emlbody.search("Pop-Up #:");
    var ordrttl_indx = emlbody.search("Order Total:");
    var notes_indx = emlbody.search("Notes");
    var pymt_indx = emlbody.search("Payment");
    var ordrhdr = emlbody.substring(dt_indx,ordrsmm);
    tmp = ordrhdr.match(/Order #:\s*([0-9]*)/);
    var Ordr_Nmbr = (tmp && tmp[1]) ? tmp[1].trim() : '';
    tmp = ordrhdr.match(/Pop-Up #:\s*([0-9]*)/);
    var popup_nb = (tmp && tmp[1]) ? tmp[1].trim() : '';
    
    
    
    
    if(sh_indx > 0 ){ 
      var cust_tbl2 = emlbody.substring(cus_indx,sh_indx);
      var shp_tbl2 = emlbody.substring(sh_indx,dt_indx);
    }
    else{
      var cust_tbl2 = emlbody.substring(cus_indx,dt_indx);
      var shp_tbl2 = emlbody.substring(cus_indx,dt_indx);
    }
    
    var ordr_tbl2 = emlbody.substring(dt_indx,notes_indx);
    var ordrsumm_tbl2 = emlbody.substring(ordrsmm,pymt_indx);
    var cogs_cnt = OrderDeets(ordrsumm_tbl2);
    
    
    var pymt_indx = emlbody.search("Payment");
    var subttl_plain = emlbody.search("Subtotal");
    var esttax_plain = emlbody.search("Tax");
    var ShpHnd_indx = emlbody.search("S&H"); 
    var ShpHndTax_plain = emlbody.search("S&H Tax");
    var x = 8;
    if(ShpHndTax_plain==-1){
      ShpHndTax_plain = emlbody.search("S&H\r\nTax");
      x=9;}
    
    
    var emlbody2 = emlbody.substring(subttl_plain, emlbody.length);
    
    shpname = cust_nm;
    tmp = ordrsumm_tbl2.match(/Items Subtotal\s*\$(\-*[0-9]*\.[0-9]*)/);
    var amt_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0;
    tmp = ordrsumm_tbl2.match(/Order Discount\s*\$(\-*[0-9]*\.[0-9]*)/);
    var disc_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0;
    tmp = ordrsumm_tbl2.match(/[^Items]\s*Subtotal\s*\$(\-*[0-9]*\.[0-9]*)/);
    var amt_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0;
    tmp = ordrsumm_tbl2.match(/  Tax\s*\$(\-*[0-9]*\.[0-9]*)/);
    var tax_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0;
    tmp = ordrsumm_tbl2.match(/S&H\s*\$(\-*[0-9]*\.[0-9]*)/);
    var sh_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0; 
    tmp = ordrsumm_tbl2.match(/S&H\s*Tax\s*\$(\-*[0-9]*\.[0-9]*)/);
    var shtax_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0; 
    tmp = ordrsumm_tbl2.match(/Total\s*\$(\-*[0-9]*\.[0-9]*)/);
    var tp_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0;
    
    var shp45_new = [];
    var shp46_new = [];
    var nmmmmm;
    tmp = emlbody2.match(/Subtotal[\s]*\$(\-*[\0-9\s]+)(Tax)/);
    if(tmp == null){esttax_plain = -1;}
    if(sh_indx>0 ){ 
      var shp45 = cust_tbl2.split("\n");
      var shp46 = shp_tbl2.split("\n");
      for(var shp_int = 0; shp_int < shp45.length; shp_int++)
      {
        shp45[shp_int]=shp45[shp_int].trim();
        if(shp45[shp_int] != ""){shp45_new.push(shp45[shp_int]);}
      }
      for(var shp_int = 0; shp_int < shp46.length; shp_int++)
      {
        shp46[shp_int]=shp46[shp_int].trim();
        if(shp46[shp_int] != ""){shp46_new.push(shp46[shp_int]);}
      }
      
      
      tmp = shp46_new[shp46_new.length-1].match(/^([A-z\s\-\.]*)\s*\,\s([A-Z]{2})\s([0-9]{5})/);
      var city = (tmp && tmp[1]) ? tmp[1].trim() : '';
      var state = (tmp && tmp[2]) ? tmp[2].trim() : '';
      var zip = (tmp && tmp[3]) ? tmp[3].trim() : '';
      
      cust_nm = shp45_new[1];
      
      
      shpname = shp46_new[1];
      addr1 = shp46_new[2];
      if (shp46_new.length==5)
      { 
        addr2 = shp46_new[3];
      }
      if (shp46_new.length==3)
      { 
        shpname = cust_nm;
        addr1 = shp46_new[1];
      }
    }
    
    
  }
  
  
  
  
  
  var disc_nb = "0";
  var clothestax = tax_nb;
  var shiptax = shtax_nb;
  
  if(clothestax==null){
    clothestax=0;}
  if(shiptax==null){
    shiptax=0;}  
  
  var tot_tax1 = parseFloat(clothestax);
  var tot_tax2 = parseFloat(shiptax);
  var tot_tax = tot_tax1 + tot_tax2;
  
  
  
  
  
  var PrintStatus = '=iferror(vlookup(R[0]C[-17],Shipments!C[-18]:C[-16],1,false),"Not Printed")';
  var ActualShipping = '=iferror(vlookup(R[0]C[-18],Shipments!C[-19]:C[-12],6,false),0)';
  var TotalProfit = '=R[0]C[-6]-R[0]C[-7]-R[0]C[-4]-R[0]C[-1]';
  
  
  
  var order_summary = [];
  order_summary[0] = cust_nm;
  order_summary[1] = Ordr_Nmbr;
  order_summary[2] = popup_nb;
  order_summary[3] = shpname;
  order_summary[4] = addr1;
  order_summary[5] = addr2;
  order_summary[6] = city;
  order_summary[7] = state; 
  order_summary[8] = zip;
  order_summary[9] = order_dt;
  order_summary[10] = amt_nb;
  order_summary[11] = disc_nb;
  order_summary[12] = sh_nb;
  order_summary[13] = tot_tax;
  order_summary[14] = tp_nb;
  order_summary[15] = recipient;
  order_summary[16] = cogs_cnt[0];
  order_summary[17] = cogs_cnt[1];
  order_summary[18] = PrintStatus;
  order_summary[19] = ActualShipping;
  order_summary[20] = TotalProfit;
  
  
  Logger.log(order_summary);
  
  
  return order_summary;  
  
};
