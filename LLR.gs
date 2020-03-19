function LogOrders()
{
  var invoicesheetID = initialSetup();
  var returnsOut = GmailApp.search("\"Lularoe Kristen Hickman\" is:unread from:(noreply@lularoebless.com) subject:(Return Receipt from LuLaRoe - Order Number)");
  Logger.log("Returns:" + returnsOut.length);
  if(returnsOut.length > 0){
    //Logger.log("Returns:" + returnsOut.length);
    getReturns(invoicesheetID,returnsOut);
  }
  var transfersOut = GmailApp.search("-to:( lularoekristenhickman@gmail.com) is:unread from:(noreply@lularoebless.com) subject:(Transfer Receipt from LuLaRoe - Order Number)");
  Logger.log("Transfer Out:" + transfersOut.length);
  if(transfersOut.length > 0){
    //Logger.log("Transfer Out:" + transfersOut.length);
    getTransfersEmail(invoicesheetID,transfersOut,0);
  }
  var transfersIn = GmailApp.search("to:( lularoekristenhickman@gmail.com) is:unread from:(noreply@lularoebless.com) subject:(Transfer Receipt from LuLaRoe - Order Number)");
  Logger.log("Transfer In:" + transfersIn.length);
  if(transfersIn.length > 0){
    //Logger.log("Transfer In:" + transfersIn.length);
    getTransfersEmail(invoicesheetID,transfersIn,1);
  }
  var unreads = GmailApp.search("is:unread from:(noreply@lularoebless.com) subject:(Purchase Receipt from LuLaRoe - Order Number) \"Lularoe Kristen Hickman\"");
  //var unreads = GmailApp.search("is:unread from:(noreply@lularoebless.com) subject:(Purchase Receipt from LuLaRoe - Order Number) \"KristenHickman\""); This changed at Noon on 2/20/2019
  Logger.log("Sales:" + unreads.length);
  if(unreads.length > 0){
    //Logger.log("Sales:" + unreads.length);
    LogNewBless(invoicesheetID,unreads);
  }
  var RetailtransfersIn = GmailApp.search("is:unread from:(noreply@lularoebless.com) subject:(Purchase Receipt from LuLaRoe - Order Number) -\"KristenHickman\"");
  Logger.log("Retail Transfer In:" + RetailtransfersIn.length);
  if(RetailtransfersIn.length > 0){
    //Logger.log("Transfer In:" + transfersIn.length);
    getTransfersEmail(invoicesheetID,RetailtransfersIn,1);
  }
  var PurchOrders= GmailApp.search("is:unread from:(noreply@lularoe.com) subject:(LuLaRoe Wholesale Order Confirmation)");
  //var PurchOrders= GmailApp.search("is:unread from:(noreply@lularoe.com) subject:(LuLaRoe Web Order Confirmation)"); //Change on 6/25/19
  Logger.log("Purchase Orders:" + PurchOrders.length);
  if(PurchOrders.length > 0){
    //Logger.log("Purchase Orders:" + PurchOrders.length);
    LLRPurchaseOrder(PurchOrders);
  }
  if(returnsOut.length + transfersOut.length + unreads.length > 0){
    dotheworksheet(invoicesheetID);
  }
  
}

function oldBlessInvoices()
{
  var OldInvoices= GmailApp.search("from:(no-reply@mylularoe.com) subject:(Purchase receipt from LuLaRoe)");
  Logger.log("Purchase Orders:" + OldInvoices.length);
  for (var a=0;a < 5; a++)
  {
    var message = OldInvoices[a].getMessages();
    var ThreadNm = OldInvoices[a].getFirstMessageSubject();
    var emlbody = OldInvoices[a].getMessages()[0].getPlainBody();
    Logger.log(emlbody);
  }
}




function OrderDeets(emlbody){
  
  //if(unreads.length > 0){
  //for (var a=0;a < unreads.length; a++)
  //{
  //var message = unreads[a].getMessages();
  //var ThreadNm = unreads[a].getFirstMessageSubject()
  //var recipient = unreads[a].getMessages()[0].getTo();
  //var emlbody = unreads[a].getMessages()[0].getPlainBody();
  //Logger.log(emlbody);
  var tmp;
  tmp = emlbody.match(/O\/S\s/g);
  var leggingsOS = tmp ? tmp.length : 0;
  tmp = emlbody.match(/T\/C\s/g) ;
  var leggingsTC = tmp ? tmp.length : 0;
  tmp = emlbody.match(/T\/C\s2\s/g)  ;
  var leggingsTC2 = tmp ? tmp.length : 0;
  tmp = emlbody.match(/S\/M\s/g);
  var leggingsSM = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Tween\s/g);
  var leggingsTW = tmp ? tmp.length : 0;
  tmp = emlbody.match(/L\/XL\s/g);
  var leggingsXL = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Noir\s*Leggings\s*.*\s*Various/g);
  var blklgsTC2 = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Leggings\s*O\/S/g);
  var blklgsOS = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Leggings\s*T\/C\s\s/g);
  var blklgsTC = tmp ? tmp.length : 0;
  
  
  
  
  
  tmp = emlbody.match(/Lynnae\s/g) ;
  var lynnae = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Cassie\s/g);
  var cassie = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Sarah\s/g)  ;
  var sarah = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Carly\s/g) ; 
  var carly = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Classic\s*T\s/g)  ;
  var classic = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Disney\s*Leggings\s*.*\s*L\/XL/g) ;
  var DisleggingsLXL = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Randy\s/g)  ;
  var randy = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Joy\s/g) ;
  var joy = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Perfect\s*T\s/g) ;
  var perfect = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Perfect\s*Tank/g) ;
  var ptank = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Madison\s/g) ;
  var madison = tmp ? tmp.length : 0;
  
  tmp = emlbody.match(/Julia\s/g) ;
  var julia = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Lindsay\s/g) ;
  var lindsay = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Maxi\s/g)  ;
  var maxi = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Disney\s*Randy\s*.*\s*Tops/g) ;
  var Disrandy = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Irma\s/g)  ;
  var irma = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Disney\s*Leggings\s*.*\s*T\/C/g)  ;
  var DisleggingsTC = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Disney\s*Leggings\s*.*\s*O\/S/g)  ;
  var DisleggingsOS = tmp ? tmp.length : 0;
  
  tmp = emlbody.match(/Disney\s*Leggings\s*.*\s*S\/M/g) ;
  var DisleggingsSM =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Amelia\s/g) ;
  var amelia =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Disney\s*Leggings\s*.*\s*Tween/g) ;
  var DisleggingsTween =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Azure\s/g) ;
  var azure =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Monroe\s/g) ;
  var monroe =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Nicole\s/g) ;
  var nicole =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Lola\s/g) ;
  var lola =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Shirley\s/g) ;
  var shirley =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Mimi\s/g) ;
  var mimi =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Disney\s*Irma\s*.*\s*Tops/g) ;
  var Disirma =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Disney\s*Classic\s*T\s*.*\s*Tops/g) ;
  var Disclassic =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Elegant\s*2017\s*Debbie\s*.*\s*Dresses/g);
  var elgDebbie = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Elegant\s*2017\s*Shirley\s*.*\s*Tops/g);
  var elgShirley = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Elegant\s*2017\s*Carly\s*.*\s*Dresses/g);
  var elgCarly = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Elegant\s*2017\s*DeAnne\s*.*\s*Dresses/g);
  var elgDeanne = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Elegant\s*2017\s*Cassie\s*.*\s*Skirts/g);
  var elgCassie = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Elegant\s*2017\s*Sarah/g);
  var elgSarah = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Jaxon\s/g);
  var jaxon = tmp ? tmp.length : 0;
  tmp = emlbody.match(/Harvey\s/g);
  var harvey = tmp ? tmp.length : 0;
  
  
  tmp = emlbody.match(/Valentine's\s*2018\s*Leggings\s*Leggings\s*T\/C/g) ;
  var vdayleggingsTC =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Valentine's\s*2018\s*Leggings\s*Leggings\s*O\/S/g) ;
  var vdayleggingsOS =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Valentine's\s*2018\s*Leggings\s*Leggings\s*Various/g) ;
  var vdayleggingsTC2 =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Valentine's\s*2018\s*Leggings\s*Leggings\s*S\/M/g) ;
  var vdayleggingsSM =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Valentine's\s*2018\s*Leggings\s*Leggings\s*L\/XL/g) ;
  var vdayleggingsXL =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Valentine's\s*2018\s*Leggings\s*Leggings\s*Tween/g) ;
  var vdayleggingsTW =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Maria\s/g) ;
  var maria =  tmp ? tmp.length : 0;
   tmp = emlbody.match(/Maurine\s/g) ;
  var maurine =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Debbie\s/g) ;
  var debbie =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Jane\s/g) ;
  var jane =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Cici/g) ;
  var cici =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Tank\s*Top/g) ;
  var tank =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Christy\s*T/g) ;
  var christy =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Dani/g) ;
  var dani =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Georgia\s/g) ;
  var grga =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Caroline\s/g) ;
  var carol =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Jessie\s/g) ;
  var jessie =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Marly\s/g) ;
  var marly =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Mae\s/g) ;
  var mae =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Adeline\s/g) ;
  var adeline =  tmp ? tmp.length : 0;
    tmp = emlbody.match(/Scarlett\s/g) ;
  var scarlett =  tmp ? tmp.length : 0;
    tmp = emlbody.match(/Sariah\s/g) ;
  var sariah =  tmp ? tmp.length : 0;
tmp = emlbody.match(/Gracie\s/g) ;
  var gracie =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Nicki\s/g) ;
  var nikki =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Amber\s/g) ;
  var amber =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Gigi\s/g) ;
  var gigi =  tmp ? tmp.length : 0;
  
   tmp = emlbody.match(/Valentina\s/g) ;
  var valentina =  tmp ? tmp.length : 0;
    tmp = emlbody.match(/Hudson\s/g) ;
  var hudson =  tmp ? tmp.length : 0;
    tmp = emlbody.match(/Lucille\s/g) ;
  var lucille =  tmp ? tmp.length : 0;
   tmp = emlbody.match(/Denim\s/g) ;
  var denim =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Mitzi\s/g) ;
  var mitzi =  tmp ? tmp.length : 0;
   tmp = emlbody.match(/Mimi\s/g) ;
  var mimi =  tmp ? tmp.length : 0;
     tmp = emlbody.match(/Ellie\s/g) ;
  var ellie =  tmp ? tmp.length : 0;
       tmp = emlbody.match(/Liv\s/g) ;
  var liv =  tmp ? tmp.length : 0;
       tmp = emlbody.match(/Michelle\s/g) ;
  var michelle =  tmp ? tmp.length : 0;  
  tmp = emlbody.match(/Jax\s/g) ;
  var jax =  tmp ? tmp.length : 0;
    tmp = emlbody.match(/Emily\s/g) ;
  var emily =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Iris\s/g) ;
  var iris =  tmp ? tmp.length : 0;
   tmp = emlbody.match(/Hudson Long Sleeve\s/g) ;
  var hudsonLS =  tmp ? tmp.length : 0;
  
  //25,25,39,70,20,42
 //elgDebbie,elgShirley,elgCarly,elgDeanne,elgCassie,elgSarah
  //var DisleggingsTC = emlbody.match(/Mimi Tops/g);
  //var leggingsOS = emlbody.match(/Disney Leggings Leggings TC/g);
  var OrdrPcs = [leggingsOS+blklgsOS+vdayleggingsOS, lynnae,cassie-elgCassie, sarah-elgSarah,carly-elgCarly,classic-Disclassic,DisleggingsLXL,randy-Disrandy,joy,perfect+ptank,madison,leggingsTC+blklgsTC+vdayleggingsTC,julia,lindsay,maxi,Disrandy,irma-Disirma,DisleggingsTC,DisleggingsOS,leggingsTC2+blklgsTC2+vdayleggingsTC2,DisleggingsSM,amelia,DisleggingsTween,azure,monroe, nicole,lola,shirley-elgShirley,mimi,Disirma, Disclassic,leggingsSM+leggingsXL+vdayleggingsSM+vdayleggingsXL,leggingsTW+vdayleggingsTW,elgDebbie,elgShirley,elgCarly,elgDeanne,elgCassie,elgSarah,jaxon+harvey,maria,maurine,debbie,jane,cici,tank,christy,dani,grga,carol,jessie,marly,mae+adeline+sariah+scarlett+gracie,nikki,amber,gigi,valentina,hudson,lucille,denim,mitzi,mimi,ellie,liv,michelle,jax,emily,iris,hudsonLS];
  //Logger.log(OrdrPcs);
  var WSprices = [10.5,18,14,30,25,16,10,16,25,17,23,10.5,18,21,21,21,15,12.5,12.5,12,10,31,11,14,10.5,23,21,22,34,22,20,8.5,9.5,25,25,39,70,20,42,50,34,36,24,24,35,12.5,17,27.5,32.5,27,25,25,5,23,21,15,22,17,30,35,17,30,26,17,24,26,25,16,19];
  var OrdrCOGS=0;
  var OrdrPcsCnt=0;
  for(var b=0;b<OrdrPcs.length;b++){
    OrdrCOGS=OrdrPcs[b]*WSprices[b]+OrdrCOGS;
    OrdrPcsCnt=OrdrPcs[b]+OrdrPcsCnt;
  }
  var cogs_cnt = [OrdrCOGS,OrdrPcsCnt];   
  Logger.log("COGS:"+ OrdrCOGS);
  Logger.log("Pieces:" + OrdrPcsCnt);
  return cogs_cnt;
  //}
  // }
  
}
/*
function parserofthetable(){
  Document doc = Jsoup.parse(html);
  Elements tables = doc.select("table");
  for (Element table : tables) {
    Elements trs = table.select("tr");
    String[][] trtd = new String[trs.size()][];
                                             for (int i = 0; i < trs.size(); i++) {
                                             Elements tds = trs.get(i).select("td");
                                             trtd[i] = new String[tds.size()];
    for (int j = 0; j < tds.size(); j++) {
      trtd[i][j] = tds.get(j).text(); 
    }
  }
  // trtd now contains the desired array for this table
}
}*/

function importCSVFromGoogleDrive(filename, destWrkbk, destShtnm) {
  
  var file = DriveApp.getFilesByName(filename).next();
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
  var ss = SpreadsheetApp.openById(destWrkbk);
  var centralStore = ss.getSheetByName (destShtnm);
  centralStore.clear();
  
  centralStore.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  ss.setActiveSheet(ss.getSheetByName(destShtnm));
  
}



function getReturns(invoicesheetID, returnsOut){
  //var returnsOut = GmailApp.search("-to:( lularoekristenhickman@gmail.com) from:(noreply@lularoebless.com) subject:(Return Receipt from LuLaRoe - Order Number)");
  //Logger.log(returnsOut);
  if(returnsOut.length > 0){
    var all_added_orders = [];
    for (var a=0;a < returnsOut.length; a++)
    {
      var message = returnsOut[a].getMessages();
      var ThreadNm = returnsOut[a].getFirstMessageSubject()
      var Ordr_Nmbr = ThreadNm.substring(ThreadNm.lastIndexOf("Number ")+7,ThreadNm.length);
      //Logger.log(Ordr_Nmbr);
      var recipient = returnsOut[a].getMessages()[0].getTo();
      recipient = recipient.substring(recipient.lastIndexOf("<") + 1 , recipient.lastIndexOf(">"));
      
      
      var order_summary = getOrderSummaryFromMessageNewBless(message[0]);
      all_added_orders.push(order_summary);
      
      returnsOut[a].markRead();
    } 
    
    if(all_added_orders.length == 1) {
      appendRowtoSheet(invoicesheetID, "Returns", all_added_orders[0]);
    }
    else{
      appendRowstoSheet(invoicesheetID, "Returns", all_added_orders);}
  }
}



function getTransfersEmail(invoicesheetID,transfers,transferInBool){
  //var transfersOut = GmailApp.search("-to:( lularoekristenhickman@gmail.com) is:unread from:(noreply@lularoebless.com) subject:(Transfer Receipt from LuLaRoe - Order Number)");
  //var transfersIn = GmailApp.search("to:( lularoekristenhickman@gmail.com) from:(noreply@lularoebless.com) subject:(Transfer Receipt from LuLaRoe - Order Number)");
  //var transfersIn = GmailApp.search("to:( lularoekristenhickman@gmail.com) is:unread from:(noreply@lularoebless.com) subject:(Transfer Receipt from LuLaRoe - Order Number)");
  //var invoicesheetID = initialSetup();
  //getReturns(invoicesheetID);
  
  if(transferInBool==0){ //TransferOut
    var all_added_orders = [];
    for (var a=0;a < transfers.length; a++)
    {
      var message = transfers[a].getMessages();
      //Logger.log(message.getPlainBody());
      var ThreadNm = transfers[a].getFirstMessageSubject()
      var Ordr_Nmbr = ThreadNm.substring(ThreadNm.lastIndexOf("Number ")+7,ThreadNm.length);
      //Logger.log(Ordr_Nmbr);
      var recipient = transfers[a].getMessages()[0].getTo();
      recipient = recipient.substring(recipient.lastIndexOf("<") + 1 , recipient.lastIndexOf(">"));
      
      
      var order_summary = getOrderSummaryFromMessageNewBless(message[0]);
      all_added_orders.push(order_summary);
      
      transfers[a].markRead();
    } 
    
    if(all_added_orders.length == 1) {
      appendRowtoSheet(invoicesheetID, "TransferOut", all_added_orders[0]);
    }
    else{
      appendRowstoSheet(invoicesheetID, "TransferOut", all_added_orders);}
  }
  
  if(transferInBool==1){ //TransferIn
    //put log of purchase transfers in here
    var all_added_transfers = [];
    for (var b=0;b < transfers.length; b++)
    {
      var message = transfers[b].getMessages();
      var ThreadNm = transfers[b].getFirstMessageSubject();
      //var emlbody = transfers[b].getMessages()[0].getPlainBody();
      var Ordr_Nmbr = ThreadNm.substring(ThreadNm.lastIndexOf("Number ")+7,ThreadNm.length);
      var order_summary = getOrderSummaryFromMessageNewBless(message[0]);
      all_added_transfers.push(order_summary);
      transfers[b].markRead();
      
    }
    var transfersToSheet = [];
    for( var c=0;c < all_added_transfers.length; c++){
      var transfersarr = [];
      transfersarr[0]= all_added_transfers[c][1];
      transfersarr[1]= "Perks";
      transfersarr[2]= "USPS";
      transfersarr[3]= "115561";
      transfersarr[4]= "Kristen Hickman";
      transfersarr[5]= all_added_transfers[c][9];
      transfersarr[6]= "lularoekristenhickman@gmail.com";
      transfersarr[7]= "7248753155";
      transfersarr[8]= "Consultant Transfer";
      transfersarr[9]= "Wholesale";
      transfersarr[10]= all_added_transfers[c][10];
      transfersarr[11]= all_added_transfers[c][12];
      transfersarr[12]= all_added_transfers[c][13];
      transfersarr[13]= all_added_transfers[c][14];
      transfersarr[14]= all_added_transfers[c][16];
      transfersarr[15]= all_added_transfers[c][17];
      transfersToSheet.push(transfersarr);
    }
    
    if(all_added_transfers.length == 1) {
      appendRowtoSheet(invoicesheetID, "TransferIn", transfersToSheet[0]);}
    
    else{
      appendRowstoSheet(invoicesheetID, "TransferIn", transfersToSheet);}               
  }
}




function LogNewBless(invoicesheetID,unreads){
  //var invoicesheetID = initialSetup();
  //getTransfersEmail(invoicesheetID);
  //var unreads1 = GmailApp.search("-to:( lularoekristenhickman@gmail.com) is:unread from:(noreply@lularoebless.com) subject:(Purchase Receipt from LuLaRoe - Order Number)");
  //var unreads= GmailApp.search("is:unread from:(noreply@lularoebless.com) subject:(Purchase Receipt from LuLaRoe - Order Number)");
  
  //Logger.log(unreads.length);
  

  importData(invoicesheetID);
  var ordersinSheet = getWorkbookColumnData(invoicesheetID, "Invoices", 2);
  //var emailsinSheet = getWorkbookColumnData(invoicesheetID, "Emails", 1);
  var all_added_orders = [];
  var all_added_order_items = [];
  if(ordersinSheet.length == 0 || ordersinSheet.length == ""){
    var spreadsheet = SpreadsheetApp.openById(invoicesheetID);
    var a1sheet = spreadsheet.getSheetByName("Invoices");
    var sheet = spreadsheet.setActiveSheet(a1sheet);
    sheet.appendRow(["Invoice Name","Order#","Popup#","Shipping Name","Address1", "Address2","City","State","Zip","Invoice Date/Time","Subtotal","Discounts","Shipping","Tax","Total Sale","Email","Status","Total COGS","Actual Shipping","Items Invoiced"]);
  }
  
    for (var a=0;a < unreads.length; a++)
    //for (var a=0;a < 100; a++) // comment out previous and uncomment this one to re-run all
    {
      var message = unreads[a].getMessages();  
      var dup_ordr = 0;
      var dup_email = 0;
      for( var v = 0; v < ordersinSheet.length; v++)
      { if(ordersinSheet[v].toString() == (Ordr_Nmbr)){dup_ordr=1;break}} 
      if(dup_ordr == 0)
      {
        var order_summary = getOrderSummaryFromMessageNewBless(message[0]);
        all_added_orders.push(order_summary);
      }//end if dup_ordr
      unreads[a].markRead();
    } // ending a for loop
     
    if(all_added_orders.length == 1) {appendRowtoSheet(invoicesheetID, "Invoices", all_added_orders[0]);}
    else{appendRowstoSheet(invoicesheetID, "Invoices", all_added_orders);}
    
  
} // ending function






function dotheworksheet(invoicesheetID){
  // set up Invoices (order summary) formulas
  var PrintStatus = '=iferror(vlookup(R[0]C[-17],Shipments!C[-18]:C[-16],1,false),"Not Printed")';
  //var email = '=iferror(vlookup(R[0]C[-15],Emails!C[-16]:C[-15],2,false),"")';
  //var TotalCOGS = '=sumifs(\'MyOrders - New Bless\'!C[-12]:C[-12],\'MyOrders - New Bless\'!C[-17]:C[-17],R[0]C[-16])';
  //var itemsInvcd = '=sumifs(\'MyOrders - New Bless\'!C[-17]:C[-17],\'MyOrders - New Bless\'!C[-19]:C[-19],R[0]C[-18])';
  var ActualShipping = '=iferror(vlookup(R[0]C[-18],Shipments!C[-19]:C[-12],6,false),0)';
  //var ChargedShipping = '=iferror(ARRAYFORMULA(INDEX(\'Order Details\'!C[-17]:C[-9],MATCH(1,(R[0]C[-16]=\'Order Details\'!C[-17]:C[-17])*(\'Order Details\'!C[-14]:C[-14]="na"),0),3)),"Shipping and Handling")';
  var TotalProfit = '=R[0]C[-6]-R[0]C[-7]-R[0]C[-4]-R[0]C[-1]';
  var formulas = [];
  
  //formulas = [PrintStatus, TotalCOGS, ActualShipping, itemsInvcd,TotalProfit];
  formulas = [PrintStatus, ActualShipping, TotalProfit];
 
  
  var sheetnamestodo = ["Invoices","Returns","TransferOut","TransferIn"]
  var spreadsheet = SpreadsheetApp.openById(invoicesheetID);
  for (var sn=0;sn<sheetnamestodo.length;sn++)
  {
    var formulas1 = [];
    var a1sheet = spreadsheet.getSheetByName(sheetnamestodo[sn]);
    var sheet = spreadsheet.setActiveSheet(a1sheet);
    //var rows1 = lastValue("N",sheet);
    
    // set the formula array to be duplicated as many times as there are rows
    var last_row = sheet.getLastRow();
    if(last_row == 1) {
      last_row = 2;
    }
    for (var p =0;p<last_row-1; p++){
      formulas1[p] = formulas;
    }
    //rows1 = rows1+1;
    var form_range = "S2" + ":" + "U" + last_row;
    //var form_range = "N" + rows1 + ":" + "Q" + last_row;
    var cell = sheet.getRange(form_range);
    cell.setFormulasR1C1(formulas1);
    
    
    //cell.copyTo(sheet.getRange(form_range), {contentsOnly: true});
    
    
    //}//row 101-97
    
    removeDuplicates(sheet);
    //sort the order summary by paid invoice date
    form_range = "A2:" + "V" + last_row;
    cell = sheet.getRange(form_range);
    cell.sort({column: 10, ascending: false});
    
  }
  //export the sheet to csv for import into Stamps.com
  //export the sheet to xlsx for easy manipulation
  exportSheet(invoicesheetID);
  
}







function LogNewBlessCatchAll(){
  var invoicesheetID = initialSetup();
  
  
  getTransfersEmail(invoicesheetID);
  //var threads = GmailApp.getInboxThreads(0, 1);
  //var msssssg = threads[0].getMessages();
  //var dddate = msssssg[0].getDate();
  //var formattedDate = Utilities.formatDate(dddate, "EST", "yyyy-MM-dd");
  // dddate = dddate;
  
  
  var unreads1 = GmailApp.search("-to:( lularoekristenhickman@gmail.com) is:unread from:(noreply@lularoebless.com) subject:(Purchase Receipt from LuLaRoe - Order Number)");
  var unreads= GmailApp.search("newer_than:3d from:(noreply@lularoebless.com) subject:(Purchase Receipt from LuLaRoe - Order Number)");
  
  Logger.log(unreads.length);
  
  if(unreads.length > 0){
    importData(invoicesheetID);}
  var ordersinSheet = getWorkbookColumnData(invoicesheetID, "Invoices", 2);
  //var emailsinSheet = getWorkbookColumnData(invoicesheetID, "Emails", 1);
  var all_added_orders = [];
  var all_added_order_items = [];
  if(ordersinSheet.length == 0 || ordersinSheet.length == ""){
    var spreadsheet = SpreadsheetApp.openById(invoicesheetID);
    var a1sheet = spreadsheet.getSheetByName("Invoices");
    var sheet = spreadsheet.setActiveSheet(a1sheet);
    var a2sheet = spreadsheet.getSheetByName("Order Details");
    var sheet2 = spreadsheet.setActiveSheet(a2sheet);
    sheet.appendRow(["Invoice Name","Order#","Popup#","Shipping Name","Address1", "Address2","City","State","Zip","Invoice Date/Time","Subtotal","Discounts","Shipping","Tax","Total Sale","Email","Status","Total COGS","Actual Shipping","Items Invoiced"]);
    sheet2.appendRow(["Order#","Item Qty","Item Description","Item Size","MAP","Wholesale Cost","Profit","Cost Category"]);
  }
  if(unreads.length > 0){
    for (var a=0;a < unreads.length; a++)
    {
      
      
      var message = unreads[a].getMessages();
      var ThreadNm = unreads[a].getFirstMessageSubject()
      var recipient = unreads[a].getMessages()[0].getTo();
      recipient = recipient.substring(recipient.lastIndexOf("<") + 1 , recipient.lastIndexOf(">"));
      var Ordr_Nmbr = ThreadNm.substring(ThreadNm.lastIndexOf("Number ")+7,ThreadNm.length);
      Logger.log(Ordr_Nmbr);
      
      var dup_ordr = 0;
      var dup_email = 0;
      
      for( var v = 0; v < ordersinSheet.length; v++)
      {
        if(ordersinSheet[v].toString() == (Ordr_Nmbr))
        {
          dup_ordr=1;
        }
      }
      
      if(dup_ordr == 0)
      {
        //var order_itms = [];
        var order_summary = getOrderSummaryFromMessageNewBless(message[0]);
        all_added_orders.push(order_summary);
        //var order_itms = getOrderDetailsFromMessageNewBless(message[a]);
        //for (var f = 0;f < order_itms.length; f++){
        // all_added_order_items.push(order_itms[f]);}
        
      }//end if dup_ordr
      
      unreads[a].markRead();
    } // ending a for loop
    
    
    if(all_added_orders.length == 1) {
      appendRowtoSheet(invoicesheetID, "Invoices", all_added_orders[0]);
    }
    else{
      appendRowstoSheet(invoicesheetID, "Invoices", all_added_orders);}
    //}// end if 0 unreads
    
    // set up Invoices (order summary) formulas
    var PrintStatus = '=iferror(vlookup(R[0]C[-15],Shipments!C[-16]:C[-14],1,false),"Not Printed")';
    var email = '=iferror(vlookup(R[0]C[-15],Emails!C[-16]:C[-15],2,false),"")';
    var TotalCOGS = '=sumifs(\'MyOrders - New Bless\'!C[-12]:C[-12],\'MyOrders - New Bless\'!C[-17]:C[-17],R[0]C[-16])';
    var itemsInvcd = '=sumifs(\'MyOrders - New Bless\'!C[-17]:C[-17],\'MyOrders - New Bless\'!C[-19]:C[-19],R[0]C[-18])';
    var ActualShipping = '=iferror(vlookup(R[0]C[-17],Shipments!C[-18]:C[-11],6,false),0)';
    var ChargedShipping = '=iferror(ARRAYFORMULA(INDEX(\'Order Details\'!C[-17]:C[-9],MATCH(1,(R[0]C[-16]=\'Order Details\'!C[-17]:C[-17])*(\'Order Details\'!C[-14]:C[-14]="na"),0),3)),"Shipping and Handling")';
    var TotalProfit = '=R[0]C[-6]-R[0]C[-7]-R[0]C[-3]-R[0]C[-2]';
    var formulas = [];
    formulas = [PrintStatus, TotalCOGS, ActualShipping, itemsInvcd,TotalProfit];
    var formulas1 = [];
    
    
    var spreadsheet = SpreadsheetApp.openById(invoicesheetID);
    var a1sheet = spreadsheet.getSheetByName("Invoices");
    var sheet = spreadsheet.setActiveSheet(a1sheet);
    var rows1 = lastValue("N",sheet);
    
    // set the formula array to be duplicated as many times as there are rows
    var last_row = sheet.getLastRow();
    for (var p =0;p<last_row-1; p++){
      formulas1[p] = formulas;
    }
    rows1 = rows1+1;
    var form_range = "Q2" + ":" + "U" + last_row;
    //var form_range = "N" + rows1 + ":" + "Q" + last_row;
    var cell = sheet.getRange(form_range);
    cell.setFormulasR1C1(formulas1);
    
    //cell.copyTo(sheet.getRange(form_range), {contentsOnly: true});
    
    
    //}//row 101-97
    
    removeDuplicates(sheet);
    //sort the order summary by paid invoice date
    form_range = "A2:" + "U" + last_row;
    cell = sheet.getRange(form_range);
    cell.sort({column: 10, ascending: false});
    
  }    
  if(unreads.length != 0){
    //export the sheet to csv for import into Stamps.com
    //export the sheet to xlsx for easy manipulation
    exportSheet(invoicesheetID);
  }
} // ending function
















function LogIntoSheet(){
  
  var unreads = GmailApp.search("is:unread from:(no-reply@mylularoe.com) subject:(Purchase receipt from LuLaRoe)");
  if(unreads.length == 0){return;}
  var invoicesheetID = initialSetup();
  importData(invoicesheetID);
  var ordersinSheet = getWorkbookColumnData(invoicesheetID, "Invoices", 2);
  var emailsinSheet = getWorkbookColumnData(invoicesheetID, "Emails", 1);
  if(ordersinSheet.length == 0 || ordersinSheet.length == ""){
    var spreadsheet = SpreadsheetApp.openById(invoicesheetID);
    var a1sheet = spreadsheet.getSheetByName("Invoices");
    var sheet = spreadsheet.setActiveSheet(a1sheet);
    var a2sheet = spreadsheet.getSheetByName("Order Details");
    var sheet2 = spreadsheet.setActiveSheet(a2sheet);
    sheet.appendRow(["Invoice Name","Order#","Shipping Name","Address1", "Address2","City","State","Zip","Invoice Date/Time","Items Invoiced","Subtotal","Discounts","Total Sale","Status","Email","Total COGS","Actual Shipping","Charged Shipping"]);
    sheet2.appendRow(["Order#","Item Qty","Item Description","Item Size","MAP","Wholesale Cost","Profit","Cost Category"]);
  }
  
  //var label = GmailApp.getUserLabelByName("LLR/Sales");
  var thread = unreads;
  var all_added_orders = [];
  var all_added_order_items = [];
  
  for (var i = 0; i < thread.length; i++) {
    
    if(thread[i].getFirstMessageSubject() == "Purchase receipt from LuLaRoe"){
      
      var message = thread[i].getMessages();
      for (var z = 0; z < message.length; z++) {
        if (message[z].getSubject() == "Purchase receipt from LuLaRoe") {
          
          var order1 = "<h4>";
          var order2 = "</h4>";
          var ordr = message[z].getBody();
          var ordr_pos1 = ordr.search(order1);
          var ordr_pos2 = ordr.search(order2);
          
          var ordr_nmbr = ordr.substring(ordr_pos1+4,ordr_pos2);
          
          var ordr_num = ordr_nmbr.substr(ordr_nmbr.search("#")+1,8);
          
          var dup_ordr = 0;
          var dup_email = 0;
          
          for( var v = 0; v < ordersinSheet.length; v++)
          {
            if(ordersinSheet[v].toString() == (ordr_num))
            {
              dup_ordr=1;
            }
          }
          
          for( var e = 0; e < emailsinSheet.length; e++)
          {
            if(emailsinSheet[e].toString() == (ordr_num))
            {
              dup_email=1;
            }
          }
          if(dup_email == 0)
          {
            //get email and order from InvoiceReport
            
            
          }
          
          if(dup_ordr == 0)
          {
            //var order_itms = [];
            var order_summary = getOrderSummaryFromMessage(message[z]);
            all_added_orders.push(order_summary);
            var order_itms = getOrderDetailsFromMessage(message[z]);
            for (var f = 0;f < order_itms.length; f++){
              all_added_order_items.push(order_itms[f]);}
            
          }//end if dup_ordr
          
        }//ending "Purche..." if stmt
      }//ending z for loop
    }//end if statement
    thread[i].markRead();
  }//ending i for loop
  
  
  if(all_added_orders.length == 1) {
    appendRowtoSheet(invoicesheetID, "Invoices", all_added_orders[0]);
  }
  else{
    appendRowstoSheet(invoicesheetID, "Invoices", all_added_orders);}
  /*if(all_added_order_items.length == 4 && all_added_order_items.length[0] == 8) {
    appendRowtoSheet(invoicesheetID, "Order Details", all_added_order_items);
  }
  else{
    appendRowstoSheet(invoicesheetID, "Order Details", all_added_order_items);
  }*/
  
  
  // set up Invoices (order summary) formulas
  var PrintStatus = '=iferror(vlookup(R[0]C[-12],Shipments!C[-13]:C[-11],1,false),"Not Printed")';
  var email = '=iferror(vlookup(R[0]C[-13],InvoicesReport!C[-14]:C[0],12,false),"")';
  var TotalCOGS = '=sumifs(\'Order Details\'!C[-10]:C[-10],\'Order Details\'!C[-15]:C[-15],R[0]C[-14],\'Order Details\'!C[-8]:C[-8],"Cost of Goods Sold")';
  var ActualShipping = '=iferror(vlookup(R[0]C[-15],Shipments!C[-16]:C[-9],6,false),0)';
  var ChargedShipping = '=iferror(ARRAYFORMULA(INDEX(\'Order Details\'!C[-17]:C[-9],MATCH(1,(R[0]C[-16]=\'Order Details\'!C[-17]:C[-17])*(\'Order Details\'!C[-14]:C[-14]="na"),0),3)),"Shipping and Handling")';
  var TotalProfit = '=R[0]C[-6]-R[0]C[-7]-R[0]C[-3]-R[0]C[-2]';
  var formulas = [];
  formulas = [PrintStatus, email, TotalCOGS, ActualShipping, ChargedShipping];
  var formulas1 = [];
  
  // set up Sheet2 (order details) formulas
  var MAP = '=vlookup(R[0]C[-2],Items!C[-4]:C[0],2,false)*R[0]C[-3]';
  var WholeSale = '=vlookup(R[0]C[-3],Items!C[-5]:C[-1],5,false)*R[0]C[-4]';
  var Profit = '=(R[0]C[-2]-R[0]C[-1])';
  var CostCat = '=if(R[0]C[-4]="na","Shipping Costs","Cost of Goods Sold")';
  var formulas2 = [MAP, WholeSale, Profit, CostCat];
  var formulas3 = [];
  
  
  var spreadsheet = SpreadsheetApp.openById(invoicesheetID);
  var a1sheet = spreadsheet.getSheetByName("Invoices");
  var sheet = spreadsheet.setActiveSheet(a1sheet);
  var rows1 = lastValue("N",sheet);
  
  var a2sheet = spreadsheet.getSheetByName("Order Details");
  var sheet2 = spreadsheet.setActiveSheet(a2sheet);
  var rows2 = lastValue("E",sheet2);
  
  
  
  var last_row = sheet.getLastRow();
  ///==============================test just adding formulas to new rows
  // if(rows1 != last_row ){
  //var last_col = sheet.getRange(last_row-order_summary.length,13,order_summary.length,4);
  
  //}
  
  //var last_no_form = sheet.setActiveRange(range)  
  
  
  
  // set the formula array to be duplicated as many times as there are rows
  var last_row = sheet.getLastRow();
  for (var p =0;p<last_row-1; p++){
    formulas1[p] = formulas;
  }
  rows1 = rows1+1;
  var form_range = "N2" + ":" + "R" + last_row;
  //var form_range = "N" + rows1 + ":" + "Q" + last_row;
  var cell = sheet.getRange(form_range);
  cell.setFormulasR1C1(formulas1);
  
  
  
  // set the formula array to be duplicated as many times as there are rows
  var last_row2 = sheet2.getLastRow();
  for (var p =0;p<last_row2-1; p++){
    formulas3[p] = formulas2;
  }
  rows2 = rows2 + 1;
  //var form_range2 = "E" + rows2+ ":" + "H" + last_row2;
  var form_range2 = "E2" + ":" + "H" + last_row2;
  var cell2 = sheet2.getRange(form_range2);
  cell2.setFormulasR1C1(formulas3);
  
  cell.copyTo(sheet.getRange(form_range), {contentsOnly: true});
  cell2.copyTo(sheet2.getRange(form_range2), {contentsOnly: true});
  
  //}//row 101-97
  
  removeDuplicates(sheet);
  removeDuplicates(sheet2);
  //sort the order summary by paid invoice date
  form_range = "A2:" + "Q" + last_row;
  cell = sheet.getRange(form_range);
  cell.sort({column: 9, ascending: false});
  
  form_range2 = "A2:" + "H" + last_row2;
  cell = sheet2.getRange(form_range2);
  cell.sort({column: 1, ascending: false});
  
  //export the sheet to csv for import into Stamps.com
  //export the sheet to xlsx for easy manipulation
  exportSheet(invoicesheetID);
  
}//ending function




function importData(wkbkID) {
  importCSVFromGoogleDrive('USPSshipments.csv', wkbkID, "Shipments");
  
};

function exportSheet(wkbkID){
  downloadSheetAsXlsx(wkbkID,"Invoices","xlsx");
  downloadSheetAsXlsx(wkbkID,"Invoices","csv");
  downloadSheetAsXlsx(wkbkID,"TransferOut","csv");
};

function test2(){
  downloadSheetAsXlsx("14VzdU_0cZZ_x4Ab3zTF1SCUS5mTYZEdkZxhLFae2T0w","Invoices","xlsx");
  downloadSheetAsXlsx("14VzdU_0cZZ_x4Ab3zTF1SCUS5mTYZEdkZxhLFae2T0w","Invoices","csv"); 
}




//99

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
  
  sheet.appendRow(arraytoAppend);
    var form_range = "S2:U2" ;
    
    var cell = sheet.getRange(form_range);
    cell.setFormulasR1C1(cell.getValues());
   sheet.insertRowsAfter(1,arraytoAppend.length);
  //sheet.
  var range = sheet.getRange("S2:U"+(arraytoAppend.length+1));
  range.setValues(arraytoAppend);
  
};


function appendRowstoSheet(workbookID, sheetNm, arraytoAppend){
  var spreadsheet = SpreadsheetApp.openById(workbookID);
  var a1sheet = spreadsheet.getSheetByName(sheetNm);
  var sheet = spreadsheet.setActiveSheet(a1sheet);
  
  for( var x = 0; x < arraytoAppend.length;x++){
    sheet.appendRow(arraytoAppend[x]); 
  }
  var form_range = "S2:U"+(arraytoAppend.length+1);
  sheet.insertRowsAfter(1,arraytoAppend.length);
  
  var range = sheet.getRange(form_range);
  range.setValues(arraytoAppend);
   range.setFormulasR1C1(range.getValues());
  
  
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

  var tmp;
  //New Email Style as of 5.15.2018
  if(method == 3){
    var ty_indx = emlbody.search("= Thank You! =");
    var cus_indx = emlbody.search("== Bill To\: ==");
    var sh_indx = emlbody.search("== Ship To\: ==");
    var ordrsmm = emlbody.search("== Order Summary ==");
    var dt_indx = emlbody.search("Date:");
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
    tmp = ordrsumm_tbl2.match(/Items Subtotal\s*\$([0-9]*\.[0-9]*)/);
  var amt_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0;
  tmp = ordrsumm_tbl2.match(/Order Discount\s*\-\$([0-9]*\.[0-9]*)/);
  var disc_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0;
  tmp = ordrsumm_tbl2.match(/  Subtotal\s*\$([0-9]*\.[0-9]*)/);
  var amt_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0;
  tmp = ordrsumm_tbl2.match(/  Tax\s*\$([0-9]*\.[0-9]*)/);
  var tax_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0;
  tmp = ordrsumm_tbl2.match(/S&H\s*\$([0-9]*\.[0-9]*)/);
  var sh_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0; 
  tmp = ordrsumm_tbl2.match(/S&H\s*Tax\s*\$([0-9]*\.[0-9]*)/);
  var shtax_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0; 
  tmp = ordrsumm_tbl2.match(/Total\s*\$([0-9]*\.[0-9]*)/);
  var tp_nb = (tmp && tmp[1]) ? tmp[1].trim() : 0;
    
    var shp45_new = [];
    var shp46_new = [];
    var nmmmmm;
    tmp = emlbody2.match(/Subtotal[\s]*\$([\0-9\s]+)(Tax)/);
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
     
      
      tmp = shp46_new[shp46_new.length-1].match(/^([A-z\s]*)\s*\,\s([A-Z]{2})\s([0-9]{5})/);
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


