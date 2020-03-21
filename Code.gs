function adhoc_export()
{
var invoicesheetID = initialSetup();
dotheworksheet(invoicesheetID);
}


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
     tmp = emlbody.match(/Teddy Bear Jacket\s/g) ;
  var teddy =  tmp ? tmp.length : 0;
     tmp = emlbody.match(/Ivy\s/g) ;
  var ivy =  tmp ? tmp.length : 0;
 tmp = emlbody.match(/Avery\s/g) ;
  var avery =  tmp ? tmp.length : 0;
 tmp = emlbody.match(/Noelle\s/g) ;
  var noelle =  tmp ? tmp.length : 0;
 tmp = emlbody.match(/Xoe\s/g) ;
  var xoe =  tmp ? tmp.length : 0;
   tmp = emlbody.match(/Elizabeth\s/g) ;
  var eliz =  tmp ? tmp.length : 0;
 tmp = emlbody.match(/Mark\s/g) ;
  var mark =  tmp ? tmp.length : 0;
  tmp = emlbody.match(/Jane\s/g) ;
  var jane =  tmp ? tmp.length : 0;
  
  
  //25,25,37,26
 //ivy,avery,noelle,xoe
  //var DisleggingsTC = emlbody.match(/Mimi Tops/g);
  //var leggingsOS = emlbody.match(/Disney Leggings Leggings TC/g);
  var OrdrPcs = [leggingsOS+blklgsOS+vdayleggingsOS, lynnae,cassie-elgCassie, sarah-elgSarah,carly-elgCarly,classic-Disclassic,DisleggingsLXL,randy-Disrandy,joy,perfect+ptank,madison,leggingsTC+blklgsTC+vdayleggingsTC,julia,lindsay,maxi,Disrandy,irma-Disirma,DisleggingsTC,DisleggingsOS,leggingsTC2+blklgsTC2+vdayleggingsTC2,DisleggingsSM,amelia,DisleggingsTween,azure,monroe, nicole,lola,shirley-elgShirley,mimi,Disirma, Disclassic,leggingsSM+leggingsXL+vdayleggingsSM+vdayleggingsXL,leggingsTW+vdayleggingsTW,elgDebbie,elgShirley,elgCarly,elgDeanne,elgCassie,elgSarah,jaxon+harvey,maria,maurine,debbie,jane,cici,tank,christy,dani,grga,carol,jessie,marly,mae+adeline+sariah+scarlett+gracie,nikki,amber,gigi,valentina,hudson,lucille,denim,mitzi,mimi,ellie,liv,michelle,jax,emily,iris,hudsonLS,teddy,ivy,avery,noelle,xoe,eliz,mark,jane];
  //Logger.log(OrdrPcs);
  var WSprices = [10.5,18,14,30,25,16,10,16,25,17,23,10.5,18,21,21,21,15,12.5,12.5,12,10,31,11,14,10.5,23,21,22,34,22,20,8.5,9.5,25,25,39,70,20,42,50,34,36,24,24,35,12.5,17,27.5,32.5,27,25,25,5,23,21,15,22,17,30,35,17,30,26,17,24,26,25,16,19,26,25,25,37,26,18,20,22];
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
    var a2sheet = spreadsheet.getSheetByName("Order Details");
    var sheet2 = spreadsheet.setActiveSheet(a2sheet);
    sheet.appendRow(["Invoice Name","Order#","Popup#","Shipping Name","Address1", "Address2","City","State","Zip","Invoice Date/Time","Subtotal","Discounts","Shipping","Tax","Total Sale","Email","Status","Total COGS","Actual Shipping","Items Invoiced"]);
    sheet2.appendRow(["Order#","Item Qty","Item Description","Item Size","MAP","Wholesale Cost","Profit","Cost Category"]);
  }
  //if(unreads.length > 0){
    for (var a=0;a < unreads.length; a++)
    //for (var a=0;a < 100; a++) // comment out previous and uncomment this one to re-run all
    {
      var message = unreads[a].getMessages();
      var ThreadNm = unreads[a].getFirstMessageSubject()
      var recipient = unreads[a].getMessages()[0].getTo();
      var emlbody = unreads[a].getMessages()[0].getPlainBody();
      //Logger.log(emlbody);
      //var cogs_count = OrderDeets(emlbody);
      recipient = recipient.substring(recipient.lastIndexOf("<") + 1 , recipient.lastIndexOf(">"));
      var Ordr_Nmbr = ThreadNm.substring(ThreadNm.lastIndexOf("Number ")+7,ThreadNm.length);     
      var dup_ordr = 0;
      var dup_email = 0;
      
      for( var v = 0; v < ordersinSheet.length; v++)
      {
        if(ordersinSheet[v].toString() == (Ordr_Nmbr))
        {dup_ordr=1;break}
      }
      
      if(dup_ordr == 0)
      {
        var order_summary = getOrderSummaryFromMessageNewBless(message[0]);
        all_added_orders.push(order_summary);
      }//end if dup_ordr
      
      unreads[a].markRead();
    } // ending a for loop
    
    
    if(all_added_orders.length == 1) {
      appendRowtoSheet(invoicesheetID, "Invoices", all_added_orders[0]);
    }
    else{
      appendRowstoSheet(invoicesheetID, "Invoices", all_added_orders);}
    
  
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
    
    //removeDuplicates(sheet);
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