function getStatusEmail(){
  var tes_threads = GmailApp.search("from:abc@xyz.com is:unread subject:\"sales\"");
  message = tes_threads;

  var openSheet = SpreadsheetApp.openByUrl('');

  var activeSheet = openSheet.setActiveSheet(openSheet.getSheetByName("Status Log"));
  var message_count = 0;

  for (var i = 0; i < tes_threads.length; i++) {

    var message = tes_threads[i].getMessages();
    // Logger.log(i)

    for (var j = 0; j < message.length; j++){
      var arr = [];
      message_count = message_count+1;

      GmailApp.markMessageRead(message[j]);

      var recipient = message[j].getBody();
      var dateTime = message[j].getDate()

      arr.push(dateTime);

      let str = recipient.replace(/[\n\r]/g, '>>');
      // console.log(str); 
      var patt1 = /(?<=<b>)([^%]+?)(?=<\/b>)/g; 
      var patt2 = /(?<=<li>)([^%]+?)(?=<\/li>)/g;
      var result1 = recipient.match(patt1);
      // Logger.log(result1)
      var result2 = str.match(patt2);
      
      subject = message[j].getSubject();
      // if(result2!==null){
      //     Logger.log(result2_clean);
      //   }
      
      if(result1!==null){
        custid = result1[0];
        let custid_clean = custid.replace("+", "");
        let custid_clean2 = custid_clean.replace(/\*/g, "");
        arr.push(custid_clean2);

        account = result1[1];
        // Logger.log(account)
        let account_clean = account.replace(/[\n\r]/g, ' ');
        // Logger.log(account_clean);
        let account_clean2 = account_clean.replace('&amp;', "&");
        let account_clean3 = account_clean2.replace(/\*/g, "");
        // Logger.log(account_clean2);
        arr.push(account_clean3);
        // Logger.log(custid+account);
        if(result2!==null){
          status = "flagged";
          arr.push(status);
          
          var result2_clean = result2[0].replace(/>>>>/g, "");
          reason = result2_clean;
          arr.push(reason);
          Logger.log(reason);
          var announce = "*" + account_clean3 + "* - The custid number " + custid_clean2 + " has been flagged";
        }
        else{
          status = "connected";
          arr.push(status);

          reason = "";
          arr.push(reason);
          var announce = "*" + account_clean3 + "* - The custid number " + custid_clean2 + " is no longer flagged";
        }
        // Logger.log(arr);
          
        arr.push(announce); 
        activeSheet.appendRow(arr);
      }
     
    }
    
  }
  multiSortColumns(openSheet, "Status Log");
  openSheet.toast('Update complete. ' + message_count + ' New Notification');
  // var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}