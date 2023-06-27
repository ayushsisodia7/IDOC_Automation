function extractIDOCReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  var i;
  var existingIDs = sheet.getDataRange().getValues().map(row => row[4]);
  
  var now = new Date();
  var startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var endOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate() +1);

  const secondsSinceEpoch = (date) => Math.floor(date.getTime() / 1000);
  var today = new Date();  
  const after = new Date();
  const before = new Date();
  var ch = 23;
  var cm = 59;
  var cs = 59;
  before.setHours(ch,cm ,cs);

  var ah = 0;
  var am = 0;
  var as = 0;
  after.setHours(ah,am,as);

  var currentDate = new Date();
  var startDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate());
  var endDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() + 1);


  var fetch = "subject:(51 STATUS IDOC REPORT)" + "after:" + secondsSinceEpoch(after) + " " + "before:" + secondsSinceEpoch(before);
  console.log(fetch);

  var threads = GmailApp.search(fetch); 

  var headers = ["Partner Number of Sender", "Partner Number of Receiver", "Direction for IDoc", "Logical Message Type", "Idoc_Num", "Idoc Created on", "Idoc Created at", "Status", "Error_Message"];

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  var headerFont = headerRange.getFontFamily();
  headerRange.setFontWeight('bold').setBackground('#ADD8E6').setFontFamily(headerFont); 

  if (sheet.getLastRow() == 0) {
    sheet.appendRow(headers);
  }
  //console.log(startDate,endDate);
  var newIDs = [];

  for (var i = 0; i < threads.length; i++) 
  {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) 
    {
      var messageDate = messages[j].getDate();
      if (messageDate >= startDate && messageDate < endDate)
      {
      var subject = messages[j].getSubject().toUpperCase();
      if (subject.indexOf("IDOC REPORT") > -1) 
      {
        var attachments = messages[j].getAttachments();
        if (attachments.length > 0) 
        {
          for (var k = 0; k < attachments.length; k++) 
          {
            var attachment = attachments[k];
            var attachmentContents = attachment.getDataAsString();
            var lines = attachmentContents.split("\n");
            var idocNumIndex = headers.indexOf("Idoc_Num");
            for (var l = 1; l < lines.length; l++) 
            {
              var values = lines[l].split("\t");
              var idocNum = values[idocNumIndex];
              if (existingIDs.indexOf(idocNum) == -1) 
              {                 
                sheet.appendRow(values);
                existingIDs.push(idocNum);
                newIDs.push(idocNum);
                
              }
            }
          }
        }
      }
      }
    }
  }
  sheet.deleteColumn(3);
  sheet.deleteColumn(3);
  sheet.deleteColumn(5);
  sheet.deleteColumn(5);
  removeDuplicates();

  var targetRange = sheet.getRange(2, 1, sheet.getLastRow(), 5); 
  var targetValues = targetRange.getValues();
  var error_msg="";

  var system="";
  for(i=0;i<sheet.getLastRow();i++)
  {
    error_msg=targetValues[i][4];
    system=targetValues[i][0];
    if(error_msg.includes("does not exist") && (system.includes("LAP") || system.includes("LWP")))
    {
      console.log("Send mail to Ricardo");
      console.log(error_msg);
      var final_msg = "Hello Ricardo,\n" + "For IDOC Number: " + targetValues[i][2] +"\n" + error_msg + "\nPlease do the needful.";
      var subject_mail = targetValues[i][1] + " - " + targetValues[i][3] + " " + "STATUS IDOC REPORT";
      GmailApp.sendEmail("xyz@colpal.com",subject_mail,final_msg,{cc:'abc@colpal.com'});
      
    }
    else if(error_msg.includes("does not exist") && (system.includes("HUP") || system.includes("HWP")))
    {
      console.log("Send Mail to Brian Huber & Brian Lane");
      console.log(error_msg);
      var final_msg = "Hello Brian,\n" + "For IDOC Number: " + targetValues[i][2] +"\n"+ error_msg + "\nPlease do the needful.";
      var subject_mail = targetValues[i][1] + " - " + "51 STATUS IDOC REPORT";
      GmailApp.sendEmail("xyz@hillspet.com",subject_mail,final_msg,{cc:'abc@colpal.com'});
    }
    else if((error_msg.includes("Material") && error_msg.includes("is not defined")) &&  (system.includes("LAP") || system.includes("LWP")))
    {
      console.log("Send mail to Ricardo");
      console.log(error_msg);
      var final_msg = "Hello Ricardo,\n" + "For IDOC Number: " + targetValues[i][2] +"\n" + error_msg + "\nPlease do the needful.";
      var subject_mail = targetValues[i][1] + " - " + targetValues[i][3] + " " + "STATUS IDOC REPORT";
      GmailApp.sendEmail("xyz@colpal.com",subject_mail,final_msg,{cc:'abc@colpal.com'});
    }
    else if((error_msg.includes("Material") && error_msg.includes("is not defined")) && (system.includes("HUP") || system.includes("HWP")))
    {
      console.log("Send Mail to Brian Huber & Brian Lane");
      console.log(error_msg);
      var final_msg = "Hello Brian,\n" + "For IDOC Number: " + targetValues[i][2] +"\n"+ error_msg + "\nPlease do the needful.";
      var subject_mail = targetValues[i][1] + " - " + "51 STATUS IDOC REPORT";
      GmailApp.sendEmail("xyz@hillspet.com",subject_mail,final_msg,{cc:'abc@colpal.com'});
    }
    
  }
}


function removeDuplicates() {
var sheet = SpreadsheetApp.getActiveSheet();
var data = sheet.getDataRange().getValues();
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
sheet.clearContents();
sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}
