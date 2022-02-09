function getDuration(start, end){
  var duration;
  //var startArr = start.split("/");
  //var endArr = end.split("/");
  
}

function kidsPageCreate(title){
  var ssNew = SpreadsheetApp.create(title);
  return ssNew.getUrl();
}


function GetSheetName() {
return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function myFunction() {
//var emailRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1").getRange("B2");
//var emailAddress = emailRange.getValues();
// Send Alert Email.
//var message = 'This is your Alert email!'; // Second column
//var subject = 'Your Google Spreadsheet Alert';
//MailApp.sendEmail(emailAddress, subject, message);


//To do:
//0) Emailer for notifications on start dates and deadlines
//1) Emailer about reports on how many tasks were completed this week
//2) Watch: https://www.youtube.com/watch?v=a8y8YZTTsGs&ab_channel=LeandroZubrezki
}
function getCoordinator(){
  var user_email = Session.getActiveUser().getEmail();
  var user_time = null;                                                         //TO DO!!!!!
  var user_name =null;
  var user_phone =null;
  var user_main_position =null;
  var user_secondary_position =null;
  var login_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');

  var data=null;

  for(var i=1;i<=login_sheet.getLastRow();i++){
    data = login_sheet.getRange(i,1,1,login_sheet.getLastColumn()).getValues()[0];
    if(Session.getActiveUser().getEmail()==data[3]){
      user_email = Session.getActiveUser().getEmail();
      user_name =data[0];
      user_phone =data[4];
      user_main_position =data[1];
      user_secondary_position =data[2];
      break;
    }else{
      user_email=null;
    }
  }
  var coordinator = {//Coordinator class
      fullname: user_name
    , mainposition: user_main_position
    , email: user_email
    , phone: user_phone
    , time: user_time    
    };
    return coordinator;
}



function onEdit(e){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All_Tasks');
  var coordinator = getCoordinator();
  var history_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('History');

    if ((sheet.getActiveCell().getRow() >= 4)&&(sheet.getActiveCell().getColumn() == 1)&&(sheet.getActiveCell().getValue()== true)){
      sheet.getRange(sheet.getActiveCell().getRow(),13).setValue(new Date);
      sheet.getRange(sheet.getActiveCell().getRow(),12).setValue(coordinator.fullname);
      history_sheet.getRange(history_sheet.getLastRow()+1,2).setValue(sheet.getRange(sheet.getActiveCell().getRow(),2).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),3).setValue(sheet.getRange(sheet.getActiveCell().getRow(),3).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),4).setValue(sheet.getRange(sheet.getActiveCell().getRow(),4).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),5).setValue(sheet.getRange(sheet.getActiveCell().getRow(),5).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),6).setValue(sheet.getRange(sheet.getActiveCell().getRow(),6).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),7).setValue(sheet.getRange(sheet.getActiveCell().getRow(),7).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),8).setValue(sheet.getRange(sheet.getActiveCell().getRow(),8).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),9).setValue(coordinator.fullname);
      history_sheet.getRange(history_sheet.getLastRow(),10).setValue(new Date);
      history_sheet.getRange(history_sheet.getLastRow(),11).setValue("ticked");
    }else if((sheet.getActiveCell().getRow() >= 4)&&(sheet.getActiveCell().getColumn() == 1)&&(sheet.getActiveCell().getValue()== false)){
      sheet.getRange(sheet.getActiveCell().getRow(),13).clearContent();
      sheet.getRange(sheet.getActiveCell().getRow(),12).clearContent();
      history_sheet.getRange(history_sheet.getLastRow()+1,2).setValue(sheet.getRange(sheet.getActiveCell().getRow(),2).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),3).setValue(sheet.getRange(sheet.getActiveCell().getRow(),3).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),4).setValue(sheet.getRange(sheet.getActiveCell().getRow(),4).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),5).setValue(sheet.getRange(sheet.getActiveCell().getRow(),5).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),6).setValue(sheet.getRange(sheet.getActiveCell().getRow(),6).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),7).setValue(sheet.getRange(sheet.getActiveCell().getRow(),7).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),8).setValue(sheet.getRange(sheet.getActiveCell().getRow(),8).getValue());
      history_sheet.getRange(history_sheet.getLastRow(),9).setValue(coordinator.fullname);
      history_sheet.getRange(history_sheet.getLastRow(),10).setValue(new Date);
      history_sheet.getRange(history_sheet.getLastRow(),11).setValue("unticked");
    }
  
}



/*function sendReminder() {
  var name = '';
  var birthDate = '';
  var emailAddress = '';
  var todayDate = new Date();
  allData.slice(1, allData.length).forEach(function (birthDayArr) {
    birthDate = new Date(birthDayArr[2]);
    if (birthDate.getDate() == todayDate.getDate() && birthDate.getMonth() == todayDate.getMonth()) {
      name = birthDayArr[0];
      emailAddress = birthDayArr[1];
      try {
        GmailApp.sendEmail(emailAddress, "Happy birthday!", "Happy birthday, " + name);
      }
      catch (e) {
        try {
          GmailApp.sendEmail(errorEmail, "Error sending email", "There was an error sending email to " + name + " - " + emailAddress + " with script ID " + ScriptApp.getScriptId);
        }
        catch (e) {
          Logger.log("Failed to send email for " + name + " to " + emailAddress + ". Make sure the data in the sheet is valid (valid email address). If it is, is the API down?");
        }
      }
    }
  })
}

function createTrigger() {
  ScriptApp.newTrigger('sendReminder')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
}
 */
