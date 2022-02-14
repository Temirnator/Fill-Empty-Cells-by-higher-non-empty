function getSheetbyId(){ 
  var sheetID = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId(); 
  SpreadsheetApp.getUi().alert(sheetID); 
} 
var CONF;

function onOpen() {
  var ui = SpreadsheetApp.getUi();  //For convenience
  //Creates a button 'Email_01' in the top panel of the Spreadsheet
  ui.createMenu('Email_01').addItem('Send all', 'sender_email_01').addItem('Send selected','slctsender_email_01').addToUi();
  ui.createMenu('Email_02').addItem('Send all', 'sender_email_02').addItem('Send selected','slctsender_email_02').addToUi();
  ui.createMenu('Email_03').addItem('Send all', 'sender_email_03').addItem('Send selected','slctsender_email_03').addToUi();
  ui.createMenu('Email_04').addItem('Send all', 'sender_email_04').addItem('Send selected','slctsender_email_04').addToUi();
  ui.createMenu('Email_05').addItem('Send all', 'sender_email_05').addItem('Send selected','slctsender_email_05').addToUi();
  
  ui.createMenu('Show Sheet ID').addItem('Show Active Sheet ID','getSheetbyId' ).addToUi(); 
  ui.createMenu('Folder Populator').addItem('Populate Kid Pages','folderPopulator' ).addToUi(); 
  ui.createMenu('RowID').addItem('Update Row IDs','rowID').addToUi();
}

function searchTutor(kidname){
  var ss = SpreadsheetApp.openById("1AaMNmhopqblI0o87lCO1DEbHFg_6bH81vtmJH4Boj8Y");
  var tutorsheet = ss.getSheetByName("Tutors");
  var lastrow = getLastRowSpecial(tutorsheet.getRange("A2:A").getValues());
  for(var i=1;i<=lastrow;i++){
    if(tutorsheet.getRange(i+1,9).getValue()==kidname){
      return i+1;
    }
  }
}

function getLastRowSpecial(range){
  var rowNum = 0;
  var blank = false;
  for(var row = 0; row < range.length; row++){

    if(range[row][0] === "" && !blank){
      rowNum = row;
      blank = true;

    }else if(range[row][0] !== ""){
      blank = false;
    };
  };
  return rowNum;
};

function rowID(){
  var ss = SpreadsheetApp.openById("1AaMNmhopqblI0o87lCO1DEbHFg_6bH81vtmJH4Boj8Y");
  var kidssheet = ss.getSheetByName("Kids");
  //SpreadsheetApp.getUi().alert(lastrow);
  //var avals = kidssheet.getRange("A2:A").getValues();
  //var alast = avals.filter(String).length;
  var lastrow = getLastRowSpecial(kidssheet.getRange("A:A").getValues());
  kidssheet.getRange("G2:G").clearContent();
  for(var i=1;i<lastrow;i++){
    kidssheet.getRange(i+1,7).setRichTextValue(SpreadsheetApp.newRichTextValue()
    .setText('Click to go to A' + (i+1))
    .setLinkUrl('#gid=' + kidssheet.getSheetId() + '&range=A' + (i+1))
    .build())
    //"#gid=680657599&range=A"+(i+1));
  }
  //SpreadsheetApp.getUi().alert(lastrow);
  return lastrow;
}

//Create folder if does not exists only
function createKidFolder(folderID, folderName,templateFolderID){
  var sourceFolder = DriveApp.getFolderById(folderID);
  var subFolders = sourceFolder.getFolders();
  var doesntExists = true;
  var newFolder = '';
  var templateFolder = DriveApp.getFolderById(templateFolderID);
  var files = templateFolder.getFiles();
  // Check if folder already exists.
  while(subFolders.hasNext()){
    var folder = subFolders.next();
    
    //If the name exists return the id of the folder
    if(folder.getName() === folderName){
      doesntExists = false;
      newFolder = folder;
      return newFolder.getId();
    }
  }
    //If the name doesn't exists, then create a new folder
    if(doesntExists == true){
      //If the file doesn't exists
      newFolder = sourceFolder.createFolder(folderName);
      while (files.hasNext()) {
        var file = files.next();
        file.makeCopy(file.getName(), newFolder);

      }    
    return newFolder.getId();
  }
};

function folderPopulator(){ 
  var sourceFOLDER_ID = '1q_lA9L9E7E0FtR0JWVP-rhxVGLLBRhEg';
  var templateFOLDER_ID = '1JpTG5zLxKymLDUaVJLbDcWEqT09Rp3eC'
  ////////////
  var ss = SpreadsheetApp.openById("1AaMNmhopqblI0o87lCO1DEbHFg_6bH81vtmJH4Boj8Y");
  //https://docs.google.com/spreadsheets/d/1AaMNmhopqblI0o87lCO1DEbHFg_6bH81vtmJH4Boj8Y/edit?usp=sharing
  var tutorsheet = ss.getSheetByName("Tutors");
  var kidssheet = ss.getSheetByName("Kids");
  var i=0;
  //var avals = kidssheet.getRange("A2:A").getValues();
  //var alast = avals.filter(String).length;
  var lastrow = getLastRowSpecial(kidssheet.getRange("A2:A").getValues());
  ////////////
  var myFolderID; 
  var folder;
  for(i=1;i<=lastrow;i++){
    myFolderID = createKidFolder(sourceFOLDER_ID, kidssheet.getRange(i+1,1).getValue(),templateFOLDER_ID);
    folder = DriveApp.getFolderById(myFolderID).addEditor(tutorsheet.getRange(searchTutor(kidssheet.getRange(i+1,1).getValue()),3).getValue());
    kidssheet.getRange(i+1,8).setValue("https://drive.google.com/drive/folders/"+myFolderID);
  }
};

function globaltutor(){
  var user_email = null;                                                         //TO DO!!!!!
  var user_name =null;
  var user_phone =null;
  var user_position =null;
  var login_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accounts');

  var data=null;

  for(var i=1;i<login_sheet.getLastRow();i++){
    data = login_sheet.getRange(i+1,1,1,login_sheet.getLastColumn()).getValues()[0];
    if(Session.getActiveUser().getEmail()==data[0]){
      user_email = data[0];
      user_name =data[1];
      user_phone =data[2];
      user_position =data[3];
      break;
    }
  }
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main Page')
  var tutor = {//Tutor class
      fullname : null//some fields change as the row changes (fullname, notionPage, email,etc)
    , notionPage: null
    , phone: null
    , email: null
    , kidname: null
    , kidphone: null
    , kidlang: null
    , kidgrade: null
    , kidgender: null
    , kidSVcategory: null
    , kidLink: null
    , parentname: null
    , parentphone: null
    , maincoord: mainSheet.getRange(2,2).getValue()//others remain the same for all rows (maincoord, maincoordnum, maincoordemail,etc)
    , maincoordnum: mainSheet.getRange(3,2).getValue()
    , maincoordemail: mainSheet.getRange(4,2).getValue()
    , username: user_name
    , userposition: user_position
    , useremail: user_email
    , userphone: user_phone
    , instructions: mainSheet.getRange(7,2).getValue()
    , telegramchat: mainSheet.getRange(8,2).getValue()
    };
    //SpreadsheetApp.getUi().alert(tutor.username); 
  return tutor;
}

function sender(subject, template_name){
  
  var ui = SpreadsheetApp.getUi();// For convenience
  CONF = ui.alert('Confirm to Send all', '',ui.ButtonSet.YES_NO);//Add Alert that asks Confirm Yes or No

  if (CONF == ui.Button.YES) {//Continue if the Confirm is Yes
  } else {
    return;//Don't execute the function otherwise
  }

  var tutorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tutors');//For convenience
  var temp = HtmlService.createTemplateFromFile(template_name);//Generate an HTML template named temp from the 'template_name' template
  var lastrow = tutorSheet.getLastRow();//Get last row number of the current Spreadsheet
  var tutor = globaltutor();

  var history_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('emails_history');//For convenience
  var date = Utilities.formatDate(new Date(), "GMT+6", 'MMMM dd, yyyy [HH:mm:ss]');
  
  if(tutor.useremail==null){
    ui.alert('You are not authorized to send emails','',ui.ButtonSet.OK);
    return;
  }

  var data=null;//Variable for storing data from one row to an array

  //For loop the iterates starting from the 13th row to the last nonempty row of the Spredsheet
  for(var i=1; i<lastrow;i++){
    data = tutorSheet.getRange(i+1, 1, 1, the_sheet.getLastColumn()).getValues()[0];//Transfers data from row to an array named 'data'
    tutor.fullname=data[0];//Transfer data from 'data' array to 'tutor' class
    tutor.phone= data[1];
    tutor.email= data[2];
    tutor.kidname= data[7];
    tutor.kidphone = data[8];
    tutor.kidlang= data[9];
    tutor.kidgrade= data[10];
    tutor.kidgender = data[11];
    tutor.kidSVcategory = data[12];
    tutor.kidLink = data[13];
    tutor.parentname= data[14];
    tutor.parentphone= data[15];
    temp.tutor=tutor;//get values for a template using fields from the 'tutor' class
    var message = temp.evaluate().getContent();//generate a message from the template and current variables
    MailApp.sendEmail({//Send email 
    to: tutor.email//to 'tutor.email'
    , subject: subject//with subject of the email as 'subject' variable contents,
    , htmlBody: message//with a message generated from an HTML template 'message'
    });
    ////////////////////////////////////////////////////////////////
    //the_sheet.getRange(the_range.getRow()+i,1).setValue('sent');//////////////////////////////
    //After sending the message, writes the word 'sent' to the first column of the current row//
    ////////////////////////////////////////////////////////////////////////////////////////////
    //Emails History//
    //////////////////
    history_sheet.getRange(history_sheet.getLastRow()+1,1).setValue(date);
    history_sheet.getRange(history_sheet.getLastRow(),2).setValue(tutor.useremail);
    history_sheet.getRange(history_sheet.getLastRow(),3).setValue(tutor.username);
    history_sheet.getRange(history_sheet.getLastRow(),4).setValue(tutor.email);
    history_sheet.getRange(history_sheet.getLastRow(),5).setValue(tutor.fullname);
    history_sheet.getRange(history_sheet.getLastRow(),6).setValue(subject);
  }
  return;
}

function slctsender(subject, template_name){
  var ui = SpreadsheetApp.getUi();  //For convenience
  CONF = ui.alert('Confirm','', ui.ButtonSet.YES_NO);//Add Alert that asks Confirm Yes or No

  if (CONF == ui.Button.YES) {        //Continue if the Confirm is Yes
  } else {
    return;                           //Don't execute the function otherwise
  }

  var tutorSheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tutors');//For convenience  
  var temp = HtmlService.createTemplateFromFile(template_name);//Generate an HTML template named temp from the 'template_name' template
  var activeSheet = SpreadsheetApp.getActiveSheet();//For convenience
  var range =  activeSheet.getActiveRange();//Gets active range
  var tutor = globaltutor();

  var history_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('emails_history');//For convenience
  var date = date = Utilities.formatDate(new Date(), "GMT+6", 'MMMM dd, yyyy [HH:mm:ss]');

  if(tutor.useremail==null){//
    ui.alert('You are not authorized to send emails','',ui.ButtonSet.OK);
    return;
  }

  var data=null;                    //Variable for storing data from one row to an array
  
  //For loop that iterates over the selected range. 
  //     !!!ATTENTION!!!
  //Works properly only if the selected range is unseparated.
  //That is, if you choose two ranges, say 'A2:B6, E10:E12' it WON'T send emails for rows 2 to 6 and 10 to 12
  //Instead, it will get the number of rows (in our case it will be 8, since 8 rows are selected)
  //And it will iterate from the first selected (A2 is the first cell, so the first row will be row #2)
  //for 8 rows (that is, from 2nd row to 9th row)
  for(var i=0; i<range.getNumRows();i++){
    data = tutorSheet.getRange(range.getRow()+i,1, 1, tutorSheet.getLastColumn()).getValues()[0];//Transfers data from row to an array named 'data'
    tutor.fullname=data[0];         //Transfer data from 'data' array to 'tutor' class
    tutor.phone= data[1];
    tutor.email= data[2];
    tutor.kidname= data[7];
    tutor.kidphone = data[8];
    tutor.kidlang= data[9];
    tutor.kidgrade= data[10];
    tutor.kidgender = data[11];
    tutor.kidSVcategory = data[12];
    tutor.kidLink = data[13];
    tutor.parentname= data[14];
    tutor.parentphone= data[15];
    temp.tutor=tutor;               //get values for a template using fields from the 'tutor' class
    var message = temp.evaluate().getContent();//generate a message from the template and current variables
    MailApp.sendEmail({             //Send email   
      to: tutor.email               //to 'tutor.email'
      , subject: subject            //with subject of the email as 'subject' variable contents,
      , htmlBody: message           //with a message generated from an HTML template 'message'
    });
    //the_sheet.getRange(the_range.getRow()+i,1).setValue('sent');
    //After sending the message, writes the word 'sent' to the first column of the current row
    history_sheet.getRange(history_sheet.getLastRow()+1,1).setValue(date);
    history_sheet.getRange(history_sheet.getLastRow(),2).setValue(tutor.useremail);
    history_sheet.getRange(history_sheet.getLastRow(),3).setValue(tutor.name);
    history_sheet.getRange(history_sheet.getLastRow(),4).setValue(tutor.email);
    history_sheet.getRange(history_sheet.getLastRow(),5).setValue(tutor.fullname);
    history_sheet.getRange(history_sheet.getLastRow(),6).setValue(subject);

  }
  return;

}

////////////
//Email 01//
////////////
function sender_email_01(){
  var subject = 'IHelper Email 01 - Welcome letter (and important information)';//'Subject' variable
  var template_name = 'ihelper_email_01';
  sender(subject,template_name);
}

function slctsender_email_01(){
  var subject = 'IHelper Email 01 - Welcome letter (and important information)';//'Subject' variable
  var template_name = 'ihelper_email_01';
  slctsender(subject,template_name);
}

////////////
//Email 02//
////////////
function sender_email_02(){
  var subject = 'IHelper Email 02 - Kid Assignment';//'Subject' variable
  var template_name = 'ihelper_email_02';
  sender(subject,template_name);
}

function slctsender_email_02(){
  var subject = 'IHelper Email 02 - Kid Assignment';//'Subject' variable
  var template_name = 'ihelper_email_02';
  slctsender(subject,template_name);
}


////////////
//Email 03//
////////////
function sender_email_03(){
  var subject = 'IHelper Email 03 - Kid Performance Assessment';
  var template_name = 'ihelper_email_03';
  sender(subject,template_name);
}

function slctsender_email_03(){
  var subject = 'IHelper Email 03 - Kid Performance Assessment';
  var template_name = 'ihelper_email_03';
  slctsender(subject,template_name);
}

////////////
//Email 04//
////////////
function sender_email_04(){
  var subject = 'IHelper Email 04 - Weekly Report';//'Subject' variable
  var template_name = 'ihelper_email_04';
  sender(subject,template_name);
}

function slctsender_email_04(){
  var subject = 'IHelper Email 04 - Weekly Report';//'Subject' variable
  var template_name = 'ihelper_email_04';
  slctsender(subject,template_name);
}

////////////
//Email 05//
////////////
function sender_email_05(){
  var subject = 'IHelper Email 05 - End of Period';//'Subject' variable
  var template_name = 'ihelper_email_05';
  sender(subject,template_name);
}

function slctsender_email_05(){
  var subject = 'IHelper Email 05 - End of Period';//'Subject' variable
  var template_name = 'ihelper_email_05';
  slctsender(subject,template_name);
}

