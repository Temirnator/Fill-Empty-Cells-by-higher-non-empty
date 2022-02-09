
function getSheetbyId(){ 
  var sheetID = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId(); 
  SpreadsheetApp.getUi().alert(sheetID); 
} 


function onOpen() {
  var ui = SpreadsheetApp.getUi();  //For convenience
  ui.createMenu('Email_01').addItem('Send all', 'sender_email_01').addItem('Send selected','slctsender_email_01').addToUi();
  //                                //Creates a button 'Email_01' in the top panel of the Spreadsheet
  ui.createMenu('Email_02').addItem('Send selected','slctsender_email_02').addToUi();
  ui.createMenu('Email_03').addItem('Send all', 'sender_email_03').addItem('Send selected','slctsender_email_03').addToUi();
  //   ui.createMenu('Email_04.1').addItem('Send all', 'sender_email_041').addItem('Send selected','slctsender_email_041').addToUi();
  //   ui.createMenu('Email_04.2').addItem('Send all', 'sender_email_042').addItem('Send selected','slctsender_email_042').addToUi();
  //   ui.createMenu('Email_04.3').addItem('Send all', 'sender_email_043').addItem('Send selected','slctsender_email_043').addToUi();
  //   ui.createMenu('Email_05').addItem('Send all', 'sender_email_05').addItem('Send selected','slctsender_email_05').addToUi();
  ui.createMenu('Show Sheet ID').addItem('Show Active Sheet ID','getSheetbyId' ).addToUi(); 

}
var CONF;

///////////////////////////////////////////
//Single tutor variable for all functions//
//It is not actually a global variable/////
///////////////////////////////////////////
function globaltutor(){
  var user_email = Session.getActiveUser().getEmail();
  var user_time = null;                                                         //TO DO!!!!!
  var user_name =null;
  var user_phone =null;
  var user_position =null;
  var login_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('');

  var data=null;

  for(var i=1;i<=login_sheet.getLastRow();i++){
    data = login_sheet.getRange(i,1,1,login_sheet.getLastColumn()).getValues()[0];
    if(Session.getActiveUser().getEmail()==data[0]){
      user_email = Session.getActiveUser().getEmail();
      user_name =data[1];
      user_phone =data[2];
      user_position =data[3];
      break;
    }else{
      user_email=null;
    }
  }
  var the_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tutors');//For convenience
  var tutor = {//Tutor class
      fullname : null//some fields change as the row changes (fullname, notionPage, email,etc)
    , notionPage: null
    , phone: null
    , email: null
    , coord: null
    , coordnum: null
    , coordemail: null
    , kidname: null
    , kidgrade: null
    , kidlang: null
    , kidspecialty: null
    , parentname: null
    , parentphone: null
    , maincoord: the_sheet.getRange(2,2).getValue()//others remain the same for all rows (maincoord, maincoordnum, maincoordemail,etc)
    , maincoordnum: the_sheet.getRange(3,2).getValue()
    , maincoordemail: the_sheet.getRange(4,2).getValue()
    , personname: user_name
    , personposition: user_position
    , personemail: user_email
    , personephone: user_phone
    , persontime: user_time
    , instructions: the_sheet.getRange(9,2).getValue()
    , telegramchat: the_sheet.getRange(10,2).getValue()
    };
  return tutor;
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
  var subject = 'IHelper Email 03 - Weekly Report';
  var template_name = 'ihelper_email_03';
  sender(subject,template_name);
}

function slctsender_email_03(){
  var subject = 'IHelper Email 03 - Weekly Report';
  var template_name = 'ihelper_email_03';
  slctsender(subject,template_name);
}

//////////////
//Email 04.1//
//////////////

//////////////
//Email 04.2//
//////////////

//////////////
//Email 04.3//
//////////////

////////////
//Email 05//
////////////

////////////
//Email 06//
////////////

/////////////////////
//Send All Function//
/////////////////////
function sender(subject, template_name){
  
  var ui = SpreadsheetApp.getUi();// For convenience
  CONF = ui.alert('Confirm to Send all', '',ui.ButtonSet.YES_NO);//Add Alert that asks Confirm Yes or No

if (CONF == ui.Button.YES) {//Continue if the Confirm is Yes
} else if (CONF== ui.Button.NO) {
  return;//Don't execute the function if the Confirm is No
} else {
  return;//Don't execute the function if the Alert is closed
}

  var the_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tutors');//For convenience
  var temp = HtmlService.createTemplateFromFile(template_name);//Generate an HTML template named temp from the 'template_name' template
  var lastrow = the_sheet.getLastRow();//Get last row number of the current Spreadsheet
  var tutor = globaltutor();

  var history_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('emails_history');//For convenience
  var date = date = Utilities.formatDate(new Date(), "GMT+6", 'MMMM dd, yyyy [HH:mm:ss]');
  
  if(tutor.personemail==null){//CHANGE THIS!!! IT HAS NO MEANING
    ui.alert('You are not authorized to send emails','',ui.ButtonSet.OK);
    return;
  }

  var data=null;//Variable for storing data from one row to an array

  //For loop the iterates starting from the 13th row to the last nonempty row of the Spredsheet
  for(var i=13; i<= lastrow;i++){
    data = the_sheet.getRange(i, 2, 1, the_sheet.getLastColumn()).getValues()[0];//Transfers data from row to an array named 'data'
    tutor.fullname=data[0];//Transfer data from 'data' array to 'tutor' class
    tutor.notionPage= data[1];
    tutor.phone= data[2];
    tutor.email= data[3];
    tutor.coord= data[4];
    tutor.coordnum= data[5];
    tutor.coordemail= data[6];
    tutor.kidname= data[7];
    tutor.kidgrade= data[8];
    tutor.kidlang= data[9];
    tutor.kidspecialty= data[10];
    tutor.parentname= data[11];
    tutor.parentphone= data[12];
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
    history_sheet.getRange(history_sheet.getLastRow(),2).setValue(tutor.personemail);
    history_sheet.getRange(history_sheet.getLastRow(),3).setValue(tutor.personname);
    history_sheet.getRange(history_sheet.getLastRow(),4).setValue(tutor.email);
    history_sheet.getRange(history_sheet.getLastRow(),5).setValue(tutor.fullname);
    history_sheet.getRange(history_sheet.getLastRow(),6).setValue(subject);
  }
  return;
}


//////////////////////////
//Send Selected Function//
//////////////////////////
function slctsender(subject, template_name){
  var ui = SpreadsheetApp.getUi();  //For convenience
  CONF = ui.alert('Confirm','', ui.ButtonSet.YES_NO);//Add Alert that asks Confirm Yes or No

if (CONF == ui.Button.YES) {        //Continue if the Confirm is Yes
} else if (CONF== ui.Button.NO) {
  return;                           //Don't execute the function if the Confirm is No
} else {
  return;                           //Don't execute the function if the Alert is closed
}

  var the_sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tutors');//For convenience  
  var temp = HtmlService.createTemplateFromFile(template_name);//Generate an HTML template named temp from the 'template_name' template
  var activeSheet = SpreadsheetApp.getActiveSheet();//For convenience
  var the_range =  activeSheet.getActiveRange();//Gets active range
  var tutor = globaltutor();

  var history_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('emails_history');//For convenience
  var date = date = Utilities.formatDate(new Date(), "GMT+6", 'MMMM dd, yyyy [HH:mm:ss]');

  if(tutor.personemail==null){//CHANGE THIS!!!!!!! IT HAS NO MEANING
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
  for(var i=0; i<the_range.getNumRows();i++){
    data = the_sheet.getRange(the_range.getRow()+i,2, 1, the_sheet.getLastColumn()).getValues()[0];//Transfers data from row to an array named 'data'
    tutor.fullname=data[0];         //Transfer data from 'data' array to 'tutor' class
    tutor.notionPage= data[1];
    tutor.phone= data[2];
    tutor.email= data[3];
    tutor.coord= data[4];
    tutor.coordnum= data[5];
    tutor.coordemail= data[6];
    tutor.kidname= data[7];
    tutor.kidgrade= data[8];
    tutor.kidlang= data[9];
    tutor.kidspecialty= data[10];
    tutor.parentname= data[11];
    tutor.parentphone= data[12];
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
    history_sheet.getRange(history_sheet.getLastRow(),2).setValue(tutor.personemail);
    history_sheet.getRange(history_sheet.getLastRow(),3).setValue(tutor.personname);
    history_sheet.getRange(history_sheet.getLastRow(),4).setValue(tutor.email);
    history_sheet.getRange(history_sheet.getLastRow(),5).setValue(tutor.fullname);
    history_sheet.getRange(history_sheet.getLastRow(),6).setValue(subject);

  }
  return;
}
