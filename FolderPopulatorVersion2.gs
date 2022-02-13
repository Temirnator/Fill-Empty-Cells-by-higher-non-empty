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
    };
  };
  //If the name doesn't exists, then create a new folder
  if(doesntExists == true){
    //If the file doesn't exists
    newFolder = sourceFolder.createFolder(folderName);
    while (files.hasNext()) {
      var file = files.next();
      file.makeCopy(file.getName(), newFolder);
    }
    return newFolder.getId();
  };

  

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
  var avals = kidssheet.getRange("A2:A").getValues();
  var alast = avals.filter(String).length;
  ////////////
  var myFolderID; 
  for(i=2;i<=alast;i++){
    myFolderID = createKidFolder(sourceFOLDER_ID, tutorsheet.getRange(i,9).getValue(),templateFOLDER_ID);
    kidssheet.getRange(i,8).setValue(myFolderID);
  }

};

function onOpen() {
  var ui = SpreadsheetApp.getUi();  //For convenience
  //Creates a button 'Email_01' in the top panel of the Spreadsheet  
  ui.createMenu('Folder Populator').addItem('Populate Kid Pages','folderPopulator' ).addToUi(); 

}
