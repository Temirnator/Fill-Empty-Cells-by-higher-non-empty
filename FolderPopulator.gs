
//Create folder if does not exists only
function createKidFolder(folderID, folderName,templateFolderID){
  var parentFolder = DriveApp.getFolderById(folderID);
  var subFolders = parentFolder.getFolders();
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
    newFolder = parentFolder.createFolder(folderName);
    while (files.hasNext()) {
      var file = files.next();
      file.makeCopy(file.getName(), newFolder);
    }
    return newFolder.getId();
  };

  

};

function start(){ 
  var parentFOLDER_ID = '1nEXUZJ6RhFDUE_qidnMuc8yN6jY8zKEU';
  var templateFOLDER_ID = '1OvZ-UEo76z5POqCS3SoyrXF2onSYvvo6'
  ////////////
  var ss = SpreadsheetApp.openById("1a1d-Ypei8ignWlMkMpie2zcU7vhboNmFEeO3uavWHiA");
  var sheet = ss.getSheetByName("Tutors");
  var i=0;
  var avals = sheet.getRange("I1:I").getValues();
  var alast = avals.filter(String).length;
  ////////////
  var myFolderID; 
  for(i=13;i<=alast;i++){
    myFolderID = createKidFolder(parentFOLDER_ID, sheet.getRange(i,9).getValue(),templateFOLDER_ID);
  }

};

