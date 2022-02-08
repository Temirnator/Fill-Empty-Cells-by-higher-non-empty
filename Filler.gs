function onOpen() { 
  var ui = SpreadsheetApp.getUi(); 
  ui.createMenu('Fill').addItem('Fill empty', 'filler').addToUi(); 
  ui.createMenu('Show Sheet ID').addItem('Show Active Sheet ID','getSheetbyId' ).addToUi(); 
} 
 
function filler(){ 
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = getSheetById(696812339); 
  //SpreadsheetApp.getUi().alert(sheet.getName()); 
  var lastRow = sheet.getLastRow(); 
  var lastColumn = sheet.getLastColumn(); 
  var lastCell = sheet.getRange(lastRow, lastColumn); 
  //SpreadsheetApp.getUi().alert(lastRow); 
  //SpreadsheetApp.getUi().alert(lastColumn); 
  var i=1; 
  var filledcell; 
  //SpreadsheetApp.getUi().alert(sheet.getRange(1,1,lastRow,lastColumn).getCell(i,1)); 
  for(i=4;i<=lastRow;i++){ 
    if(sheet.getRange(1,1,lastRow,lastColumn).getCell(i,1).isBlank()){ 
      filledcell = sheet.getRange(1,1,lastRow,lastColumn).getCell(i-1,1).getValue(); 
      sheet.getRange(1,1,lastRow,lastColumn).getCell(i,1).setValue(filledcell).getValue(); 
    } 
  } 
   
} 
 
function getSheetbyId(){ 
  var sheetID = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId(); 
  SpreadsheetApp.getUi().alert(sheetID); 
} 
 
function getSheetById(id) { 
  return SpreadsheetApp.getActive().getSheets().filter( 
    function(s) {return s.getSheetId() === id;} 
  )[0]; 
}
