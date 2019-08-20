// Google Form PDF Generator using Slides as template
// @shadroi
// inspired from @andrewroberts

// Replace this with ID of your template slides.
var TEMPLATE_ID = '1Y3GLCaytUEh69e8j_e1cS-qgHN1cL_cwnq7LNKtPhGk' //DSC Certificate Slide

//PDF FILE NAME
var PDF_FILE_NAME = ''

/**
 *  Add a menu.
**/

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu('PDF Generator')
    .addItem('Create PDF', 'showDialog')
    .addToUi()

}

function showDialog() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Please confirm',
     'Did you select the ROW you want to generate?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // if "Yes".
    createPdf();
  } else {
    // if "No" close
  }
}


//Find and Replace hot keys inside '<< >>' and Create New PDF from template

function createPdf() {

  if (TEMPLATE_ID === '') {
    
    SpreadsheetApp.getUi().alert('TEMPLATE_ID needs to be defined')
    return
  }

  // Set up the slides and the spreadsheet access
  
  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(),
      copyId = copyFile.getId(),
      copySlide = SlidesApp.openById(copyId),
      copySBody = copySlide.getSlides()[0],
      activeSheet = SpreadsheetApp.getActiveSheet(),
      numberOfColumns = activeSheet.getLastColumn(),
      activeRowIndex = activeSheet.getActiveRange().getRowIndex(),
      activeRow = activeSheet.getRange(activeRowIndex, 1, 1, numberOfColumns).getValues(),
      headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues(),
      columnIndex = 0;
  
 
  // Replace the keys with the spreadsheet values
 
  for (;columnIndex < headerRow[0].length; columnIndex++) {
    
    copySBody.replaceAllText('<<' + headerRow[0][columnIndex] + '>>', 
                         activeRow[0][columnIndex])
    
  }
  
  // Create the PDF file, rename it if required and delete the slides copy
    
  copySlide.saveAndClose()

  var newFile = DriveApp.createFile(copyFile.getAs('application/pdf'))  
  
  PDF_FILE_NAME = 'DSC Lead 2018 - ' + activeRow[0][0] + ' - ' + activeRow[0][1] + '.pdf';

  if (PDF_FILE_NAME !== '') {
  
    newFile.setName(PDF_FILE_NAME)
  } 
  
  copyFile.setTrashed(true)
  
  SpreadsheetApp.getUi().alert('New PDF file created in the root of your Google Drive')
  
} // 
