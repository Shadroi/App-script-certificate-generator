// Google Form PDF Generator using Slides as template
// @shadroi

function onFormSubmit(e) {
  
  var value = e.values;
  
  var email = value[1];
  var name = value[2];
  
  Logger.log(value);
  
  var link = createPdf(name,email);
  
  sendEmail(name,email,link);
}

function sendEmail(name,email,link){
  
  var subject = "ICLS 2019 - Thanks for joining!"
  var htmlBody = "<strong>Hello, "+name+"! </strong><br><br> Here's your free <a href="+link+">certificate</a>. Thanks for listening!<br><br>"+"<strong>Best,<br>Shad</strong>"
  
  GmailApp.sendEmail(email, subject, '', {htmlBody: htmlBody});
  
}

// Replace this with ID of your template slides.
var TEMPLATE_ID = '1RnhhRe7J5jSVnljQkg1UJaXKUrB3hSH5Ll1rE40JaME' //Certificate Slide

//PDF FILE NAME
var PDF_FILE_NAME = ''

// Create New PDF from submission using the template

function createPdf(name,email) {
  
  //name = "Shad"
  //email = "shad@shadroi.com"

  if (TEMPLATE_ID === '') {
    
    SpreadsheetApp.getUi().alert('TEMPLATE_ID needs to be defined')
    return
  }

  // Set up the slides and the spreadsheet access
  
  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(),
      copyId = copyFile.getId(),
      copySlide = SlidesApp.openById(copyId),
      copySBody = copySlide.getSlides()[0];

  
 
  // Replace the keys with the spreadsheet values by having correct and matching headers
  
  //name
    copySBody.replaceAllText('<<' + "Name" + '>>', name)
   
  //email
    copySBody.replaceAllText('<<' + "Email" + '>>', email)
  
  // Create the PDF file, rename it if required and delete the slides copy
    
  copySlide.saveAndClose()

  //CREATING NEW FILE
  var newFile = DriveApp.getRootFolder().createFile(copyFile.getAs('application/pdf'))  
  //NEW FILE NAME
  PDF_FILE_NAME = 'ICLS 2019 - ' + name + '.pdf';
  newFile.setName(PDF_FILE_NAME)
   
  //TRASH EDITED FILE
  copyFile.setTrashed(true)
  
  return newFile.getUrl()
  
  SpreadsheetApp.getUi().alert('New PDF file created in the root of your Google Drive - '+newFile.getUrl())
  
}
