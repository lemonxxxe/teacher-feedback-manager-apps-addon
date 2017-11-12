var teacherfolders = {};
var templateId = '1csy2kO0-NymbIC23gWnhpCY81G9CJrYpHWfgoS3HwUw'; //template id for google form


function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Survey Manager')
    .addItem('Set Up Spreadsheet', 'setupSheet')
    .addItem('Generate Forms and Folders', 'generateForms')
    .addItem('Share Folders', 'shareFolders')
    .addItem('Update Response Numbers', 'updateNumResponses')
    .addToUi(); 

   
  var htmlOutput = HtmlService
       .createHtmlOutputFromFile('index')
       .setTitle('Student Feedback Surveys Manager');
  
  SpreadsheetApp.getUi()
       .showSidebar(htmlOutput);
}

function setTitle(title) {
  SpreadsheetApp.getActive().rename(title);
}

function setupSheet() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var projectProperties = documentProperties.getProperties();
  
  if (!projectProperties.hasOwnProperty('setup') || !projectProperties.setup) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var headers = [['teacher_first_name', 'teacher_email', 'course', 'period/section', 'semester', 'year', 'num_responses', 'student_survey_url', 'response_summary_url', 'edit_url', 'form_id']];
    var range = sheet.getRange(1, 1, 1, 11);
    range.setValues(headers);
    range.setBackground('#cfe2f3');
    range = sheet.getRange('G2:K');
    range.setBackground('#efefef');
    documentProperties.setProperty(setup, true);
  } else {
    SpreadsheetApp.getUi().alert('hell no');
  }
}

function preGenerateCheck () {
  var alreadyForms = []
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var rangeVals = range.getValues();
  var formId, thisForm;
  for (var i = 1; i < rangeVals.length; i++) {
    formId = rangeVals[i][10];
    if (formId !== '') {
      alreadyForms.push(formId);
      thisForm = FormApp.openById(formId);
      Logger.log(thisForm);
    }
  }
}



function generateForms() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    'Are you sure you want to do this?',
    'It cannot be undone without deleting the project folder.',
    ui.ButtonSet.YES_NO);
  
  if (confirm == ui.Button.YES) {
    keepOn();
  }
  
  function keepOn() {
    var newForm, title, template, destination, editUrl, surveyUrl, formId, targetRange, updatedData, summaryUrl;
    
    var folders = {};
    var sheet = SpreadsheetApp.getActiveSheet();
    var masterFolder = DriveApp.getFileById(sheet.getParent().getId()).getParents().next();
    var range = sheet.getDataRange();
    if (!masterFolder.getFoldersByName(SpreadsheetApp.getActiveSpreadsheet().getName()).hasNext()){
      masterFolder.createFolder(SpreadsheetApp.getActiveSpreadsheet().getName());
    }
    destination = DriveApp.getFoldersByName(SpreadsheetApp.getActiveSpreadsheet().getName()).next();
    createNameFolders();
    
    var currentVals = range.getValues();
    var currentCourse;
    
    currentVals.forEach(function(val, ind, arr){
      // ind = 0 is header row
      if (ind !== 0) {
        currentCourse = new Course(val);
        template = DriveApp.getFileById(templateId);
        
        //teacherFolders[val[0]].email = val[1];
        title = val[0] + ' : ' + val[2] + ' (period ' + val[3] + ') - ' + val[4] + ' ' + val[5];
        //template is the a copy of the student feedback survey
        
        newForm = currentCourse.createForm(teacherFolders[currentCourse.teacherEmail].folder, template);
        //newForm.addEditor(val[1]);
        editUrl = newForm.getUrl();
        formId = newForm.getId();
        summaryUrl = FormApp.openById(formId).getSummaryUrl();
        surveyUrl = FormApp.openById(formId).getPublishedUrl();
        arr[ind].splice(7, 4, surveyUrl, summaryUrl, editUrl, formId);
        targetRange = sheet.getRange(ind+1, 1, 1, 11);
        targetRange.setValues([arr[ind]]);
      }
    });
  }
}

function createNameFolders() {
      var sheet = SpreadsheetApp.getActiveSheet();
      var names = sheet.getRange('A2:B' + sheet.getDataRange().getLastRow());
      var namesArr = [];
      var emailsArr = [];
      var destination = DriveApp.getFoldersByName(SpreadsheetApp.getActiveSpreadsheet().getName()).next();
      var parent = destination;
      names.getValues().forEach(function(val, ind, arr){
        Logger.log(val[1]);
        if (emailsArr.indexOf(val[1]) === -1){
          emailsArr.push(val[1]);
          namesArr.push(val[0]);
        }
      });
      Logger.log(namesArr);
      Logger.log(emailsArr);
      namesArr.forEach(function(val, ind, arr){
        if (!parent.getFoldersByName(val).hasNext()) {
          teacherFolders[val] = {folder: parent.createFolder(val)};
        } else {
          SpreadsheetApp.getUi().alert("something's wrong -- there's already a folder with name: " + val);
        }
      });
    }



function shareFolders() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var projectFolder = DriveApp.getFoldersByName(SpreadsheetApp.getActiveSpreadsheet().getName()).next();
  for (user in folders) {
    folders[user].folder.addEditor(folders[user].email);
  }
}
  

function updateNumResponses() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var currVals = range.getValues();
  var currRowVals, currForm, numResponseCell;
  for (var i = 0; i < currVals.length; i++) {
    if (i > 0) {
      currRowVals = currVals[i];
      currForm = FormApp.openById(currRowVals[10]);
      numResponseCell = sheet.getRange(i+1, 7);
      numResponseCell.setValue(currForm.getResponses().length);
    }
  }
}

function Course(teacherName, teacherEmail, courseName, period, semester, year) {
  this.teacherName = teacherName;
  this.teacherEmail = teacherEmail;
  this.courseName = courseName;
  this.period = period;
  this.semester = semester;
  this.year = year;
  
}

Course.prototype.createForm = function(destination, template) {
  var title = this.teacherName + ' : ' + this.courseName + ' (period ' + this.period + ') - ' + this.semester + ' ' + this.year;
  template.makeCopy(title, destination); 
}