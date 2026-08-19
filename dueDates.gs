/**
* @OnlyCurrentDoc
*/

/* Start of Configuration */
function getConfig(option) {
  if (typeof option === 'undefined') {
    return;
  }
  var item;
  switch (option) {
    case 'header':
      // headers to process
      // type: 1: regular, 2: date, 3: quiz-only date, 5: boolean, 11: calculated
      item = [ {
        'title' : 'Title',
        'field' : 'title',
        'type' : 'string'
      }, {
        'title' : 'Due',
        'field' : 'due_at',
        'type' : 'date',
        'daysEnd' : true,
      }, {
        'title' : 'Available from',
        'field' : 'unlock_at',
        'type' : 'date'
      }, {
        'title' : 'Available until',
        'field' : 'lock_at',
        'type' : 'date',
        'daysEnd' : true,
      }, {
        'title' : 'Show Answers',
        'field' : 'show_correct_answers_at',
        'type' : 'date',
        'location' : 'Quiz'
      }, {
        'title' : 'Hide Answers',
        'field' : 'hide_correct_answers_at',
        'type' : 'date',
        'location' : 'Quiz'
      }, {
        'title' : 'Points',
        'field' : 'points_possible',
        'type' : 'number',
        'location' : 'Assignment'
      }, {
        'title' : 'Published',
        'field' : 'published',
        'type' : 'boolean'
      }, {
        'title' : 'Type',
        'field' : 'type',
        'type' : 'string',
        'required' : true
      }, {
        'title' : 'Canvas ID',
        'field' : 'id',
        'type' : 'integer',
        'required' : true
      } ];
      break;
    case 'sheetName' : 
      // name of the sheet to use
      item = 'Dates';
      break;
  }
  return item;
}
/* End of Configuration */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Canvas')
  .addItem('Specify Course', 'getCourseDlg')
  .addItem('Load Due Dates', 'listDueDates')
  .addItem('Save Due Dates', 'setDueDates')
  .addSeparator()
  .addItem('Shift All Dates...', 'shiftAllDatesDialog')
  .addItem('Remap Date Range...', 'remapDateRangeDialog')
  .addItem('Apply Availability Rules...', 'availabilityDefaultsDialog')
  .addSeparator()
  .addItem('Hide Times', 'hideTimes')
  .addItem('Show Times', 'showTimes')
  .addSeparator()
  .addItem('Configure API Settings', 'configurationDialog')
  .addItem('Forget API Settings', 'resetApiSettings')
  .addSeparator()
  .addItem('Show Help', 'helpDialog')
  .addToUi();
  return;
}

/**
* @function getHeaders return header information for spreadsheet This is where you should modify information you want displayed
*/
function getHeaders(list) {
  var obj = {};
  var hasList = typeof list == 'object' && Array.isArray(list);
  var headerInfo = getConfig('header');
  for (var i = 0; i < headerInfo.length; i++) {
    var headerItem = headerInfo[i];
    var key = headerItem.field.split('_')[0];
    var header = {
      'name' : headerItem.title,
      'field' : headerItem.field,
      'type' : headerItem.type,
      'location' : typeof headerItem.location === 'undefined' ? null : headerItem.location,
      'daysEnd' : typeof headerItem.daysEnd === 'undefined' ? null : headerItem.daysEnd,
      'key' : key,
      'c1' : i,
      'c2' : hasList ? list.indexOf(headerItem.title) : -1
    };
    obj[key] = header;
  }
  return obj;
}

/**
* @function getCourseDlg Ask for the course ID and clear existing data if it is given
*/
function getCourseDlg() {
  var courseId = getCourseDialog();
  if (courseId) {
    getDataSheet(true);
    listDueDates();
  }
  return;
}

/**
* @function itemToRow Convert JSON item into spreadsheet format
* @param {object} item an object containing information from the API
* @param {object} headers to use in the conversion
* @returns {Array}
*/
function itemToRow(item, headers) {
  if (typeof headers === 'undefined') {
    headers = getHeaders();
  }
  var type;
  var dt;
  if (typeof item.type !== 'undefined') {
    type = item.type;
  } else {
    type = typeof item.title === 'undefined' ? 'Assignment' : 'Quiz';
  }
  var row = [];
  for ( var j in headers) {
    if (headers.hasOwnProperty(j)) {
      var hdr = headers[j];
      var value = item[hdr.field];
      if (hdr.key == 'title' && type != 'Quiz') {
        value = item.name;
      } else if (hdr.key == 'type') {
        value = type;
      } else if (hdr.type == 'boolean') {
        value = item[hdr.key] ? 1 : 0;
      }
      if (typeof value === 'undefined') {
        row.push('');
      } else {
        if (hdr.type == 'date') {
          dt = fromIso8601(value);
          row.push(dt ? dt.substr(0, 19) : '');
        } else {
          row.push(value);
        }
      }
    }
  }
  return row;
}

function getCourseId() {
  var courseId;
  try {
    var userProperties = PropertiesService.getUserProperties();
    courseId = userProperties.getProperty('courseid');
    if (!courseId) {
      throw new Error('You must specify a courseID through the Canvas menu before proceeding.');
    }
  } catch (e) {
    Logger.log(e);
    throw e; // Bubble to UI
  }
  return courseId;
}

function getDueDates(courseId) {
  var data = {};
  var quizAssignments = {};
  var key;
  try {
    if (typeof courseId === 'undefined') {
      courseId = getCourseId();
    }
    if (typeof courseId === 'undefined') {
      throw new Error('You must specify a course ID before running this');
    }
    var quizList = canvasAPI('GET /api/v1/courses/:course_id/quizzes', {
      ':course_id' : courseId
    }, [ 'id', 'title', 'due_at', 'unlock_at', 'lock_at', 'show_correct_answers_at', 'hide_correct_answers_at', 'published', 'assignment_id' ]);
	
    var assignmentList = canvasAPI('GET /api/v1/courses/:course_id/assignments', {
      ':course_id' : courseId
    }, [ 'id', 'name', 'due_at', 'unlock_at', 'lock_at', 'published', 'points_possible' ]);
	var assignmentById = {};

if (typeof assignmentList !== 'undefined') {
  for (var i = 0; i < assignmentList.length; i++) {
    assignmentById[assignmentList[i].id] =
      assignmentList[i];
  }
}
    if (typeof quizList !== 'undefined') {
      for (var i = 0; i < quizList.length; i++) {
        key = 'q' + quizList[i].id;
        if (typeof assignmentList !== 'undefined' && quizList[i].assignment_id) {
          var quizAssignmentId = quizList[i].assignment_id;
          quizAssignments[quizAssignmentId] = quizList[i].id;
        }
        quizList[i].type = 'Quiz';

if (
  quizList[i].assignment_id &&
  assignmentById[quizList[i].assignment_id]
) {
  quizList[i].points_possible =
    assignmentById[
      quizList[i].assignment_id
    ].points_possible;
}


data[key] = quizList[i];
      }
    }
    


    if (typeof assignmentList !== 'undefined') {
      for (var i = 0; i < assignmentList.length; i++) {
        var assignmentId = assignmentList[i].id;
        if (typeof quizAssignments[assignmentId] === 'undefined') {
          key = 'a' + assignmentList[i].id;
          assignmentList[i].type = 'Assignment';
          data[key] = assignmentList[i];
        }
      }
    }
  } catch (e) {
    Logger.log(e);
    throw e; // Bubble to UI
  }
  return data;
}

function getDataSheet(create) {
  var dsheet;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (typeof ss === 'undefined') {
      throw new Error('No active spreadsheet');
    }
    var userProperties = PropertiesService.getUserProperties();
    var courseId = userProperties.getProperty('courseid');
    if (!courseId) {
      throw new Error('You must specify the course ID before you can change the dates.');
    }
    var dateSheetName = getConfig('sheetName');
    dsheet = ss.getSheetByName(dateSheetName);
    if (dsheet == null) {
      ss.insertSheet(dateSheetName, 0);
      dsheet = ss.getSheetByName(dateSheetName);
    }
    if (dsheet == null) {
      throw new Error('Unable to create "' + dateSheetName + '" sheet');
    }
    var headerInfo = getConfig('header');
    if (create) {
      var numColumns = dsheet.getMaxColumns();
      var newColumns = headerInfo.length;
      dsheet.insertColumnsAfter(numColumns, newColumns + 4);
      dsheet.deleteColumns(1, numColumns);
    }
  } catch (e) {
    Logger.log(e);
    throw e; // Bubble to UI
  }
  return dsheet;
}

function listDueDates() {
  try {
    var courseId = getCourseId();
    var existingData = getDueDates(courseId);
    var data = [];
    if (typeof existingData === 'undefined' || Object.keys(existingData).length == 0) {
      throw new Error('No data returned. Cowardly refusing to do anything stupid.');
    }
    var dsheet = getDataSheet(1);
    var hdrs = getHeaders();
    var headers = [];
    for ( var j in hdrs) {
      if (hdrs.hasOwnProperty(j)) {
        headers.push(hdrs[j].name);
      }
    }
    var cols = headers.length;
    data.push(headers);  
    for ( var key in existingData) {
      if (existingData.hasOwnProperty(key)) {
        var row = itemToRow(existingData[key], hdrs);
        data.push(row);
      }
    }
    var rows = data.length;
    dsheet.getRange(1,1,rows,cols).setValues(data);
    if (dsheet.getMaxRows() > rows) {
      dsheet.deleteRows(rows + 1, dsheet.getMaxRows() - rows);
    }
    // Automatically format the due dates.
    formatSpreadsheet(courseId);
  } catch (e) {
    Logger.log(e);
    if (typeof showError === 'function') showError('Failed to Load Due Dates', e);
    return;
  }
  return;
}

function formatSpreadsheet(courseId) {
  try {
    var datetimeFmt = 'yyyy-MM-dd hh:mm';
    var dsheet = getDataSheet();
    if (dsheet.getRange(1, 1).isBlank()) {
      throw new Error('You have no data to process. Refusing to continue');
    }
    if (typeof courseId === 'undefined') {
      courseId = getCourseId();
    }
    var validationRule = null;
    if (courseId) {
      var course = canvasAPI('GET /api/v1/courses/:course_id', {
        ':course_id' : courseId
      }, [ 'start_at', 'end_at' ]);
      if (course) {
        var startDate = fromIso8601(course.start_at, true);
        var endDate = fromIso8601(course.end_at, true);
        var helpText = 'This entry must be a valid date';
        validationRule = SpreadsheetApp.newDataValidation();
        if (startDate) {
          if (endDate) {
            validationRule.requireDateBetween(startDate, endDate);
            helpText += ' between the course dates of ' + startDate.toDateString() + ' and ' + endDate.toDateString();
          } else {
            validationRule.requireDateAfter(startDate);
            helpText += ' after the course start date of ' + startDate.toDateString();
          }
        } else if (endDate) {
          validationRule.requireDateBefore(endDate);
          helpText += ' before the course end date of ' + endDate.toDateString();
        } else {
          validationRule.requireDate();
        }
        validationRule.setHelpText(helpText + '.');
        validationRule.build();
      }
    }
    var booleanValidationRule = SpreadsheetApp.newDataValidation().requireValueInList([ '0', '1' ], false).setHelpText('1 = Yes, 0 = No').build();
    var typeValidationRule =
        SpreadsheetApp.newDataValidation().requireValueInList([ 'Assignment', 'Quiz' ], false).setAllowInvalid(false).setHelpText('Can be "Assignment" or "Quiz", but changing this will not work.')
    .build();
    dsheet.setActiveRange(dsheet.getRange(1, 1));
    var range = dsheet.getDataRange();
    var rowCount = range.getNumRows() - 1;
    var colCount = range.getNumColumns();
    var rows = range.getValues();
    var sheetHeaders = rows[0];
    var hdrs = getHeaders(sheetHeaders);
    dsheet.setFrozenRows(1);
    dsheet.getRange(1, 1, 1, colCount).setFontWeight('bold');
    for ( var item in hdrs) {
      if (hdrs.hasOwnProperty(item)) {
        var hdr = hdrs[item];
        if (hdr.c2 < 0) {
          continue;
        }
        var col = 1 + hdr.c2;
        dsheet.getRange(1, col, 1 + rowCount, 1).setHorizontalAlignment(hdr.key == 'title' ? 'left' : 'right');
        if (hdr.type === 'date' && validationRule) {
          dsheet.getRange(2, col, rowCount, 1).setNumberFormat(datetimeFmt);
          dsheet.getRange(2, col, rowCount, 1).setDataValidation(validationRule);
        } else if (hdr.type === 'boolean') {
          dsheet.getRange(2, col, rowCount, 1).setDataValidation(booleanValidationRule);
        } else if (hdr.field === 'type') {
          dsheet.getRange(2, col, rowCount, 1).setDataValidation(typeValidationRule);
        }
        dsheet.autoResizeColumn(col);
        var autowidth = dsheet.getColumnWidth(col);
        dsheet.setColumnWidth(col, (hdr.key == 'title' && autowidth > 200) ? 200 : 10 + autowidth);
      }
    }
    dsheet.getRange(2, 1, rowCount, colCount).sort([ {
      'column' : 1 + hdrs.due.c2,
      'ascending' : true
    }, {
      'column' : 1 + hdrs.title.c2,
      'ascending' : true
    } ]);
  } catch (e) {
    Logger.log(e);
    throw e; // Bubble to UI
  }
  return;
}

function hideTimes() {
  formatDueDates(false);
  return;
}

function showTimes() {
  formatDueDates(true);
  return;
}

function formatDueDates(showTimes) {
  if (typeof showTimes === 'undefined') {
    showTimes = true;
  }
  var dateFmt = 'yyyy-MM-dd';
  var datetimeFmt = 'yyyy-MM-dd hh:mm';
  var hr;
  var min;
  try {
    var dsheet = getDataSheet();
    if (dsheet.getRange(1, 1).isBlank()) {
      throw new Error('You have no data to process. Refusing to continue');
    }
    dsheet.setActiveRange(dsheet.getRange(1, 1));
    var range = dsheet.getDataRange();
    var rowCount = range.getNumRows() - 1;
    var rows = range.getValues();
    var sheetHeaders = rows[0];
    var hdrs = getHeaders(sheetHeaders);
    for ( var item in hdrs) {
      if (hdrs.hasOwnProperty(item)) {
        var hdr = hdrs[item];
        if (hdr.c2 < 0 || hdr.type !== 'date') {
          continue;
        }
        var col = 1 + hdr.c2;
        if (showTimes) {
          dsheet.getRange(2, col, rowCount, 1).setNumberFormat(datetimeFmt);
        } else {
          dsheet.getRange(2, col, rowCount, 1).setNumberFormat(dateFmt);
          for (var j = 1; j <= rowCount; j++) {
            if (!rows[j][col - 1]) {
              continue;
            }
            var value = rows[j][col - 1];
            
            // Defensively ensure the cell value is actually a Date object 
            // before attempting to extract hours and minutes
            if (Object.prototype.toString.call(value) !== '[object Date]') {
              continue; 
            }
            
            var needsTime = false;
            hr = value.getHours();
            min = value.getMinutes();
            if (hdr.daysEnd) {
              if (hr !== 23 || min !== 59) {
                needsTime = true;
              }
            } else {
              if (hr !== 0 || min !== 0) {
                needsTime = true;
              }
            }
            if (needsTime) {
              dsheet.getRange(j + 1, col).setNumberFormat(datetimeFmt);
            }
          }
        }
        dsheet.autoResizeColumn(col);
        var autowidth = dsheet.getColumnWidth(col);
        dsheet.setColumnWidth(col, (hdr.key == 'title' && autowidth > 200) ? 200 : 10 + autowidth);
      }
    }
  } catch (e) {
    Logger.log(e);
    throw e; // Bubble to UI
  }
  return;
}

function setDueDates() {
  try {
    var courseId = getCourseId();
    if (!courseId) {
      throw new Error('You must specify the course ID before you can change the dates.');
    }
    var dsheet = getDataSheet();
    dsheet.setActiveRange(dsheet.getRange(1, 1));
    var range = dsheet.getDataRange();
    var rows = range.getValues();
    var hdrs = getHeaders(rows[0]);
    if (hdrs.id.c2 < 0) {
      throw new Error('You do not have Canvas IDs in here and without those, I cannot do anything.');
    }
    if (hdrs.type.c2 < 0) {
      throw new Error('You do not have a column specifying whether this is a quiz or assignment. I need that to know how to process the information.');
    }
    
    var existingData = getDueDates(courseId);
    var changes = {};
    var updateCount = 0; // Tracks successful API requests
    
    for (var i = 1, rowCount = rows.length; i < rowCount; i++) {
      var row = rows[i];
      var canvasId = row[hdrs.id.c2];
      var type = row[hdrs.type.c2];
      if (type != 'Quiz' && type != 'Assignment') {
        throw new Error('The type must be Quiz or Assignment in row ' + (i + 1));
      }
      var itemKey = type.substr(0, 1).toLowerCase() + canvasId;
      if (typeof existingData[itemKey] === 'undefined') {
        throw new Error('You are trying to replace a quiz/assignment that is not in the system. Look for Canvas Id: ' + canvasId + ' in row ' + (i + 1) + ' of the spreadsheet.');
      }
      // We have a match, now it's time to look for any changes between the original and the new
      var existing = existingData[itemKey];
      for ( var key in hdrs) {
        if (hdrs.hasOwnProperty(key)) {
          var hdr = hdrs[key];
          if (hdr.c2 < 0 || hdr.required) {
            continue;
          }
          var value = row[hdr.c2];
          switch (hdr.type) {
            case 'date':
              value = toIso8601(value, hdr.daysEnd);
              break;
            case 'boolean':
              value = !(value == 0);
              break;
          }
          var field = hdr.field;
          var code = itemKey;
          // Points are displayed for reference only.
// Canvas classic quiz point updates are not supported.
if (field === 'points_possible') {
  continue;
}

          if (hdrs[hdr.key].location) {
            if (
    type == 'Quiz' &&
    hdrs[hdr.key].location == 'Assignment' &&
    existingData[itemKey].assignment_id
) {

    code =
        'a' +
        existingData[itemKey].assignment_id;

    
}
          }
          if (field == 'title' && type == 'Assignment') {
            field = 'name';
          }
          if (value == '' && (existing[field] == null || existing[field] == '')) {
            continue;
          }
          
          if (value == existing[field]) {
            continue;
          }
          if (hdr.type === 'date' && secondResolution(value, existing[field])) {
            continue;
          }
          if (typeof changes[code] === 'undefined') {
            changes[code] = {};
          }
          changes[code][field] = value;
          
        }
      }
    }
    
    var item;
    var result;
    var failedItems = []; // Array to track items that fail to update

    for ( var changeKey in changes) {
      if (changes.hasOwnProperty(changeKey)) {
        var change = changes[changeKey];
        var ltype = changeKey.substr(0, 1);
        var prefix = ltype == 'q' ? 'quiz' : 'assignment';
        var id = changeKey.substr(1);
        item = {
          ':course_id' : courseId,
          ':id' : id,
        };
        for ( var key2 in change) {
          if (change.hasOwnProperty(key2)) {
            var prefixedKey = prefix + '[' + key2 + ']';
            item[prefixedKey] = change[key2];
          }
        }
        
        // Wrap the API call in a try/catch so one failure doesn't kill the batch
        try {
          switch (ltype) {
            case 'a':
            
              result = canvasAPI(
    'PUT /api/v1/courses/:course_id/assignments/:id',
    item
);


              
              
              break;
            case 'q':
              result = canvasAPI('PUT /api/v1/courses/:course_id/quizzes/:id', item);
              break;
          }
          updateCount++;
        } catch (apiError) {
          Logger.log('Failed to update Canvas ID ' + id + ': ' + apiError.toString());
          failedItems.push(id);
        }
      }
    }
    
    // Provide a highly accurate summary to the user
    if (updateCount === 0 && failedItems.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Changes Detected', 
        'No due date changes were found to update.', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else if (updateCount === 0 && failedItems.length > 0) {
      SpreadsheetApp.getUi().alert(
        'Update Failed', 
        '0 Canvas items updated.\n\nThe following Canvas IDs failed and require manual review: ' + failedItems.join(', '), 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else if (failedItems.length > 0) {
      SpreadsheetApp.getUi().alert(
        'Partial Success', 
        updateCount + ' Canvas item(s) successfully updated.\n\nHowever, the following Canvas IDs failed to update and require manual review: ' + failedItems.join(', '), 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        'Success!', 
        updateCount + ' Canvas item(s) successfully updated.', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return;
    
  } catch (e) {
    Logger.log(e);
    if (typeof showError === 'function') showError('Failed to Save Due Dates', e);
    return;
  }
}

function secondResolution(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') {
    return false;
  }
  return (a.substr(0, 19) === b.substr(0, 19));
}

function helpDialog() {
  // Removed the deprecated .setSandboxMode() method
  var html = HtmlService.createTemplateFromFile('help').evaluate();
  var props = SpreadsheetApp.getUi().showModalDialog(html,'Course Due Dates');
  return;
}

function shiftAllDatesDialog() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Shift All Dates',
    'Enter the number of days to shift Due, Available From, Available Until, ' +
'and Show/Hide Answers dates by.\n' +
'Use a negative number to shift earlier (e.g. -14).\n\n' +
'Decimal values are allowed.\n\n' +
'Examples:\n' +
'1 = 1 day\n' +
'0.5 = 12 hours\n' +
'0.25 = 6 hours\n' +
'0.041667 = 1 hour\n' +
'-0.25 = 6 hours earlier\n\n' +
'This only changes the spreadsheet; you still need to choose Save Due Dates afterward.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  var text = response.getResponseText().trim();
  if (!/^-?\d+(\.\d+)?$/.test(text)) {
    ui.alert('Invalid Input', 'Please enter a number of days.\n\nExamples:\n7 = 7 days\n0.5 = 12 hours\n-0.25 = 6 hours earlier', ui.ButtonSet.OK);
    return;
  }
  shiftAllDates(parseFloat(text));
}

function shiftAllDates(days) {
  try {
    var dsheet = getDataSheet();
    if (dsheet.getRange(1, 1).isBlank()) {
      throw new Error('No data to process. Load due dates first.');
    }
    var range = dsheet.getDataRange();
    var rowCount = range.getNumRows() - 1;
    var rows = range.getValues();
    var hdrs = getHeaders(rows[0]);
    var dateCols = [];
    
    for (var key in hdrs) {
      if (hdrs.hasOwnProperty(key) && hdrs[key].type === 'date' && hdrs[key].c2 >= 0) {
        dateCols.push(hdrs[key].c2);
      }
    }
    
    if (dateCols.length === 0) {
      throw new Error('No date columns found in the sheet.');
    }
    
    var msInDay = 24 * 60 * 60 * 1000;
    var changed = 0;
    
    // Step 1: Count how many date values will be shifted
    for (var i = 1; i <= rowCount; i++) {
      for (var c = 0; c < dateCols.length; c++) {
        var col = dateCols[c];
        var val = rows[i][col];
        if (Object.prototype.toString.call(val) === '[object Date]') {
          changed++;
        }
      }
    }
    
    // Safety check: Exit early if nothing will be updated
    if (changed === 0) {
      SpreadsheetApp.getUi().alert(
        'No Dates Found', 
        'There are no valid dates to shift.', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    var ui = SpreadsheetApp.getUi();

    // Step 2: Ask for confirmation BEFORE modifying the array
   

var displayDays = Number(Math.abs(days).toFixed(2));
var direction = days >= 0 ? 'later' : 'earlier';
var response = ui.alert(
  'Confirm Date Shift',
  'You are about to shift ' + changed + ' date value(s) ' + displayDays + ' day(s) ' + direction + '.\n\n' +
  'This will modify the spreadsheet immediately.\n' +
  'You can still review the changes before saving them to Canvas.\n\n' +
  'Continue?',
  ui.ButtonSet.YES_NO
);

    if (response !== ui.Button.YES) {
      return;
    }
    
    // Step 3: The user confirmed. Perform the math and update the array.
    for (var j = 1; j <= rowCount; j++) {
      for (var d = 0; d < dateCols.length; d++) {
        var execCol = dateCols[d];
        var execVal = rows[j][execCol];
        if (Object.prototype.toString.call(execVal) === '[object Date]') {
          rows[j][execCol] = new Date(execVal.getTime() + days * msInDay);
        }
      }
    }

    // Step 4: Write the updated array back to the spreadsheet
    range.setValues(rows);
    ui.alert(
      'Dates Shifted',
      'Shifted by' + days + ' day(s) across ' + changed + ' date value(s).\n\n' +
      'Review the sheet, then choose Save Due Dates to push the changes to Canvas.',
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    Logger.log(e);
    if (typeof showError === 'function') showError('Shift Dates Error', e);
  }
}
function remapDateRangeDialog() {
  var html = HtmlService.createHtmlOutputFromFile('remapDates')
    .setWidth(320)
    .setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, 'Remap Date Range');
}

function processRemapForm(formObject) {
  // Parse the native datetime-local string directly
  var oldDate = new Date(formObject.old_date);
  var newDate = new Date(formObject.new_date);
  if (isNaN(oldDate) || isNaN(newDate)) {
    throw new Error('Could not understand one of the dates entered.');
  }
  var msInDay = 24 * 60 * 60 * 1000;
  // Use precise decimal days to allow hour/minute shifting (Do not round)
  var deltaDays = (newDate.getTime() - oldDate.getTime()) / msInDay;
  shiftAllDates(deltaDays);
}

function availabilityDefaultsDialog() {
  var userProperties = PropertiesService.getUserProperties();
  var template = HtmlService.createTemplateFromFile('availabilityDefaults');
  template.fromOffset = userProperties.getProperty('availFromOffset') || '0';
  template.fromTime = userProperties.getProperty('availFromTime') || '';
  template.untilOffset = userProperties.getProperty('availUntilOffset') || '';
  template.untilTime = userProperties.getProperty('availUntilTime') || '';
  
  var html = template.evaluate().setWidth(420).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Default Availability Dates');
}

function processAvailabilityDefaultsForm(formObject) {
  var fromOffset = formObject.from_offset === '' ? null : Number(formObject.from_offset);
  var untilOffset = formObject.until_offset === '' ? null : Number(formObject.until_offset);
  var fromTime = formObject.from_time;
  var untilTime = formObject.until_time;
  if (fromOffset < 0) {
  throw new Error(
    'Available From days must be positive.'
  );
}
  if (fromOffset !== null && isNaN(fromOffset)) {
    throw new Error('Available From offset must be a number.');
  }
  if (untilOffsetOffset < 0) {
  throw new Error(
    'Available Until days must be positive.'
  );
}
  if (untilOffset !== null && isNaN(untilOffset)) {
    throw new Error('Available Until offset must be a number.');
  }
  
  var userProperties = PropertiesService.getUserProperties();
  if (fromOffset !== null) userProperties.setProperty('availFromOffset', String(fromOffset));
  if (untilOffset !== null) userProperties.setProperty('availUntilOffset', String(untilOffset));
  userProperties.setProperty('availFromTime', fromTime);
  userProperties.setProperty('availUntilTime', untilTime);
  
  var overwrite = (formObject.overwrite === true || formObject.overwrite === 'true');
  applyAvailabilityDefaults(fromOffset, fromTime, untilOffset, untilTime, overwrite);
}

function applyAvailabilityDefaults(fromOffsetDays, fromTime, untilOffsetDays, untilTime, overwrite) {
  try {
    var dsheet = getDataSheet();
    if (dsheet.getRange(1, 1).isBlank()) {
      throw new Error('No data to process. Load due dates first.');
    }
    var range = dsheet.getDataRange();
    var rowCount = range.getNumRows() - 1;
    var rows = range.getValues();
    var hdrs = getHeaders(rows[0]);
    if (hdrs.due.c2 < 0) {
      throw new Error('No Due column found in the sheet.');
    }
    var msInDay = 24 * 60 * 60 * 1000;
    var filledFrom = 0;
    var filledUntil = 0;
    
    for (var i = 1; i <= rowCount; i++) {
      var dueVal = rows[i][hdrs.due.c2];
      if (Object.prototype.toString.call(dueVal) !== '[object Date]') {
        continue;
      }
      
      // Process Available From
      if (hdrs.unlock && hdrs.unlock.c2 >= 0 && fromOffsetDays !== null) {
        var curFrom = rows[i][hdrs.unlock.c2];
        var isBlankFrom = Object.prototype.toString.call(curFrom) !== '[object Date]';
        if (overwrite || isBlankFrom) {
          var newFrom = new Date(dueVal.getTime() - (fromOffsetDays * msInDay));
          if (fromTime) {
            var fParts = fromTime.split(':');
            newFrom.setHours(parseInt(fParts[0], 10), parseInt(fParts[1], 10), 0, 0);
          }
          rows[i][hdrs.unlock.c2] = newFrom;
          filledFrom++;
        }
      }
      
      // Process Available Until
      if (hdrs.lock && hdrs.lock.c2 >= 0 && untilOffsetDays !== null) {
        var curUntil = rows[i][hdrs.lock.c2];
        var isBlankUntil = Object.prototype.toString.call(curUntil) !== '[object Date]';
        if (overwrite || isBlankUntil) {
          var newUntil = new Date(dueVal.getTime() + (untilOffsetDays * msInDay));
          if (untilTime) {
            var uParts = untilTime.split(':');
            newUntil.setHours(parseInt(uParts[0], 10), parseInt(uParts[1], 10), 0, 0);
          }
          rows[i][hdrs.lock.c2] = newUntil;
          filledUntil++;
        }
      }
    }
    
    range.setValues(rows);
    var msg = 'Available From set on ' + filledFrom + ' row(s).\n' +
              'Available Until set on ' + filledUntil + ' row(s).\n\n' +
              'Review the sheet, then choose Save Due Dates to push the changes to Canvas.';
    SpreadsheetApp.getUi().alert(
      overwrite ? 'Availability Reset' : 'Availability Defaults Applied',
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log(e);
    if (typeof showError === 'function') showError('Availability Defaults Error', e);
  }
}