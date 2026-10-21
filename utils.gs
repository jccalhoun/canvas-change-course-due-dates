/**
* @OnlyCurrentDoc
*/
var _cachedTz = null;

function assertTimeZonesMatch() {
  var sheetTz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var scriptTz = Session.getScriptTimeZone();
  if (sheetTz !== scriptTz) {
    throw new Error('Time zone mismatch: the spreadsheet uses ' + sheetTz + 
      ' but the script uses ' + scriptTz + 
      '. Make them match (File > Settings, or appsscript.json) before working with dates.');
  }
}

/**
* @function Add days by calendar date so wall-clock times survive DST changes.
*           Fractional days are added as exact time after the whole days.
*/
function addDays(date, days) {
  var wholeDays = days < 0 ? Math.ceil(days) : Math.floor(days);
  var remainderMs = Math.round((days - wholeDays) * 86400000 / 60000) * 60000;
  var d = new Date(date.getTime());
  d.setDate(d.getDate() + wholeDays);
  d.setTime(d.getTime() + remainderMs);
  return d;
}

/**
* @function Display an error in the UI
* @param {String} title - The title of the error dialog
* @param {Error|String} err - The error object or message
*/
function showError(title, err) {
  SpreadsheetApp.getUi().alert(
    title, 
    err.toString(), 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
* @function Converts an ISO8601 timestamp into a Google time
* @params {String} timestamp - ISO8601 timestamp from Canvas
* @returns {String} Formatted Google date/time
*/
function fromIso8601(timestamp, asDateObject) {
  if (typeof timestamp === 'undefined' || timestamp == null) {
    return false;
  }
  var regex = /(:\d\d)Z$/;
  timestamp = timestamp.replace(regex, '$1.000Z');
  var dt = new Date(timestamp);
  if (isNaN(dt)) {
    return false;
  }
  if (typeof asDateObject !== 'undefined' && asDateObject) {
    return dt;
  }
  if (!_cachedTz) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    _cachedTz = ss.getSpreadsheetTimeZone();
  }
  return Utilities.formatDate(dt, _cachedTz, 'yyyy-MM-dd HH:mm:ss');
}

/**
* @function Convert a Google date/time object to an ISO8601 timestamp
* @param {Object} datetime - Date object 
* @returns {String} ISO8601 timestamp
*/
function toIso8601(datetime, daysEnd) {
  if (!(datetime instanceof Date)) {
    return '';
  }
  
  // Clone the date object to prevent mutating the original in-memory array
  var dt = new Date(datetime.getTime());
  
  var hr;
  var min;
  
  if (typeof daysEnd !== 'undefined' && daysEnd) {
    hr = dt.getHours();
    min = dt.getMinutes();
    if (hr === 0 && min === 0) {
      hr = 23;
      min = 59;
    }
    if (min === 59) {
      dt.setHours(hr, min, 59);
    }
    else {
      dt.setHours(hr, min, 0);
    }
  }
  return Utilities.formatDate(dt, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

/**
* @function Obtain the course ID through the UI and save it into the user properties
*/
function getCourseDialog() {
  var courseId = null;
  try {
    var ui = SpreadsheetApp.getUi();
    var settings = getApiSettings();
    if (settings === false) {
      return false;
    }
    var userProperties = PropertiesService.getUserProperties();
    var msg = 'Specify the Canvas Course ID for your course.\n\nYou may enter this as an integer or you may paste a URL from your course into the box.\nThe URL should look like this: https://' + settings.host + '/courses/123\n \n';
    var btnset = ui.ButtonSet.OK;
    var savedCourseId = userProperties.getProperty('courseid');
    if (savedCourseId) {
      savedCourseId = savedCourseId.trim();
      if (/^[0-9]+$/.test(savedCourseId)) {
        var savedCourse;
        // Safely attempt to fetch the saved course. 
        // If it was reset/deleted in Canvas, catch the 404 so the script doesn't crash.
        try {
          savedCourse = canvasAPI('GET /api/v1/courses/' + savedCourseId );
        } catch (oldCourseError) {
          userProperties.deleteProperty('courseid');
        }
        
        if (typeof savedCourse !== 'undefined') {
          msg += 'Press Cancel to keep the current information.\nCourse ID : ' + savedCourse.id + '\nCourse name: ' + savedCourse.name + '\n \n';
          btnset = ui.ButtonSet.OK_CANCEL;
        }
      }
      else {
        userProperties.deleteProperty('courseid');
      }
    }
    var response = ui.prompt('Specify Course ID', msg, btnset);
    if (response.getSelectedButton() == ui.Button.OK) {
      var txt = response.getResponseText();
      txt = txt.trim();
      var courseRegex = new RegExp('^https://.*/courses/([0-9]+)','i');
      var match = courseRegex.exec(txt);
      if (match) {
        courseId = match[1];
      }
      else if (/^[0-9]+$/.test(txt)) {
        courseId = txt;
      }
      if (typeof courseId !== 'undefined' && courseId) {
        var course;
        // Fix: Ensure we actually successfully fetched the savedCourse before reusing it!
        if (savedCourse && courseId == savedCourseId) {
          course = savedCourse;
        } else {
          course = canvasAPI('GET /api/v1/courses/' + courseId );
        }
        // Because canvasAPI throws on error, if we reach this line, course is definitely valid.
        var proceed = ui.alert(course.name,'Course ID : ' + course.id + '\nCourse name: ' + course.name + '\n\nWould you like to proceed?',ui.ButtonSet.YES_NO);
        if (proceed == ui.Button.YES) {
          userProperties.setProperty('courseid', courseId);
          userProperties.setProperty('coursename', course.name);
        } else {
          courseId = null;
        }
      }
      else {
        ui.alert('Bad Input','I was unable to determine your Canvas Course ID.',ui.ButtonSet.OK);
        courseId = null;
      }
    }
  } catch(e) {
    Logger.log(e);
    if (typeof showError === 'function') {
      showError('Course Dialog Error', e);
    }
    return;
  }
  return courseId;
}