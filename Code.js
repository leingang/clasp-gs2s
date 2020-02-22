var APP_NAME="GS2S"

const GS_SHEET_NAME='scores (from Gradescope)';
const GS_FIRSTNAME_COLUMN=1;
const GS_FIRSTNAME_COLUMN_NAME="First Name";
const GS_LASTNAME_COLUMN=2;
const GS_LASTNAME_COLUMN_NAME="Last Name";
const GS_SID_COLUMN=3;
const GS_SID_COLUMN_NAME="SID";
const GS_SCORE_COLUMN=6;
const GS_SCORE_COLUMN_NAME="Total Score";
const SK_SHEET_NAME='To Sakai';
const SK_SID_COLUMN=1;
const SK_SID_COLUMN_NAME='Student ID';
const SK_NAME_COLUMN=2;
const SK_NAME_COLUMN_NAME='Student Name';
const SK_SCORE_COLUMN=3;
const SK_COMMENT_COLUMN=4;
const SK_MSG_KUDOS="Good job!";

function onOpen() {
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .createMenu(APP_NAME)
        .addItem('Set Late Policies...','setLatePolicies')
        .addItem('Rescale...','rescale')
        .addToUi();
  }

  function setLatePolicies() {
      var ui = SpreadsheetApp.getUi();
      ui.alert("Not implemented yet");
  }


/**
 * Read assignment data from a Gradescope spreadsheet
 *
 * The return value has properties "name", "maxPoints", and "records".  
 * The "records" property is an array of objects.  Each one has properties
 * "firstName", "lastName", "sid", and "totalPoints".
 *
 * @param {SpreadSheet} sheet 
 * @return {Object}
 */
function readGradescopeSheet(sheet) {
    if (! isGradescopeSheet(sheet)) {
        throw new Error("not a Gradescope score sheet");
    }
    var data = sheet.getDataRange().getValues();
    var assignment = {};
    assignment.name = nameFromGradescopeSheet(sheet);
    var headers = data.shift();
    assignment.maxPoints = sumPoints(headers);
    assignment.records = [];
    for (const row of data) {
        // Logger.log("readData:row: " + row);
        var record = {};
        row.unshift(""); // spreadsheet column numbers start with 1
        record.firstName = row[GS_FIRSTNAME_COLUMN];
        record.lastName = row[GS_LASTNAME_COLUMN];
        record.sid = row[GS_SID_COLUMN];
        record.totalPoints = row[GS_SCORE_COLUMN];
        // Logger.log("readData:record: " + record);
        assignment.records.push(record);
    }
    return assignment;
}

/**
 * Decide if a sheet “looks like” a Gradescope score sheet.
 *
 * Checks the header row and matches the fields with expected ones.
 *
 * @param {SpreadSheet} sheet 
 * @returns {boolean}
 */
function isGradescopeSheet(sheet) {
    var data = sheet.getDataRange().getValues();
    var header = data.shift();
    header.unshift("");
    return (header[GS_FIRSTNAME_COLUMN] == GS_FIRSTNAME_COLUMN_NAME)
        && (header[GS_LASTNAME_COLUMN] == GS_LASTNAME_COLUMN_NAME)
        && (header[GS_SID_COLUMN] == GS_SID_COLUMN_NAME)
        && (header[GS_SCORE_COLUMN] == GS_SCORE_COLUMN_NAME);
}

/**
 * Parse the name of the assignment from a Gradescope score sheet.
 *
 * For example, when exporting an assignemnt named "HW 02", the exported file is
 * named "HW_02_scores.csv".  When imported to Google Spreadheets, the extension
 * is cut off.  This function will strip off "_scores.csv", replace the
 * underscore, and return "HW 02".
 *
 * @param {*} sheet 
 * @returns {string} the name
 */
function nameFromGradescopeSheet(sheet) {
    var regex = /(.*)_scores.*/;
    var sheetName = sheet.getName();
    var name = "Assignment"; // default value
    if (m = regex.exec(sheetName)) {
        name = m[1].replace("_"," ");
    }
    return name;
}


/**
 * Scan field names and sum up point values.
 * @param {array} fields 
 */
function sumPoints(fields) {
    var regex = /\((\d+\.\d+) pts\)/;
    var tot=0;

    for (const field of fields) {
        if (m = regex.exec(field)) {
            tot += Number(m[1]);
        }
    }
    return tot;
}

/**
 * process an assignment, rescaling the total
 * 
 * @param {Assignment} assignment 
 */
function process(assignment) {
    var newMaxPoints = 100;
    for (i in assignment.records) {
        record = assignment.records[i];
        record.formattedName = record.lastName + ", " + record.firstName;
        record.comment = "";
        if (record.totalPoints == assignment.maxPoints) {
            record.comment += SK_MSG_KUDOS;
        }
        record.totalPoints *= newMaxPoints/assignment.maxPoints;
    }
    assignment.maxPoints = newMaxPoints;
}

/**
 * Get a new sheet with name "name", deleting any existing sheet
 * 
 * @param {SpreadSheet} spreadSheet
 * @param {string} name 
 * @returns {Sheet}
 */
function getCleanSheet(spreadSheet, name) {
    var sheet = spreadSheet.getSheetByName(name);
    if (sheet) {
        spreadSheet.deleteSheet(sheet);
    }
    return spreadSheet.insertSheet(name);
}


function writeSakaiSheet(sheet,assignment) {
    var headers = [
        SK_SID_COLUMN_NAME,
        SK_NAME_COLUMN_NAME,
        assignment.name + " [" + assignment.maxPoints + "]",
        "* " + assignment.name
    ];
    var rows = [];
    rows.push(headers);
    for (const record of assignment.records) {
        rows.push(["'" + record.sid,record.formattedName,record.totalPoints,record.comment]);
    }
    sheet.getRange(1,1,rows.length,rows[0].length).setValues(rows)
}


/**
 * rescale the values in the Gradescope report
 */
function rescale() {

    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var gs = SpreadsheetApp.getActiveSheet();

    assignment = readGradescopeSheet(gs);
    process(assignment);

    var sk = getCleanSheet(ss,assignment.name + " (rescaled, for Sakai)");
    writeSakaiSheet(sk,assignment);

}

