var APP_NAME="GS2S"

const GS_SHEET_NAME='scores (from Gradescope)';
const GS_FIRSTNAME_COLUMN=1;
const GS_LASTNAME_COLUMN=2;
const GS_SID_COLUMN=3;
const GS_SCORE_COLUMN=6;
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
   * Get the maximum point value from the Gradescope assignment sheet
   * 
   * @param {Sheet} sheet 
   * @returns number
   */
  function getMaxPoints(sheet) {
    var ui = SpreadsheetApp.getUi();
    var j=1; 
    var tot=0;
    var regex = /\((\d+\.\d+) pts\)/;

    while (!sheet.getRange(1,++j).isBlank()) {
     if(m = regex.exec(sheet.getRange(1,j).getValue())) {
            tot += Number(m[1]);
        }
    }
    return tot;
}

  /**
   * rescale the values in the Gradescope report
   */
  function rescale() {
    // TODO: get from a dialog
    var assignmentName = 'Assignment';
    var newMaxPoints = 100;

    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var gs = ss.getSheetByName(GS_SHEET_NAME);
    var oldMaxPoints = getMaxPoints(gs);
    var oldScore,newScore;

    var s=ss.getSheetByName(SK_SHEET_NAME);
    if (s) {
        ss.deleteSheet(s);
    }
    s = ss.insertSheet(SK_SHEET_NAME);

    // header row
    s.getRange(1,SK_SID_COLUMN).setValue(SK_SID_COLUMN_NAME);
    s.getRange(1,SK_NAME_COLUMN).setValue(SK_NAME_COLUMN_NAME);
    s.getRange(1,SK_SCORE_COLUMN).setValue(assignmentName + " [" + newMaxPoints + "]");
    s.getRange(1,SK_COMMENT_COLUMN).setValue(" * " + assignmentName);

    var i = 1; // row index
    while (!gs.getRange(++i,1).isBlank()) {
        // first column is the NetID (from column 3)
        gs.getRange(i,GS_SID_COLUMN).copyTo(s.getRange(i,SK_SID_COLUMN));
        // next column is "lastname, firstname"
        s.getRange(i,SK_NAME_COLUMN).setValue(
            gs.getRange(i,GS_LASTNAME_COLUMN).getValue()
            + ", "
            + gs.getRange(i,GS_FIRSTNAME_COLUMN).getValue()
        );
        // next column is score
        oldScore = gs.getRange(i,GS_SCORE_COLUMN).getValue();
        newScore = oldScore/oldMaxPoints*newMaxPoints;
        s.getRange(i,SK_SCORE_COLUMN).setValue(newScore);
        if (oldScore == oldMaxPoints) {
            s.getRange(i,SK_COMMENT_COLUMN).setValue(SK_MSG_KUDOS);
        }

        
    }

    


}

