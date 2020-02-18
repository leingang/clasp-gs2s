var APP_NAME="GS2S"

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

  function rescale() {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Not implemented yet");
}

