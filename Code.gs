function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Formulas")
    .addItem("Error Wrap", "errorWrap")
    .addSeparator()
    .addItem("Flip Sign", "flipSign")
    .addToUi();
}

function errorWrap() {
  var cell = SpreadsheetApp.getActiveRange().getCell(1, 1);
  var formula = cell.getFormula();

  // check if the formula starts with '='
  var startsWithEqualSign = formula.startsWith('=');
  
  // remove the equal sign if it starts with '='
  if (startsWithEqualSign) {
    formula = formula.substring(1);
  }

  if (!/^iferror/i.test(formula)) {
    cell.setFormula('=IFERROR(' + formula + ', "NA")');
  } else {
    cell.setFormula(formula.replace(/^iferror\((.*),.*\)/i, '=$1'));
  }

}

function flipSign() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (typeof values[i][j] === "number") {
        values[i][j] *= -1;
      } else if (typeof values[i][j] === "string") {
        if (values[i][j].charAt(0) === "=") {
          values[i][j] = "=" + values[i][j].substring(1).replace(/(\d+)/g, "-$1");
        }
      } else if (values[i][j] instanceof Array) {
        // This code is for array formulas.
        values[i][j].forEach(function(element, index) {
          if (typeof element === "number") {
            values[i][j][index] *= -1;
          } else if (typeof element === "string" && element.charAt(0) === "=") {
            values[i][j][index] = "=" + element.substring(1).replace(/(\d+)/g, "-$1");
          }
        });
      }
    }
  }
  
  range.setValues(values);
}
