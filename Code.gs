function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Formulas & Data")
    .addItem("Error Wrap", "errorWrap")
    .addSeparator()
    .addItem("Flip Sign", "flipSign")
    .addSeparator()
    .addItem("Comment Cells", "commentCells")
    .addSeparator()
    .addItem("Clean Cells", "cleanCells")
    .addSeparator()
    .addItem("Anchor Formula", "anchorFormulas")
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

function commentCells() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var selectedRange = sheet.getActiveRange();
  var cells = selectedRange.getValues();

  for (var i = 0; i < cells.length; i++) {
    for (var j = 0; j < cells[i].length; j++) {
      var cell = cells[i][j];
      if (typeof cell === 'number' || cell.toString().startsWith('=')) {
        var range = sheet.getRange(selectedRange.getRowIndex() + i, selectedRange.getColumnIndex() + j);
        var comment = range.getComment();
        if (comment == null) {
          range.setValue("'" + cell.toString());
        } else {
          range.clearComment();
          range.setValue(cell);
        }
      }
    }
  }
}

function cleanCells() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var selectedRange = sheet.getActiveRange();
  var cells = selectedRange.getValues();

  for (var i = 0; i < cells.length; i++) {
    for (var j = 0; j < cells[i].length; j++) {
      var cell = cells[i][j];
      if (typeof cell === 'string') {
        // Trim extraneous spaces
        cell = cell.replace(/\s{2,}/g, ' ').trim();
        // Remove worksheet names from formulas
        var sheetName = sheet.getName();
        if (cell.toString().startsWith('=' + sheetName)) {
          cell = cell.replace('=' + sheetName, '=');
        }
        sheet.getRange(selectedRange.getRowIndex() + i, selectedRange.getColumnIndex() + j).setValue(cell);
      }
    }
  }
}

function anchorFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var selectedRange = sheet.getActiveRange();
  var formulas = selectedRange.getFormulas();
  
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      var formula = formulas[i][j];
      var newFormula = formula;
      if (formula) {
        var cell = sheet.getRange(selectedRange.getRowIndex() + i, selectedRange.getColumnIndex() + j);
        var anchors = getAnchors(formula);
        newFormula = anchorFormula(formula, anchors);
        cell.setFormula(newFormula);
      }
    }
  }
}

function getAnchors(formula) {
  var anchors = [];
  var regex = /\$?[A-Z]+\$?\d+/g;
  var match;
  while ((match = regex.exec(formula)) != null) {
    anchors.push(match[0]);
  }
  return anchors;
}

function anchorFormula(formula, anchors) {
  for (var i = 0; i < anchors.length; i++) {
    var anchor = anchors[i];
    var newAnchor = anchor;
    var regex = new RegExp("(?<=\\W)" + anchor + "(?=\\W)", "g");
    if (!/\$/.test(anchor)) {
      newAnchor = "$" + newAnchor.replace(/\d+/, "$&");
    }
    if (!/\$/.test(newAnchor.substring(0, newAnchor.indexOf("$")))) {
      newAnchor = newAnchor.replace(/[A-Z]+/, "$&");
    }
    formula = formula.replace(regex, newAnchor);
  }
  return formula;
}
