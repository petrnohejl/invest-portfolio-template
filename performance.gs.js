/* Documentation: https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app */

/* settings */
const startDate = "DATE(2015,04,01)";
const benchmark = "IVV";
const benchmarkId = -1;

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Invest Portfolio')
    .addItem('Generate performance sheet', 'generatePerformanceSheet')
    .addToUi();
}

function generatePerformanceSheet() {
  /* reference to the sheets */
  let spreadsheet = SpreadsheetApp.getActive();
  let stocksSheet = spreadsheet.getSheetByName("Stocks");
  let stocksExitsSheet = spreadsheet.getSheetByName("Stocks Exits");
  let performanceSheet = createSheet(spreadsheet, "Stocks Performance");

  /* get sheets values */
  let stocksValues = stocksSheet.getDataRange().getValues();
  let stocksExitsValues = stocksExitsSheet.getDataRange().getValues();

  /* arrays */
  let tickers = [];
  let shares = [];
  let purchases = [];
  let exits = [];
  let prices = [];

  /* headers */
  tickers.push("Date");
  shares.push("Shares");
  purchases.push("Purchase");
  exits.push("Exit");

  /* date column */
  prices.push('=QUERY(GoogleFinance("' + benchmark + '", "price", ' + startDate + ', TODAY(), "daily"), "select Col1 label Col1\'\'")');

  /* stocks historical prices for each ticker */
  for (var i = 1; i <= stocksValues.length; i++) {
    if (stocksValues[i][0] == "") break;
    tickers.push(stocksValues[i][0]);
    shares.push(stocksValues[i][3]);
    purchases.push(stocksValues[i][8]);
    exits.push(0);
    prices.push('=QUERY(GoogleFinance(' + columnToLetter(i + 1) + '1, "price", ' + startDate + ', TODAY(), "daily"), "select Col2 label Col2\'\'")');
  }

  /* stocks exits historical prices for each ticker */
  let stocksValuesCount = tickers.length;
  for (var i = 1; i <= stocksExitsValues.length; i++) {
    if (stocksExitsValues[i][0] == "") break;
    tickers.push(stocksExitsValues[i][0]);
    shares.push(stocksExitsValues[i][5]);
    purchases.push(stocksExitsValues[i][10]);
    exits.push(stocksExitsValues[i][14]);
    prices.push('=QUERY(GoogleFinance(' + columnToLetter(i + stocksValuesCount) + '1, "price", ' + startDate + ', TODAY(), "daily"), "select Col2 label Col2\'\'")');
  }

  /* add benchmark */
  tickers.push("Benchmark");
  shares.push(1);
  purchases.push(benchmarkId);
  exits.push(0);
  prices.push('=QUERY(GoogleFinance("' + benchmark + '", "price", ' + startDate + ', TODAY(), "daily"), "select Col2 label Col2\'\'")');

  /* logs */
  Logger.log(tickers);
  Logger.log(shares);
  Logger.log(purchases);
  Logger.log(exits);

  /* write Google Finance data to the performance sheet */
  let performanceValues = [];
  performanceValues.push(tickers);
  performanceValues.push(prices);
  performanceSheet.getRange(1, 1, performanceValues.length, performanceValues[0].length).setValues(performanceValues);

  /* convert Google Finance formulas to values */
  for (var i = 1; i <= tickers.length; i++) {
    SpreadsheetApp.flush();
    formulasToValues(performanceSheet, i);
  }

  /* calculate gain */
  for (var i = 1; i < tickers.length; i++) {
    let column = i + 1;
    let holdingRange = getFullColumn(performanceSheet, columnToLetter(column), 2);
    let purchaseDate = null;
    let exitDate = null;

    if (i < stocksValuesCount) {
      purchaseDate = "Stocks!C" + column;
    } else {
      purchaseDate = "'Stocks Exits'!C" + (column - stocksValuesCount + 1);
      exitDate = "'Stocks Exits'!D" + (column - stocksValuesCount + 1);
    }

    calculateGain(holdingRange, shares[i], purchases[i], exits[i], purchaseDate, exitDate);
  }

  /* calculate invested */
  let investedValues = [];
  let holdingValues = performanceSheet.getRange("B2:" + columnToLetter(performanceSheet.getLastColumn() - 1) + performanceSheet.getLastRow()).getValues();
  for (var i = 0; i < holdingValues.length; i++) {
    var invested = 0;
    for (var j = 0; j < holdingValues[i].length; j++) {
      if (holdingValues[i][j] != 0) {
        var exitGain = exits[j + 1] - purchases[j + 1]; /* add +1 to skip the Purchase header */
        if (Math.round(parseFloat(holdingValues[i][j]) * 1000) != Math.round(parseFloat(exitGain) * 1000)) {
          invested += purchases[j + 1];
        }
      }
    }
    investedValues.push([invested]);
  }

  /* add invested column */
  let investedColumn = columnToLetter(performanceSheet.getLastColumn() + 1);
  performanceSheet.getRange(investedColumn + "1").setValue("Invested");
  performanceSheet.getRange(investedColumn + "2:" + investedColumn + performanceSheet.getLastRow()).setValues(investedValues);

  /* add dividend and sum columns */
  let lastCol = tickers.length + 1;
  addColumnWithFormula(performanceSheet, lastCol + 1, "Gain", "=SUM(B2:" + columnToLetter(lastCol - 2) + "2)-$A$1");
  addColumnWithFormula(performanceSheet, lastCol + 2, "Return %", "=" + columnToLetter(lastCol + 1) + "2/" + columnToLetter(lastCol) + "2");
  addColumnWithFormula(performanceSheet, lastCol + 3, "Dividend", "=SUMIFS('Stocks Dividends'!$F:$F,'Stocks Dividends'!$A:$A,TO_DATE(DATEVALUE(A2)))");
  addColumnWithFormula(performanceSheet, lastCol + 4, "Dividend Acc", "=SUM(" + columnToLetter(lastCol + 3) + "$2:" + columnToLetter(lastCol + 3) + "2)");
  addColumnWithFormula(performanceSheet, lastCol + 5, "Gain Dividend", "=SUM(" + columnToLetter(lastCol + 1) + "2," + columnToLetter(lastCol + 4) + "2)");

  /* A1 cell for setting a gain init value */
  performanceSheet.getRange("A1").setValue("0.00");

  /* format the performance sheet */
  formatPerformanceSheet(performanceSheet);
}

function formulasToValues(sheet, column) {
  /* https://coderedirect.com/questions/701099/convert-formula-value-to-static-value-if-cell-is-not-blank */
  var range = sheet.getRange(1, column, sheet.getLastRow(), 1);
  var values = range.getValues();
  var formulas = range.getFormulas();

  /* fill zeros at the beginning if the stock history is too short */
  var rowsCount = values.length;
  values = values.filter(String);
  if (rowsCount > values.length) {
    var diff = rowsCount - values.length;
    for (var i = 0; i < diff; i++) {
      values.splice(1, 0, Array.of(0));
    }
  }

  var convertedValues = values.map(function (e, i) { return e[0] && formulas[i][0] || e[0] ? [e[0]] : [formulas[i][0]] });
  range.setValues(convertedValues);
}

function calculateGain(range, shares, purchase, exit, purchaseDate, exitDate) {
  let values = range.getValues();
  let benchmark = purchase == benchmarkId;
  if (benchmark) purchase = values[0][0]; /* purchase price for the benchmark */
  for (var i = 0; i < values.length; i++) {
    var cell = values[i][0];
    var rowDate = "A" + (i + 2);
    var gain = cell * shares - purchase;
    if (benchmark) {
      values[i][0] = gain / purchase; /* benchmark formula */
    } else {
      /* holding formula */
      if (exitDate == null) {
        values[i][0] = "=IF(" + purchaseDate + " <= " + rowDate + ", " + gain + ", 0)";
      } else {
        var exitGain = exit - purchase;
        values[i][0] = "=IF(" + purchaseDate + " <= " + rowDate + ", " + "IF(" + exitDate + " >= " + rowDate + ", " + gain + ", " + exitGain + ")" + ", 0)";
      }
    }
  }
  range.setValues(values);
}

function addColumnWithFormula(sheet, column, title, formula) {
  let rowsCount = sheet.getDataRange().getValues().length;
  let columnLetter = columnToLetter(column);
  sheet.getRange(columnLetter + "1").setValue(title);
  sheet.getRange(columnLetter + "2").setFormula(formula);
  sheet.getRange(columnLetter + "2").copyTo(sheet.getRange(columnLetter + "3:" + columnLetter + rowsCount));
}

function formatPerformanceSheet(sheet) {
  /* format the whole sheet */
  let performanceDataRange = sheet.getDataRange();
  performanceDataRange.setFontFamily("Arial");
  performanceDataRange.setFontSize(9);
  performanceDataRange.setBackground("#f3f3f3");

  /* header format */
  let headerRange = getFullRow(sheet, "A", 1);
  headerRange.setBackground("#c27ba0");
  headerRange.setFontColor("white");
  headerRange.setHorizontalAlignment("center");

  /* date format */
  let dateRange = getFullColumn(sheet, "A", 2);
  dateRange.setNumberFormat("yyyy-MM-dd");

  /* sum format */
  let sumColumns = 7;
  let sumRange = sheet.getRange(columnToLetter(sheet.getLastColumn() - sumColumns + 1) + 2 + ':' + columnToLetter(sheet.getLastColumn()) + sheet.getLastRow());
  sumRange.setBackground("#fff2cc");

  /* benchmark percent format */
  let benchmarkColumn = columnToLetter(sheet.getLastColumn() - sumColumns + 1);
  let benchmarkRange = sheet.getRange(benchmarkColumn + 2 + ':' + benchmarkColumn + sheet.getLastRow());
  benchmarkRange.setNumberFormat("0.00%");

  /* return percent format */
  let returnColumn = columnToLetter(sheet.getLastColumn() - sumColumns + 4);
  let returnRange = sheet.getRange(returnColumn + 2 + ':' + returnColumn + sheet.getLastRow());
  returnRange.setNumberFormat("0.00%");
}

function createSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet != null) {
    spreadsheet.deleteSheet(sheet);
  }
  sheet = spreadsheet.insertSheet(sheetName, 42);
  return sheet;
}

function getFullRow(sheet, startColumn, row) {
  var lastColumn = columnToLetter(sheet.getLastColumn());
  return sheet.getRange(startColumn + row + ':' + lastColumn + row);
}

function getFullColumn(sheet, column, startRow) {
  var lastRow = sheet.getLastRow();
  return sheet.getRange(column + startRow + ':' + column + lastRow);
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter) {
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}
