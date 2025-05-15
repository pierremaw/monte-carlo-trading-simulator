//@function       Main function to run the trade simulator. It performs simulations, calculates summary statistics, updates the statistics, and refreshes the equity value.
//@param          None.
//@returns        None.
function TradeSimulator() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const simCountCell = spreadsheet.getRangeByName("simCount");
  const simsToPerformCount = simCountCell.getValue();
  const namedRanges = getNamedRanges(spreadsheet);
  const results = performSimulations(namedRanges, simsToPerformCount);
  const summaryStats = calculateSummaryStatistics(results);
  updateSummaryStatistics(namedRanges, summaryStats);
  refreshEquityValueIsInRange(namedRanges.fourHundredTradesEquity, namedRanges.minValue, namedRanges.maxValue);
  getMedianQ1Q3(namedRanges, results);
}

//@function       Retrieves named ranges from the spreadsheet.
//@param          spreadsheet (Spreadsheet) The active spreadsheet instance.
//@returns        (Object) An object containing the named ranges in the spreadsheet.
function getNamedRanges(spreadsheet) {
  return {
    expectedValue: spreadsheet.getRangeByName("expectedValue"),
    medianValue: spreadsheet.getRangeByName("medianValue"),
    minValue: spreadsheet.getRangeByName("minValue"),
    maxValue: spreadsheet.getRangeByName("maxValue"),
    standardDeviation: spreadsheet.getRangeByName("standardDeviation"),
    fourHundredTradesEquity: spreadsheet.getRangeByName("fourHundredTradesEquity"),
    q1Value: spreadsheet.getRangeByName("q1Value"),
    q3Value: spreadsheet.getRangeByName("q3Value")
  };
}

//@function       Performs simulations based on the given simulation count.
//@param          namedRanges (Object) Named ranges in the spreadsheet.
//@param          simsToPerformCount (number) The count of simulations to perform.
//@returns        (Array) An array containing the results of the performed simulations.
function performSimulations(namedRanges, simsToPerformCount) {
  const results = [];
  for (let i = 0; i < simsToPerformCount; i++) {
    const fourHundredTradesEquity = namedRanges.fourHundredTradesEquity.getValue();
    results.push(fourHundredTradesEquity);
    updateInProgress(namedRanges, i + 1);
  }
  return results;
}

//@function       Updates progress for the simulation process.
//@param          namedRanges (Object) Named ranges in the spreadsheet.
//@param          currentSimulation (number) The current simulation number.
//@returns        None.
function updateInProgress(namedRanges, currentSimulation) {
  for (const key in namedRanges) {
    if (namedRanges[key] !== namedRanges.fourHundredTradesEquity) {
      namedRanges[key].setValue(`Updating...Sim ${currentSimulation}`);
    }
  }
}

//@function       Calculates summary statistics for the simulation results.
//@param          results (Array) An array containing the results of the simulations.
//@returns        (Object) An object containing the summary statistics.
function calculateSummaryStatistics(results) {
  const sortedResults = results.slice().sort((a, b) => a - b);
  const sum = results.reduce((acc, val) => acc + val, 0);
  return {
    expectedValue: sum / results.length,
    median: sortedResults[Math.floor(sortedResults.length / 2)],
    minValue: Math.min(...results),
    maxValue: Math.max(...results),
    standardDeviation: calculateStandardDeviation(results),
    q1: sortedResults[Math.floor(sortedResults.length / 4)],
    q3: sortedResults[Math.floor(sortedResults.length * 3 / 4)]
  };
}

//@function       Calculates the standard deviation of the simulation results.
//@param          results (Array) An array containing the results of the simulations.
//@returns        (number) The standard deviation of the results.
function calculateStandardDeviation(results) {
  const mean = results.reduce((acc, val) => acc + val, 0) / results.length;
  const variance = results.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / results.length;
  return Math.sqrt(variance);
}

//@function       Calculates the median, Q1, and Q3 from the simulation results and updates the spreadsheet.
//@param          namedRanges (Object) Named ranges in the spreadsheet.
//@param          results (Array) An array containing the results of the simulations.
//@returns        None.
function getMedianQ1Q3(namedRanges, results) {
  const sortedResults = results.slice().sort((a, b) => a - b);
  const median = sortedResults[Math.floor(sortedResults.length / 2)];
  const q1 = sortedResults[Math.floor(sortedResults.length / 4)];
  const q3 = sortedResults[Math.floor(sortedResults.length * 3 / 4)];
  namedRanges.medianValue.setValue(median);
  namedRanges.q1Value.setValue(q1);
  namedRanges.q3Value.setValue(q3);
}

//@function       Updates the summary statistics in the spreadsheet.
//@param          namedRanges (Object) Named ranges in the spreadsheet.
//@param          summaryStats (Object) Summary statistics to be updated.
//@returns        None.
function updateSummaryStatistics(namedRanges, summaryStats) {
  for (const key in namedRanges) {
    if (namedRanges[key] !== namedRanges.fourHundredTradesEquity) {
      namedRanges[key].setValue(summaryStats[key]);
    }
  }
}

//@function       Checks and updates the equity value to ensure it's within the specified range.
//@param          fourHundredTradesEquityCell (Range) The cell containing the equity value at 400 trades.
//@param          minValueCell (Range) The cell containing the minimum value.
//@param          maxValueCell (Range) The cell containing the maximum value.
//@returns        None.
function refreshEquityValueIsInRange(fourHundredTradesEquityCell, minValueCell, maxValueCell) {
  let fourHundredTradesEquityCellValue = fourHundredTradesEquityCell.getValue();
  let minValueCellValue = minValueCell.getValue();
  let maxValueCellValue = maxValueCell.getValue();
  while ((fourHundredTradesEquityCellValue > maxValueCellValue) || (fourHundredTradesEquityCellValue < minValueCellValue)) {
    SpreadsheetApp.flush(); // Refresh the spreadsheet to update values
    // Update values after the flush
    fourHundredTradesEquityCellValue = fourHundredTradesEquityCell.getValue();
    minValueCellValue = minValueCell.getValue();
    maxValueCellValue = maxValueCell.getValue();
    Utilities.sleep(50);
  }
}

