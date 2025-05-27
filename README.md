# Monte Carlo Trading Simulator for Google Sheets

[https://github.com/user-attachments/assets/a379e21c-1924-41cb-b9ea-fccf72632510](https://github.com/user-attachments/assets/a379e21c-1924-41cb-b9ea-fccf72632510)

The **Monte Carlo Trading Simulator** is a dynamic Google Sheets tool designed to model the performance of a fixed-fraction trading strategy across many randomized trade sequences. By simulating hundreds of possible outcomes based on your inputs—such as win rate and risk per trade—it reveals how probability, compounding, and variance shape the long-term trajectory of a trading account.

This simulator provides a practical demonstration of how randomness and edge interact in trading systems. Even strategies with a positive expected return can experience significant drawdowns due to variance. Monte Carlo simulations help you visualize this uncertainty and quantify the statistical expectations of a given trading setup.

## Features

This simulator gives you full control over key trading assumptions and simulates thousands of trade sequences. For each run, it calculates and stores the final equity value. Once all simulations are complete, it compiles comprehensive statistical summaries including mean, median, quartiles, and standard deviation. The simulator updates both summary metrics and visual charts in real time.

## Getting Started

### 1. Copy the Simulator

Begin by duplicating the simulator to your own Google Drive:

[**→ Copy This Google Sheet**](https://docs.google.com/spreadsheets/d/1llfX-jLt7N-dviIor4SKhrk3zuYnIL6PwtTl-bU366w/edit?gid=1990509670#gid=1990509670)

### 2. Configure Your Parameters

In the **Input Variables** section, enter your desired simulation settings:

* **Account Start**: Starting balance (e.g., \$5,000)
* **Risk % Per Trade**: Percentage of equity risked per trade (e.g., 1%)
* **Reward % Per Trade**: Percentage gained per winning trade (e.g., 5%)
* **Estimated Win Rate**: Probability of a trade being successful (e.g., 41%)
* **Simulation Count**: Number of full trade sequences to run (e.g., 111)

### 3. Run the Simulation

Click the **“Perform Simulation”** button. The Google Apps Script backend will simulate the specified number of trade paths, compute summary statistics, and update the associated charts and tables.

## How It Works

Each simulation involves a random sequence of trade outcomes based on your defined win rate. Each trade adjusts the equity based on current size, applying your defined reward and risk parameters. A single simulation might contains 450 trades, capturing the compounding impact of gains and drawdowns.

After running all simulations, the system aggregates the final results and calculates the statistical distribution of potential account outcomes. This reveals both the typical and extreme trajectories under your chosen strategy.

## Spreadsheet Overview

### Input Variables

This section captures the foundational assumptions of the trading model. You specify starting capital, trade risk/reward percentages, estimated win rate, and how many simulations to perform.

### Summary Statistics

After all simulations are complete, the spreadsheet reports key distribution metrics across the simulated outcomes:

* **Expected Value (Mean)**
* **Median, Q1, Q3**
* **Minimum and Maximum Values**
* **Standard Deviation**

These statistics help quantify the central tendency and spread of possible final outcomes, providing context for risk and reward expectations.

### Active Simulation Summary

For the most recent trade path (typically 450 trades), the simulator displays:

* **Number of Trades Simulated**
* **Count of Wins and Losses**
* **Win Percentage**
* **Final Account Balance**
* **Net % Gain or Loss**

This provides a detailed look at a single example of how a trade sequence could play out under your inputs.

### Trade-by-Trade Breakdown

The simulator also lists every trade in the most recent simulation path:

* **Trade Outcome**: Win or Loss
* **Equity After Each Trade**
* **Cumulative % Change from Start**

This granular view makes it easy to observe how streaks, volatility, and compounding play out over a sequence.

## Apps Script Backend

The core functionality is implemented using Google Apps Script. Here is the source code for the backend logic found under `Extensions > Apps Script > Code.gs`:

```javascript
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

function performSimulations(namedRanges, simsToPerformCount) {
  const results = [];
  for (let i = 0; i < simsToPerformCount; i++) {
    const fourHundredTradesEquity = namedRanges.fourHundredTradesEquity.getValue();
    results.push(fourHundredTradesEquity);
    updateInProgress(namedRanges, i + 1);
  }
  return results;
}

function updateInProgress(namedRanges, currentSimulation) {
  for (const key in namedRanges) {
    if (namedRanges[key] !== namedRanges.fourHundredTradesEquity) {
      namedRanges[key].setValue(`Updating...Sim ${currentSimulation}`);
    }
  }
}

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

function calculateStandardDeviation(results) {
  const mean = results.reduce((acc, val) => acc + val, 0) / results.length;
  const variance = results.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / results.length;
  return Math.sqrt(variance);
}

function getMedianQ1Q3(namedRanges, results) {
  const sortedResults = results.slice().sort((a, b) => a - b);
  const median = sortedResults[Math.floor(sortedResults.length / 2)];
  const q1 = sortedResults[Math.floor(sortedResults.length / 4)];
  const q3 = sortedResults[Math.floor(sortedResults.length * 3 / 4)];
  namedRanges.medianValue.setValue(median);
  namedRanges.q1Value.setValue(q1);
  namedRanges.q3Value.setValue(q3);
}

function updateSummaryStatistics(namedRanges, summaryStats) {
  for (const key in namedRanges) {
    if (namedRanges[key] !== namedRanges.fourHundredTradesEquity) {
      namedRanges[key].setValue(summaryStats[key]);
    }
  }
}

function refreshEquityValueIsInRange(fourHundredTradesEquityCell, minValueCell, maxValueCell) {
  let fourHundredTradesEquityCellValue = fourHundredTradesEquityCell.getValue();
  let minValueCellValue = minValueCell.getValue();
  let maxValueCellValue = maxValueCell.getValue();
  while ((fourHundredTradesEquityCellValue > maxValueCellValue) || (fourHundredTradesEquityCellValue < minValueCellValue)) {
    SpreadsheetApp.flush();
    fourHundredTradesEquityCellValue = fourHundredTradesEquityCell.getValue();
    minValueCellValue = minValueCell.getValue();
    maxValueCellValue = maxValueCell.getValue();
    Utilities.sleep(50);
  }
}
```

This backend ensures smooth automation, accurate results, and real-time updates to your spreadsheet's UI.
