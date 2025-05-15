# Monte Carlo Trading Simulator for Google Sheets

This project features a **Monte Carlo Trading Simulator** created within a Google Sheet document. It models how a simple fixed-fraction trading strategy can cause a trading account to grow or shrink over many random sequences. You can tweak your risk assumptions, observe probabilistic outcomes, and analyze how compounding and volatility affect long-term results.

## Features and Capabilities

This simulator allows full user control over trade assumptions and simulates random sequences of wins and losses based on your estimated win rate. Each simulation models a complete sequence of trades and records the resulting ending balance. Once all simulations are complete, it reports statistical summaries (mean, median, quartiles, standard deviation, min, max) and overlays equity curves for visual insight into performance dispersion.

The key benefit: even if your system has a mathematical edge, variance can still produce significant short-term drawdowns—or outsized gains. By visualizing this, you gain a realistic perspective on strategy risk.

## Quick Start

### 1. Copy the Simulator Sheet

Click the link below and duplicate the simulator to your own Google Drive:

[**→ Copy This Google Sheet**](https://docs.google.com/spreadsheets/d/1llfX-jLt7N-dviIor4SKhrk3zuYnIL6PwtTl-bU366w/edit?gid=1990509670#gid=1990509670)

Then go to **File → Make a copy…**.

### 2. Update Parameters

In the **Input Variables** section, set:

* **Account Start** – e.g., \$5,000
* **Risk % Per Trade** – how much you risk per trade (e.g., 1%)
* **Reward % Per Trade** – how much you win on a successful trade (e.g., 5%)
* **Reward/Risk Ratio** – automatically calculated
* **Estimated Win Rate** – e.g., 41%
* **Simulation Count** – number of trade paths to simulate (e.g., 111)

The simulator recalculates risk/reward as percentages of your *current* account, not your starting balance.

### 3. Click “Perform Simulation”

After setting your inputs, just click the **“Perform Simulation”** button. The script runs the simulations, calculates summary statistics, and updates charts automatically.

## Spreadsheet Sections

### Input Variables

You define the base conditions of your trading model: starting balance, risk, reward, win rate, and the number of simulations.

### Summary Statistics

The simulator aggregates final account values across all runs and reports:

* **Expected Value (Mean)**
* **Median, Q1, Q3**
* **Minimum and Maximum Outcomes**
* **Standard Deviation**

These help you understand the distribution of possible outcomes—what’s typical, what’s rare, and what’s extreme.

### Active Simulation Statistics

You can drill down into one specific simulated path to see:

* Total trades simulated (e.g., 450)
* Count of wins and losses
* Final balance
* Total % gain or loss

### Active Simulation Trades

Each individual trade in the selected path is listed, showing:

* Trade result (Win or Loss)
* Account size after that trade
* % gain from the original starting equity

This reveals how compounding and drawdowns evolve within a sequence.

---

## What’s Under the Hood?

* **Randomized trade outcomes** based on the given win rate
* **Risk and reward per trade** are percentages of the *current* equity
* **Each simulation runs a full trade sequence** (e.g., 450 trades)
* **Aggregated results** across simulations show overall performance distribution

---

## Full Script Source (`Code.gs`)

Paste this into the Apps Script editor:

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

---

## Why Use This Simulator?

Monte Carlo simulations help you visualize how randomness and edge interact in trading. Even with a positive expected return, variance can produce losing stretches. By running hundreds or thousands of simulations, you can measure the risks of ruin, the potential for windfall gains, and the range of outcomes you’re statistically likely to see. This makes the tool not only educational, but vital for serious risk management.

