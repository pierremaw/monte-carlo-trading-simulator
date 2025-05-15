# Monte Carlo Trading Simulator for Google Sheets

https://github.com/user-attachments/assets/a379e21c-1924-41cb-b9ea-fccf72632510

This project features a **Monte Carlo Trading Simulator** created within a Google Sheet document. It models how a simple fixed-fraction trading strategy can cause a trading account to grow or shrink over many random sequences. You can tweak your risk assumptions, observe probabilistic outcomes, and analyze how compounding and volatility affect long-term results.

Monte Carlo simulations help visualize how randomness and edge interact in trading. Even with a positive expected return, variance can produce losing stretches. By running simulations, you can visualize the potential for gains/losses and what is statistically likely to happen based on specific input variables.

## Functionality

This simulator allows full user control over trade assumptions and simulates random sequences of wins and losses based on your estimated win rate. Each simulation models a complete sequence of trades and records the resulting ending balance. Once all simulations are complete, it reports statistical summaries (mean, median, quartiles, standard deviation, min, max) for the set of simulations that were performed.

## Quick Start

### 1. Copy the Simulator Sheet

Click the link below and duplicate the simulator to your own Google Drive:

[**→ Copy This Google Sheet**](https://docs.google.com/spreadsheets/d/1llfX-jLt7N-dviIor4SKhrk3zuYnIL6PwtTl-bU366w/edit?gid=1990509670#gid=1990509670)

### 2. Update Parameters

In the **Input Variables** section, if you want, you can update:

* **Account Start** – e.g., \$5,000
* **Risk % Per Trade** – how much you risk per trade (e.g., 1%)
* **Reward % Per Trade** – how much you win on a successful trade (e.g., 5%)
* **Reward/Risk Ratio** – automatically calculated
* **Estimated Win Rate** – e.g., 41%
* **Simulation Count** – number of trade paths to simulate (e.g., 111)

### 3. Click “Perform Simulation”

After setting your inputs, just click the **“Perform Simulation”** button. The script runs the simulations, calculates summary statistics, and updates charts automatically.

## How It Works

* **Randomized trade outcomes** based on the given win rate
* **Risk and reward per trade** are percentages of the *current* equity
* **Each simulation runs a full trade sequence** (e.g., 450 trades)
* **Aggregated results** across simulations show overall performance distribution

## Spreadsheet Sections

### Input Variables

You define the base conditions of your trading model: starting balance, risk, reward, win rate, and the number of simulations.

### Summary Statistics

The simulator aggregates final account values across all runs and reports:

* **Expected Value (Mean)**
* **Median, Q1, Q3**
* **Minimum and Maximum Outcomes**
* **Standard Deviation**

These help provide insight on the distribution of possible outcomes—what’s typical, what’s rare, and what’s extreme.

### Active Simulation Statistics

After you hit **Perform Simulation**, the sheet runs all the requested paths, then displays details for the **most recent** 450-trade path:

- **Total Trades** simulated (450)  
- **Win/Loss Count** and win percentage for that run  
- **Ending Balance**  
- **Overall % Gain or Loss** relative to your starting equity  

This snapshot shows exactly how one full sequence played out.

### Active Simulation Trades

For that latest 450-trade path, the sheet lists every trade in order, including:

1. **Result** (Win / Loss)  
2. **Account Size** immediately after the trade  
3. **% Change** from the original starting balance  

Watching these rows lets you see how compounding gains—and drawdowns—unfold in a single simulated journey.

## Full Script Source (`Extensions/Apps Script/Code.gs`)

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
