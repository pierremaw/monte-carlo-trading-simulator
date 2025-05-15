# Monte Carlo Trading Simulator

A Google Sheets–based Monte Carlo trading simulator that lets you model thousands of potential trading outcomes and see summary statistics (expected value, median, quartiles, etc.) and equity curves over a set number of trades. Everything runs in your spreadsheet via an Apps Script—no external tools required.

---

## 🔍 Features

- **User-configurable inputs**  
  - Account start balance  
  - Risk % per trade  
  - Reward % per trade  
  - Reward/Risk ratio  
  - Estimated win rate  
  - Simulation count  

- **Monte Carlo simulations**  
  - Runs _N_ independent equity‐curve simulations of a fixed trade sequence (e.g. 100 or 450 trades).  
  - Automatically retries any out‐of‐bounds equity results until they fall within the specified min/max envelope.

- **Summary statistics**  
  - Expected value (mean)  
  - Median, Q1, Q3  
  - Min / Max  
  - Standard deviation  

- **Equity curve charts**  
  - Live equity charts for sample simulation runs  
  - Overlay of multiple runs for visualizing dispersion

---

## 🚀 Getting Started

### Prerequisites

- A Google account with access to **Google Sheets** and **Apps Script**.
- A blank (or copy of this) Google Sheet with named ranges set up as below.

### 1. Make a Copy of the Spreadsheet

1. Open the [template sheet](./MONTE_CARLO_TRADING_SIMULATOR_TEMPLATE.xlsx) (or your own blank sheet).  
2. **File → Make a copy…** in your Google Drive.

### 2. Define the Named Ranges

In your sheet, create the following **named ranges** (exact names required):

| Named Range                | Cell / Range Purpose                                  |
|----------------------------|--------------------------------------------------------|
| `simCount`                 | Number of simulations to perform (e.g. 111, 711, etc.) |
| `fourHundredTradesEquity`  | Formula cell that computes equity at trade # (e.g. 100)|
| `expectedValue`            | Output: mean of simulation results                     |
| `medianValue`              | Output: median                                          |
| `minValue`                 | Output: minimum                                         |
| `maxValue`                 | Output: maximum                                         |
| `standardDeviation`        | Output: standard deviation                             |
| `q1Value`                  | Output: 1st quartile                                    |
| `q3Value`                  | Output: 3rd quartile                                    |

> **Tip:** Use **Data → Named ranges…** in the Google Sheets menu to create or edit these.

### 3. Install the Apps Script

1. In your copied sheet, open **Extensions → Apps Script**.  
2. Delete any existing code in `Code.gs`.  
3. Copy & paste the entire contents of [`Code.gs`](./src/Code.gs) (see below).  
4. Save the project (e.g. name it “MonteCarloSim”).  
5. Back in the sheet, add a button (Insert → Drawing or a shape), label it “Perform Simulation,” and assign it the script function `TradeSimulator`.

### 4. Run Your First Simulation

1. Enter your input parameters in the sheet (account start, risk %, reward %, win rate, simCount).  
2. Click **Perform Simulation**.  
3. Watch the “Updating… Sim X” progress and then see your summary statistics and equity curves update instantly.

---

## 📜 Apps Script Overview

```js
/**
 * Main entry point: runs Monte Carlo simulations,
 * computes summary stats, and updates the sheet.
 */
function TradeSimulator() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const simCnt  = ss.getRangeByName("simCount").getValue();
  const ranges  = getNamedRanges(ss);
  const results = performSimulations(ranges, simCnt);
  const stats   = calculateSummaryStatistics(results);

  updateSummaryStatistics(ranges, stats);
  refreshEquityValueIsInRange(ranges.fourHundredTradesEquity, ranges.minValue, ranges.maxValue);
  getMedianQ1Q3(ranges, results);
}

// … (see full source in src/Code.gs)
