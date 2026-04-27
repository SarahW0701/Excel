# Excel Projects
This repository contains a collection of applications developed using Microsoft Excel.

---

## Overview

The repository includes multiple independent projects. Each project is provided as an `.xlsx` or `.xlsm` file and can be used directly.

---

## Project List

| Name | Description | File |
|------|------------|------|
| Budget Tracker | Tracks and analyzes personal finances with automated aggregation and a dynamic overview dashboard | `Budget_Tracker.xlsm` |
| Stock Portfolio Manager | Tracks stock investments, performance, and dividends using current market data and automated calculations | `Stock_Portfolio_Manager.xlsx` |

---
---

## Detailed Descriptions of Projects
### Budget Tracker 

#### Purpose
Track and analyze personal income and expenses with structured data entry, automated aggregation, and multi-level financial overview.

#### Features
- Centralized expense entry with categorized tracking
- Automated aggregation across days, weeks, months, and years
- Conditional formatting based on spending limits
- Dynamic overview dashboard with time-based analysis (days, weeks, months)
- Annual financial planning with income, expenses, and balance calculation
- Supporting tools (currency conversion, quick totals, temporary notes)

#### Inputs
- **`Eingabe` sheet (primary usage):**
  - Date, amount, category (predefined), and details
  - Optional notes for temporary entries
  - Daily budget ("Tagespensum") for evaluation logic
- **`Finanzplan` sheets:**
  - Manual input for income and fixed expenses
- **System structure:**
  - Categories define aggregation across all sheets


#### Outputs
- **`Übersicht` dashboard:**
  - Aggregated expenses for last 30 days, 10 weeks, and 4 months
  - Category-based breakdowns with conditional formatting
- **`Ausgaben` sheets:**
  - Full-year tracking of daily, weekly, and monthly expenses
- **`Finanzplan` sheets:**
  - Income, expenses, and resulting balance (yearly and monthly)
- **Visual indicators:**
  - Color-coded performance relative to the defined daily budget
- **`Einzelbeträge`:**
  - Central database of all recorded transactions
  - Filter options by date, amount, category, and detail

#### How to Use
1. Enter expenses in the `Eingabe` sheet and click "Hinzufügen"
2. Maintain asset categories consistently
3. Optionally define a daily budget ("Tagespensum")
4. Review short-term trends in `Übersicht`
5. Analyze yearly data in `Ausgaben` and `Financial Plan` sheets

#### Notes / Limitations
- Only the `Eingabe` sheet should be used for regular data entry
- Data consistency depends on correct category usage
- Budget evaluation is based on a fixed daily threshold ("Tagespensum")
- Real-time editing of historical data should be done in `Einzelbeträge`
- Example data is included for demonstration purposes

---

### Stock Portfolio Manager

#### Purpose
Track stock and ETF investments, including performance, current value, and dividends.

#### Features
- Automated portfolio calculations (value, averages, weights)
- Real-time price integration via external data source
- Dashboard with performance metrics and visualizations
- Dividend tracking with yearly aggregation

#### Inputs
- **`Eingabe` sheet:**
  - Transactions (buy/sell: date, asset, type, price per share)
  - Dividends (date, asset, dividend)
- **`Portfolio` sheet (initial setup only):**
  - Asset name, WKN, ISIN

#### Outputs
- **`Portfolio` sheet:**
  - Holdings, average price, current value, and allocation
- **`Dashboard`:**
  - Total investment, current value, dividends, and profit
  - Performance indicators with color coding
  - Portfolio allocation (pie chart)
  - Per-position performance overview
  - Dividend history with yearly breakdown and line chart

#### How to Use
1. Add new assets in the `Portfolio` sheet (name, WKN, ISIN)
2. Record transactions and dividends in the `Eingabe` sheet
3. Refresh data to update real-time prices
4. Review performance and analytics in the `Dashboard`

#### Notes / Limitations
- Real-time prices depend on data from an external website (boerse.de)
- Data must be refreshed manually
- Only the `Eingabe` and initial `Portfolio` setup should be modified
- Asset identification is based on WKN
