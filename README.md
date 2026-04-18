# Excel Projects
This repository contains a collection of applications developed using Microsoft Excel.

---

## Overview

The repository includes multiple independent projects. Each project is provided as an `.xlsx` or `.xlsm` file and can be used directly.

---

## Project List

| Name | Description | File |
|------|------------|------|
| Budget Tracker | Tracks income and expenses in a structured format | `Budget_Tracker.xlsm` |
| Stock Portfolio Manager | Tracks stock investments, performance, and dividends with automated calculations | `Stock_Portfolio_Manager.xlsx` |

Records and monitors personal stock investments

---
---

## Detailed Descriptions of Projects
### Budget Tracker 

#### Purpose
Brief explanation of what the workbook is used for.

#### Features
- Feature 1
- Feature 2
- Feature 3

#### Inputs
- What the user enters (sheets, cells, forms, etc.)

#### Outputs
- What the workbook produces (tables, summaries, charts, etc.)

#### How to Use
1. Step 1
2. Step 2
3. Step 3

#### Notes / Limitations
- Known constraints
- Things the user should avoid changing
- Assumptions in the design

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
- **Eingabe sheet:**
  - Transactions (buy/sell: date, asset, type, price per share)
  - Dividends (date, asset, dividend)
- **Portfolio sheet (initial setup only):**
  - Asset name, WKN, ISIN

#### Outputs
- **Portfolio sheet:**
  - Holdings, average price, current value, and allocation
- **Dashboard:**
  - Total investment, current value, dividends, and profit
  - Performance indicators with color coding
  - Portfolio allocation (pie chart)
  - Per-position performance overview
  - Dividend history with yearly breakdown and line chart

#### How to Use
1. Add new assets in the *Portfolio* sheet (name, WKN, ISIN)
2. Record transactions and dividends in the *Eingabe* sheet
3. Refresh data to update real-time prices
4. Review performance and analytics in the *Dashboard*

#### Notes / Limitations
- Real-time prices depend on data from an external website (börse.de)
- Data must be refreshed manually
- Only the *Eingabe* and initial *Portfolio* setup should be modified
- Asset identification is based on WKN
