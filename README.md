# Midterm — Comprehensive Business Analytics & Forecasting

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)
![Domain](https://img.shields.io/badge/Domain-Forecasting%20%7C%20Analytics%20%7C%20BI-0078D4?style=flat)
![Type](https://img.shields.io/badge/Type-Midterm%20%7C%20Graded%20Work-8A2BE2?style=flat)
![Status](https://img.shields.io/badge/Status-Complete-2ea44f?style=flat)

## Overview

A full-scope midterm exam covering the entire business analytics pipeline — from data cleaning and preparation through time series forecasting, dashboard design, customer profiling, and conjoint analysis. Applied across two major business cases: a **multi-product café** (cookies, soda, coffee, specialty coffee) and the **Natural Soap Company**.

This project demonstrates end-to-end analytical thinking: defining the problem, cleaning data, selecting the right model, forecasting future demand, visualizing results, and making business recommendations.

---

## What's Inside

| Sheet | Description |
|-------|-------------|
| `Sheet1` | Data analytics process — problem definition methodology |
| `Sheet2` | Data cleaning process — normalization and preparation approach |
| `Sheet3` | Additional analytical framework |
| `Data Determination` | Monthly profit margin data — café product line analysis |
| `Forecast_Cookies` | Time series forecast — cookie sales with confidence intervals |
| `Forecast_Soda` | Time series forecast — soda sales with confidence intervals |
| `Forecast_Coffee` | Time series forecast — coffee sales with confidence intervals |
| `Forecast_Specialty` | Time series forecast — specialty coffee with confidence intervals |
| `Detail1 / Detail2` | Profit margin breakdown by product |
| `Visualization Data` | Raw data for dashboard construction |
| `Dashboard` | Business intelligence dashboard — revenue and performance KPIs |
| `Natural Soap` | Customer data overview — soap purchase and profile data |
| `Soap Sales` | Transaction-level purchase data — soap units, pricing, lotion cross-sell |
| `Soap Customer Profiles` | Customer demographics — age, gender, income, preferences |
| `Merged_soap` | Cleaned merged dataset — purchases + customer profiles joined |
| `(b)(c)(d)` | Regression and analysis steps on soap customer data |
| `Soap Conjoint` | Conjoint analysis — Natural Soap Company product optimization |

---

## Key Analysis Performed

### Part 1 — Café Sales Forecasting (4 Products)
**Business question:** A café needs to forecast demand for cookies, soda, coffee, and specialty coffee to optimize inventory and staffing.

- Analyzed monthly profit margin data across all 4 product lines
- Built individual time series forecasts per product using Excel's forecasting engine
- Generated **upper and lower confidence bounds** for each forecast
- Identified which products have stable vs. volatile demand patterns
- Recommended inventory ordering strategy based on forecast uncertainty bands

**Sample forecast output:**

| Product | Forecast Method | Confidence Interval |
|---------|----------------|---------------------|
| Cookies | Time series | ±upper/lower bounds |
| Soda | Time series | ±upper/lower bounds (high volume: ~2,989/month) |
| Coffee | Time series | ±upper/lower bounds (~4,630/month baseline) |
| Specialty Coffee | Time series | ±upper/lower bounds (~234/month) |

### Part 2 — Business Intelligence Dashboard
- Built a multi-chart Excel dashboard summarizing café performance
- Visualized annual revenue by product (Product 4 leading at $324K+)
- Designed for executive-level consumption — key KPIs visible at a glance

### Part 3 — Natural Soap Company Customer Analytics
- **Data merging:** Joined purchase transaction data with customer profile data on Customer ID
- **Regression analysis:** Modeled soap purchase behavior against age, income, ingredient preference, and fragrance preference
- **Customer profiling:** Segmented soap buyers by demographics and product preferences
- **Conjoint analysis:** Quantified which soap attributes (ingredients, scent, packaging, price) drive purchase decisions most

---

## Data Pipeline Demonstrated

```
Raw Data (Soap Sales + Customer Profiles)
        ↓
Data Cleaning & Normalization (Sheet2 methodology)
        ↓
Data Merging (Merged_soap — purchase + profile join)
        ↓
Regression Analysis (b, c, d sheets)
        ↓
Conjoint Analysis (Soap Conjoint)
        ↓
Business Recommendations
```

---

## Key Findings

> Specialty coffee has the lowest volume (~234 units/month) but likely the highest margin per unit — forecasting with tight confidence bounds allows the café to minimize overstock risk on this premium product.

> Customer income and soap ingredient preference are the strongest predictors of soap purchase value — higher-income customers show strong preference for natural/organic ingredients, suggesting a premium product line strategy.

---

## Skills Demonstrated

- Time series forecasting with confidence intervals
- Multi-product demand planning
- Data cleaning and preparation methodology
- Data merging (VLOOKUP/Power Query equivalent)
- Regression analysis on customer behavioral data
- Conjoint analysis and part-worth utility calculation
- Business dashboard design (Excel)
- End-to-end analytics pipeline execution
- Excel: Forecast.ETS, PivotTables, advanced charting, Data Analysis ToolPak

---

## Tools Used

- Microsoft Excel (Forecast.ETS, Data Analysis ToolPak, PivotTables, Dashboard design)
