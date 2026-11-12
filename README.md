# Superstore-Business-Performance-Analysis

An Excel-based retail sales analysis of a superstore, exploring sales performance, profitability, and the relationship between discounting and profit. The project uses data cleaning, formula-based calculations, PivotTables, and an interactive dashboard to identify key performance trends, loss-making areas, and opportunities for improvement.


**Tools:** Microsoft Excel (Tables, formulas, PivotTables, PivotCharts, Slicers)

---

## Table of Contents
1. [Project Overview](#project-overview)
2. [Aim and Objectives](#aim-and-objectives)
3. [Business Questions](#business-questions)
4. [Dataset](#dataset)
5. [Analysis Process](#analysis-process)
6. [Key Findings](#key-findings)
7. [Recommendations](#recommendations)
8. [Dashboard](#dashboard)
9. [Limitations](#limitations)
10. [Skills Demonstrated](#skills-demonstrated)
11. [Repository Structure](#repository-structure)

---

## Project Overview

Sales figures alone don't show whether a business is healthy. A category can sell a lot and still make very little money. This project looks past total sales to examine **profit, losses and discounting** across product categories, customer segments, regions and states.

This self-directed project explores retail sales and profitability using a public Superstore dataset. The analysis was conducted entirely in Excel, covering data cleaning, calculated metrics, PivotTable analysis, and the development of an interactive dashboard to identify sales trends, profitability patterns, and areas of concern.

## Aim and Objectives

### Aim
To analyze retail order data to identify where the business makes and loses money, and to measure the effect of discounting on profit, so that clear, data-backed recommendations can be made.

### Objectives
1. **Clean and validate** the raw data by checking for missing values and removing duplicate records.
2. **Calculate new metrics** (pre-discount price, cost, losses, discount value, discount band and loss flag) that the raw data does not provide.
3. **Compare performance** across categories, sub-categories, customer segments, regions and states using sales, quantity, profit and loss.
4. **Quantify discounting** by measuring how much money is given away and at what discount level orders stop being profitable.
5. **Identify the loss-making areas** of the business, including products and states.
6. **Build an interactive dashboard** with Pivot tables and slicers so the results can be explored and filtered.
7. **Recommend actions** based on the findings.

## Business Questions

1. Which product categories and sub-categories drive sales, and which drive profit?
2. Which sub-categories lose money?
3. How do sales differ across customer segments?
4. How much money is given away through discounts, and where?
5. Is discounting hurting profit, and at what level?
6. Which states combine heavy discounting with poor profit?

## Dataset

**Source:** Sample Superstore dataset from [Kaggle](https://www.kaggle.com/datasets/bravehart101/sample-supermarket-dataset) 

| Item | Detail |
|---|---|
| Rows | 9,994 raw orders (9,977 after cleaning) |
| Columns | 13 original, 21 after calculated columns |

### Key Variables

All 21 columns in the cleaned table: 13 from the original dataset and 8 calculated columns I added.

| # | Column | Type | Source | Description |
|---|---|---|---|---|
| 1 | Ship Mode | Text | Original | Shipping method used for the order |
| 2 | Segment | Text | Original | Customer segment: Consumer, Corporate or Home Office |
| 3 | Country | Text | Original | Country of the order (United States) |
| 4 | City | Text | Original | City the order was shipped to |
| 5 | State | Text | Original | US state the order was shipped to |
| 6 | Postal Code | Number | Original | Postal (ZIP) code of the delivery location |
| 7 | Region | Text | Original | Sales region: Central, East, South or West |
| 8 | Category | Text | Original | Top-level product group: Furniture, Office Supplies or Technology |
| 9 | Sub-Category | Text | Original | Product type within a category (e.g. Tables, Phones) |
| 10 | Sales | Currency | Original | Revenue from the order line, recorded after the discount |
| 11 | Quantity | Number | Original | Number of units sold |
| 12 | Discount | Percentage | Original | Discount applied to the order line, stored as a fraction (0.2 = 20%) |
| 13 | Profit | Currency | Original | Profit on the order line; a negative value means a loss |
| 14 | Actual price | Currency | Calculated | Pre-discount price: `Sales / (1 - Discount)` |
| 15 | Cost | Currency | Calculated | Estimated cost of the order line: `Sales - Profit` |
| 16 | Actual profit | Currency | Calculated | Profit on profitable orders only, otherwise 0 |
| 17 | Loss | Currency | Calculated | Profit on loss-making orders only (negative), otherwise 0 |
| 18 | Losses | Currency | Calculated | Absolute value of Loss, so losses chart as positive numbers |
| 19 | Discount value | Currency | Calculated | Dollar amount given away in discounts: `Actual price * Discount` |
| 20 | Discount Band | Text | Calculated | Discount group for each order: None, Up to 20%, Above 20% to 40%, or Above 40% |
| 21 | Loss Flag | Number | Calculated | 1 if the order lost money (Profit below 0), otherwise 0; its average gives the share of loss-making orders |

## Analysis Process

### 1. Data cleaning and validation
Before building any metric, I checked whether the data could be trusted. Each check, what it found and what I did about it:

| Check | Finding | Action |
|---|---|---|
| **Missing values** (all 13 columns) | None found | No action needed |
| **Duplicate records** | 17 exact duplicate rows | Removed (9,994 → 9,977 rows), since they would inflate sales and profit |
| **Category consistency** | Each category column has the expected values: 4 ship modes, 3 segments, 4 regions, 3 categories, 17 sub-categories, 1 country | No action needed |
| **Spelling and spacing** | No stray spaces or mixed-case duplicates in any text column | No action needed |
| **Numeric ranges** | Discount between 0% and 80%; Quantity between 1 and 14 units; Sales always above zero; Profit never exceeds Sales | No impossible values found |

I then converted the cleaned range into an Excel Table so every formula and PivotTable updates automatically if the data changes.

### 2. Calculated columns
Eight calculated columns were added using structured Table formulas:

| New column | Logic | Purpose |
|---|---|---|
| **Actual price** | `Sales / (1 - Discount)` | Recovers the pre-discount price |
| **Cost** | `Sales - Profit` | Estimates the cost of each order |
| **Actual profit** | `IF(Profit > 0, Profit, 0)` | Isolates profitable orders |
| **Loss** | `IF(Profit < 0, Profit, 0)` | Isolates loss-making orders |
| **Losses** | `ABS(Loss)` | Positive loss values for charting |
| **Discount value** | `Actual price * Discount` | Dollar amount given away in discounts |
| **Discount Band** | `IFS(...)` on Discount | Groups orders into None, up to 20%, 20-40% and above 40% |
| **Loss Flag** | `IF(Profit < 0, 1, 0)` | Flags loss-making orders so the share of losing orders can be averaged |

### 3. PivotTable analysis
Twelve PivotTables summarise the data by segment, category, sub-category and state, covering sales, quantity, profit, cost, loss, discount and discount value. A further pivot groups orders by discount band to show orders, total profit and the share of loss-making orders at each discount level.

### 4. Dashboard
The results are presented in an interactive dashboard with PivotCharts and **slicers for Category, Region and Sub-Category**, so a viewer can filter the whole analysis.

## Key Findings

### Overall performance
| Metric | Value |
|---|---|
| Total sales | ~$2.30M |
| Total profit | ~$286K |
| Profit margin | ~12.5% |
| Total discount value | ~$567K |
| Orders that lost money | 1,869 (about 19% of orders) |
| Total losses on those orders | ~$156K |

Discounts given away appr($567K) were roughly double the profit the business kept appr($286K).

### 1. Furniture sells a lot but earns very little
| Category | Sales | Profit | Margin | Units sold |
|---|---|---|---|---|
| Technology | $836K | $145K | 17.4% | 6,939 |
| Furniture | $741K | $18K | 2.5% | 8,020 |
| Office Supplies | $719K | $122K | 17.0% | 22,861 |

Furniture generates almost as much revenue as Office Supplies but about one-sixth of the profit. Office Supplies sells by far the most units, at a healthy margin.

### 2. Three sub-categories lose money
| Sub-Category | Profit |
|---|---|
| Tables | **−$17.7K** |
| Bookcases | **−$3.5K** |
| Supplies | **−$1.2K** |

The strongest sub-categories are **Copiers appr($55.6K), Phones appr($44.5K) and Accessories appr($41.9K)**. Tables and Bookcases both sit in Furniture, which explains that category's weak margin.

### 3. Higher Discounts Are Associated with Significant Profit Losses
Grouping orders by discount level shows a sharp change:

| Discount level | Orders | Total profit | Share of orders losing money |
|---|---|---|---|
| None | 4,787 | +$321K | 0% |
| Up to 20% | 3,799 | +$101K | 14% |
| Above 20% to 40% | 459 | −$36K | 90% |
| Above 40% | 932 | −$100K | 100% |

Every order discounted at more than 40% lost money. Profit holds up to around 20% and breaks down beyond it.

### 4. Heavy-discount states are often loss-making states
| State | Discount value | Profit |
|---|---|---|
| Texas | $109K | −$25.8K |
| Pennsylvania | $71K | −$15.6K |
| California | $70K | **+$76.3K** |
| Ohio | $56K | −$17.0K |
| Illinois | $53K | −$12.6K |

Texas, Pennsylvania, Ohio and Illinois give large discounts and lose money. California discounts at a similar scale to Pennsylvania but is highly profitable, so discount volume alone doesn't explain the difference. A likely factor is *which products* are discounted and *how deeply*, which would be worth investigating.

### 5. Regional and segment picture
- **Profit by region:** West appr($108K) and East appr($92K) lead; South appr($47K) and Central appr($40K) trail.
- **Sales by segment:** Consumer 51% appr($1.16M), Corporate 31% appr($706K), Home Office 19% appr($429K).

## Recommendations

1. **Cap discounts at around 20%**, or require approval for anything above it. Orders above this level are overwhelmingly loss-making.
2. **Review pricing and discounting on Tables and Bookcases.** They drag down Furniture, the second-largest category by sales.
3. **Investigate heavy-discount, loss-making states** (Texas, Ohio, Pennsylvania, Illinois). Compare them with California to see what differs in product mix and discount depth.
4. **Protect and grow high-margin lines**, such as Copiers, Phones, Accessories and the main Office Supplies sub-categories.

## Dashboard

[Superstore Business Performance Analysis.png](Superstore-Business-Performance-Analysis.png)
The dashboard contains:
- Sales by category
- Total sales by segment
- Profit by sub-category
- Total quantity by category
- Total discount value by state
- Slicers for **Category**, **Region** and **Sub-Category**

## Limitations

- **No date field**, so there is no trend, seasonality or growth analysis.
- **Profit is taken as given.** Cost was derived as Sales − Profit, so it reflects the dataset's own profit figures, not an independent cost source.
- **Correlation, not causation.** The analysis shows that high discounts coincide with losses, but it can't prove discounts alone caused them (product mix, shipping and returns could also contribute).
- **Sample data.** Results demonstrate method, not real business performance.
