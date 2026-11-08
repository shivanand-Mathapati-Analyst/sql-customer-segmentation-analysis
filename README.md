<div align="center">

# 🗄️ Customer Segmentation & Revenue Analysis — SQL Server

### End-to-end SQL pipeline: raw retail data → cleaned star schema → RFM segmentation → 26-view reporting layer

[![Tool](https://img.shields.io/badge/Tool-Microsoft%20SQL%20Server-CC2927?style=for-the-badge&logo=microsoftsqlserver&logoColor=white)](https://www.microsoft.com/en-us/sql-server)
[![T-SQL](https://img.shields.io/badge/T--SQL-26%20Views%20%2B%202%20Procedures-217346?style=for-the-badge)](#-reporting-layer)
[![Data Model](https://img.shields.io/badge/Data%20Model-Star%20Schema-blue?style=for-the-badge)](#️-data-model)
[![RFM](https://img.shields.io/badge/Segmentation-RFM%20via%20NTILE(5)-8A2BE2?style=for-the-badge)](#-rfm--segmentation-logic)
[![Status](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge)](#)

</div>

---

## 📑 Table of Contents

- [Overview](#-overview)
- [Key Features](#-key-features)
- [Data Model](#️-data-model)
- [RFM & Segmentation Logic](#-rfm--segmentation-logic)
- [Reporting Layer](#-reporting-layer)
- [Data Cleaning](#-data-cleaning)
- [Tech Stack](#️-tech-stack)
- [Repository Structure](#-repository-structure)
- [How to Use](#-how-to-use)
- [Key Findings](#-key-findings)
- [Root Causes](#-root-causes)
- [Business Insights](#-business-insights)
- [Recommendations & Actions](#-recommendations--actions)
- [Related Projects](#-related-projects)
- [Author](#-author)
- [License](#-license)

---

## 📌 Overview

This project rebuilds the same **Online Retail RFM Customer Segmentation** analysis found in the Excel and Power BI versions of this portfolio — but entirely in **Microsoft SQL Server (T-SQL)**, from raw CSV import through to a full reporting layer of views and stored procedures. Every step — data cleaning, dimensional modeling, RFM scoring, and business reporting — is done in native SQL, with no external BI tool required to answer the core business questions.

This project is a comprehensive **Customer Segmentation & Revenue Analysis solution built using Microsoft SQL Server** and the UK-based **Online Retail** dataset (2009–2011, UK retail operations). It transforms 541,909 transaction records into meaningful business insights, covering **customer segmentation, purchasing behavior, revenue contribution, customer retention, and churn analysis**.

**This project is part of a three-tool analytics portfolio built on the same UK-based Online Retail dataset.**

This project uses the same Online Retail dataset as my [Power BI]() and [Excel]() projects, demonstrating how the same business problem can be analyzed and solved using three different analytics tools: **SQL Server, Power BI, and Microsoft Excel**.

The goal was to go beyond basic SQL queries and develop a structured, query-driven analytical solution to extract meaningful insights directly from transactional data. Leveraging **T-SQL, data cleaning, aggregations, joins, Common Table Expressions (CTEs), subqueries, window functions, and views**, this project applies **RFM (Recency, Frequency, Monetary) analysis** to segment customers based on purchasing behavior, identify high-value and at-risk customers, evaluate revenue contribution, analyze customer retention, and uncover actionable business opportunities.

The single `.sql` script is fully documented inline: every query has its own heading and purpose comment, every cleaning decision is validated with a `COUNT`/`CHECK` query before and after, and the final section of the script contains a full write-up of findings, root causes, business insights, and recommendations — derived directly from the data itself.


## 📌 Overview

This project rebuilds the same **Online Retail RFM Customer Segmentation** analysis featured in the Excel and Power BI versions of this portfolio, but entirely in **Microsoft SQL Server (T-SQL)**. The solution covers the complete analytical workflow — from raw CSV data import and data cleaning to dimensional modeling, RFM scoring, customer segmentation, and a structured reporting layer using SQL views and stored procedures. Every stage is implemented using native SQL, without relying on external BI tools for the core analysis.

Built on the UK-based **Online Retail** dataset (2009–2011, UK retail operations), the project transforms 541,909 transaction records into actionable business insights across **customer segmentation, purchasing behavior, revenue contribution, customer retention, and churn analysis**.

**This project is part of a three-tool analytics portfolio built on the same UK-based Online Retail dataset.**

The same business problem is independently explored using three different analytics tools: [Power BI](YOUR_POWER_BI_REPOSITORY_LINK), [Microsoft Excel](YOUR_EXCEL_REPOSITORY_LINK), and SQL Server. This demonstrates how different analytical platforms can be used to approach the same business challenge and derive meaningful insights.

The objective was to go beyond basic SQL queries and develop an end-to-end, query-driven analytical solution. Using **T-SQL, data cleaning, aggregations, joins, Common Table Expressions (CTEs), subqueries, window functions, views, and stored procedures**, the project applies **RFM (Recency, Frequency, Monetary) analysis** to evaluate customer value, identify high-value and at-risk segments, measure revenue contribution, assess customer retention, and uncover opportunities for business growth.

The entire workflow is documented in a single `.sql` script, with clearly structured sections, descriptive comments, and explanations of each analytical step. Data-cleaning decisions are supported by validation queries using `COUNT` and other checks before and after transformations. The script concludes with a comprehensive analysis of **key findings, root causes, business insights, and actionable recommendations**, all derived directly from the dataset.

| | |
|---|---|
| 🗂️ **Raw records** | 541,909 rows → 401,560 cleaned transaction lines |
| 💰 **Gross Revenue** | £8.89M (Net Revenue: £8.28M after returns) |
| 👥 **Customers** | 4,371 unique customers |
| 🛍️ **Products** | 3,683 unique SKUs |
| 🌎 **Coverage** | 37 countries |
| 📅 **Time Range** | Dec 2010 – Dec 2011 |
| 🗄️ **Database** | `CustomerSegmentDB`, 3 schemas (`original`, `datamodel`, `reporting`) |
| 📊 **Reporting objects** | 26 views + 2 stored procedures + 1 reference table |

---

## ✨ Key Features

- **Full ETL pipeline in T-SQL** — raw import → validation → cleaning → star schema, with every step counted and verified before moving to the next (no silent data loss)
- **A defensive 3-schema design** — `original` (raw, never touched again), `datamodel` (the trusted star schema), `reporting` (the consumption layer) — mirroring how a real data warehouse separates untrusted, modeled, and reporting-ready data
- **RFM scoring engine built with window functions** — `NTILE(5)` for percentile-based R/F/M scoring, with a documented edge-case fix for customers with return-only purchase histories (who would otherwise score incorrectly)
- **A recursive CTE-generated date dimension** — `Dim_Date` built entirely in SQL (no external calendar table), including a documented workaround for SQL Server's rule against aggregates inside recursive CTE members
- **26 reporting views across 7 analytical sections** — Executive Overview, Segmentation, Customer, Product, Geography, Time, and Insights & Actions — plus 2 parameterized stored procedures for customer-level and segment-level lookups
- **A fully written findings section** — Key Findings, Root Causes, Business Insights, and Recommendations, all derived from and cited against the actual query outputs, including two cross-referenced anomalies (a partial final month, and a bulk-order-and-cancel pattern skewing product rankings) that only became visible by comparing multiple views together

---

## 🗃️ Data Model

A star schema with one fact table and four dimension tables, built with `SELECT ... INTO` and enforced with primary/foreign keys.

| Table | Type | Description |
|---|---|---|
| **datamodel.Fact_Sales** | Fact | 401,560 rows — invoice line items: quantity, unit price, sales amount, return flag, order date, customer/product keys |
| **datamodel.Dim_Customers** | Dimension | 4,371 rows — one per customer, with Recency/Frequency/Monetary, R/F/M scores, RFM_Score, and Segment calculated as columns |
| **datamodel.Dim_Products** | Dimension | 3,683 rows — one per unique product, description resolved by most-frequent-occurrence where a StockCode had inconsistent naming |
| **datamodel.Dim_Countries** | Dimension | 37 rows — one per unique country |
| **datamodel.Dim_Date** | Dimension | 374 rows (Dec 2010 – Dec 2011), generated via a recursive CTE, with Year/Month/Quarter/Weekday/IsWeekend columns |

**Relationships**: `Fact_Sales` → `Dim_Products` (StockCode), `Fact_Sales` → `Dim_Customers` (CustomerID), `Fact_Sales` → `Dim_Date` (OrderDate), and `Dim_Customers` → `Dim_Countries` (CustomersCountry) — a snowflaked country dimension, since country lives on the customer, not the transaction.

---

## 🎯 RFM & Segmentation Logic

Recency, Frequency, and Monetary are calculated per customer, then converted to 1–5 scores using `NTILE(5)` window functions, and combined into a final 5-segment classification.

```sql
-- R_Score: NTILE ordered DESC on Recency, so the smallest (best) Recency
-- values land in the highest tile -- NTILE has no separate inversion step,
-- so the sort direction itself has to encode "lower is better."
NTILE(5) OVER (ORDER BY Recency DESC)   AS R_Score,
NTILE(5) OVER (ORDER BY Frequency ASC)  AS F_Score,
NTILE(5) OVER (ORDER BY Monetary ASC)   AS M_Score

-- Final segment: ordered most-specific-first so a Champion never falls
-- through to a broader match; "Loyal Customers" is an unconditional
-- catch-all, so every customer lands in exactly one of 5 segments.
CASE
    WHEN R_Score >= 4 AND F_Score >= 4 AND M_Score >= 4 THEN 'Champions'
    WHEN R_Score >= 4 THEN 'Potential Loyalists'
    WHEN R_Score <= 2 AND F_Score <= 2 AND M_Score <= 2 THEN 'Lost Customers'
    WHEN R_Score <= 2 THEN 'At Risk'
    ELSE 'Loyal Customers'
END
```

| Segment | Customers | Revenue | Revenue % |
|---|---:|---:|---:|
| 🏆 Champions | 1,009 | £5.85M | 65.84% |
| 💙 Loyal Customers | 874 | — | — |
| 🔵 Potential Loyalists | 739 | — | — |
| 🟠 At Risk | 933 | — | — |
| 🔴 Lost Customers | 816 | £175.17K | 1.97% |

> One notable data-quality fix: 33 customers had return-only purchase histories (no completed, non-return orders). A naive `INNER JOIN`-based calculation left their Recency/Frequency/Monetary as `NULL`, which then caused a ranking-function edge case. The fix uses a `LEFT JOIN` with `COALESCE()` fallbacks (a dynamically calculated max-Recency for Recency, and `0` for Frequency/Monetary) so these customers are honestly and correctly classified as **Lost Customers**, rather than accidentally scoring well due to how blank values sort.

---

## 📊 Reporting Layer

26 views across 7 sections, plus 2 stored procedures, all under the `reporting` schema:

| Section | Example Views |
|---|---|
| **Executive Overview** | `vw_ExcecutiveKPIs`, `vw_RevenueTrendbyMonth`, `vw_TopCountriesbyRevenue`, `vw_RevenueReturnRatebyMonth` |
| **Segmentation Analysis** | `vw_SegmentationKPIs`, `vw_SegmentationDetail`, `vw_SegmentDistribution` |
| **Customer Analysis** | `vw_Top10CustomersByRevenue`, `vw_RecencyDistribution`, `vw_CLVDistribution`, `vw_CustomerRevenuePareto` |
| **Product Analysis** | `vw_Top10ProductsByRevenue`, `vw_ProductRevenuePareto`, `vw_ReturnsByProduct` |
| **Geographical Analysis** | `vw_RevenueByCountry`, `vw_SegmentByCountry`, `vw_ReturnRateByCountry` |
| **Time Analysis** | `vw_RevenueByQuarter`, `vw_OrdersByDayOfWeek`, `vw_NewCustomersByMonth` |
| **Insights & Actions** | `vw_RevenueRisk`, `vw_AtRiskCustomers`, `vw_RecommendedActions` |

**Stored procedures:**
```sql
EXEC reporting.sp_GetCustomerDetails @CustomerID = 14646;
-- Returns a customer's RFM profile + full transaction history

EXEC reporting.sp_GetCustomerSegmentDetails @Segment = 'Champions';
-- Returns every customer in a given segment, sorted by value
```

---

## 🔄 Data Cleaning

- Raw CSV imported with **every column as `NVARCHAR`**, so the import can never fail on a type mismatch in messy source data — type conversion happens deliberately, later, once cleaning has started
- Removed 135,080 rows with a blank `CustomerID` (541,909 → 406,829)
- Removed 5,225 true duplicate rows via `ROW_NUMBER()` + `PARTITION BY` across all 8 columns (406,829 → 401,604)
- Fixed a UK-date-format bug: `InvoiceDate` (`dd-mm-yyyy`) was being misread under the server's default `us_english` locale — resolved with `SET DATEFORMAT dmy` before casting to `DATETIME`
- Added an `IsReturn` flag (`'Yes'`/`'No'`) based on the `'C'`-prefixed invoice convention
- Removed 44 rows with `UnitPrice <= 0` (401,604 → 401,560)
- Derived `SalesAmount = UnitPrice * Quantity`

---

## 🛠️ Tech Stack

| Tool | Purpose |
|---|---|
| **Microsoft SQL Server** | Database engine, T-SQL |
| **T-SQL window functions** | `NTILE()`, `ROW_NUMBER()`, `RANK()`, `SUM() OVER()` for scoring, deduplication, and Pareto/running-total analysis |
| **Recursive CTEs** | Generating the Dim_Date calendar table |
| **CSV (Online Retail, UCI)** | Source dataset |

---

## 📁 Repository Structure

```
Customer-Segmentation-SQL-Server/
│
├── sql/
│   └── Customer-Segmentation-SQLServer-FINAL.sql   # Full, single-file script
├── raw data/
│   └── OnlineRetail.csv                            # Source dataset
├── images/
│   └── data-model.png
└── README.md
```

---

## 🚀 How to Use

1. Open `Customer-Segmentation-SQLServer-FINAL.sql` in **SQL Server Management Studio (SSMS)**
2. Update the raw file path if re-importing via SSMS's **Import Flat File** wizard (`original.OnlineRetail`)
3. Run the script from top to bottom — it's organized into 8 numbered sections and is safe to run sequentially
4. Query any of the 26 views under `reporting.*`, or call the two stored procedures directly, to explore the analysis

---

## 🧩 Key Findings

- **Champions (1,009 customers, 23.1% of the base) generated 65.84% of total revenue** (£5.85M) — the same concentration pattern repeats in products (top 20% of SKUs drive 78.80% of product revenue) and countries (top 8 of 38 countries drive 96.01% of revenue)
- **The UK accounts for 81.97% of total revenue** (£7.29M), 3,920 of 4,371 customers, and 4.24M of total units sold — this is functionally a UK-domestic business
- **33.9% of customers (1,482) haven't purchased in 90+ days**, and 34.16% of all customers made only one purchase and never returned
- **£608.69K (6.85%) of gross revenue was lost to returns**, bringing revenue from £8.89M gross to £8.28M net
- **£1.33M in revenue sits with At Risk and Lost Customers**; recovering just 20% of it would recapture an estimated £265.73K

*(Full findings, root causes, business insights, and recommendations are written directly into Section 8 of the SQL script.)*

## 🔍 Root Causes

- **Revenue concentration is structural** — the same 80/20 pattern shows up independently in customers, products, and countries, pointing to one underlying dynamic rather than three coincidences
- **December 2011 figures are misleading** — the dataset ends 9 December 2011, so December is a partial month; its apparent -55.27% revenue decline and drop in new customers are data-boundary artifacts, not real trend signals
- **The #1 product ranking is driven by a single bulk order** — Customer 16446 purchased and then fully cancelled 80,995 units of one SKU, distorting the "top product" ranking; a near-identical pattern was found for Customer 12346 on a second product
- **The USA's 57.93% return rate is a small-sample effect** — total USA volume is only 2,458 units, so a handful of returns produces a dramatic-looking percentage
- **The "Lost" population reflects a retention gap, not a reach gap** — repeat customers already generate 93.09% of revenue when retained; the business's weak point is converting a first purchase into a second one

## 💡 Business Insights

- This business is effectively **UK-domestic**, with no meaningfully diversified international market to absorb a UK demand shock
- Financial health is **tightly coupled to ~1,000 Champions**, not the customer base broadly
- **The second purchase, not the first, is the highest-leverage conversion point** in the customer lifecycle
- **Raw Top-10 rankings can mislead** without checking for bulk-and-cancel transactions or small-market statistical noise
- **Returns are concentrated, not systemic** — most of the catalog already performs well

## ✅ Recommendations & Actions

| # | Recommendation |
|---|---|
| 1 | Build a formal retention program for the Champions segment |
| 2 | Launch a time-boxed second-purchase campaign for first-time buyers |
| 3 | Run a value-prioritized win-back campaign for At Risk customers, targeting the £265.73K recovery estimate |
| 4 | Filter bulk-and-cancel transactions out of product performance reporting |
| 5 | Flag or exclude partial-month December data in any trend analysis |
| 6 | Manually review the small number of USA transactions before treating the return rate as a quality issue |
| 7 | Formalize the "Jumbo Bag" product pairing into a cross-sell bundle |
| 8 | Prioritize Switzerland and Belgium for cautious international growth, given their existing 75% repeat-customer rates |

---

## 🔗 Related Projects

Part of a multi-tool Customer Segmentation portfolio, and a broader retail analytics portfolio built on two different datasets:

| Project | Tool | Link |
|---|---|---|
| 📊 Customer Segmentation Analysis (this same dataset) | Power BI | *[Add link here]* |
| 📗 Customer Segmentation Analysis (this same dataset) | Excel | *[Add link here]* |
| 📊 Superstore Sales Analysis | Power BI | [github.com/shivanand-Mathapati-Analyst/powerbi-sales-analysis-project](https://github.com/shivanand-Mathapati-Analyst/powerbi-sales-analysis-project) |
| 🗄️ Superstore Sales Analysis | SQL Server | [github.com/shivanand-Mathapati-Analyst/sql-superstore-sales-analysis](https://github.com/shivanand-Mathapati-Analyst/sql-superstore-sales-analysis) |
| 📗 Superstore Sales Analysis | Excel | [github.com/shivanand-Mathapati-Analyst/excel-superstore-sales-analysis](https://github.com/shivanand-Mathapati-Analyst/excel-superstore-sales-analysis) |

---

## 👤 Author

**Shivanand S. Mathapati**

- 🌐 Portfolio: [shivanand-mathapati.vercel.app](https://shivanand-mathapati.vercel.app)
- 💼 LinkedIn: *[Add link here]*
- 📺 YouTube: *[Add link here]*
- ✉️ Email: *[Add email here]*

---

## 📄 License

This project is licensed under the **MIT License** — feel free to use the schema design and query patterns for your own learning or portfolio projects. See [LICENSE](LICENSE) for details.

---

<div align="center">

⭐ **If you found this project useful, consider giving it a star!** ⭐

</div>
