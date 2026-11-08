
/*===================================================================================
  PROJECT      : CUSTOMER SEGMENTATION USING RFM ANALYSIS (SQL SERVER)
  AUTHOR       : Shivanand S Mathapati
  DATASET      : Online Retail Dataset (UK-based online retailer, transactional data) (Kaggle)
  TOOLS USED   : Microsoft SQL Server, T-SQL (JOINS, CTEs, Views, Stored Procedures, Window Functions, Subqueries)


  SCRIPT STRUCTURE
  -----------------------------------------------------------------------------------
    1. Database & schema setup
    2. Data validation
    3. Data cleaning
    4. Dimension Tables Creation &  modeling
    5. Referential integrity (PK & FK)
    6. RFM calculation
    7. Analysis & reporting layer (Views & Standalone)
		- Executive Overview Analysis
		- Customer Segmentation Analysis
		- Customers Analysis
		- Product Analysis
		- Geographical Analysis
		- Time Based / Seasonality Analysis
        - Stored procedures (Customer level & Segment level)
	8. Data Cleaning Summary, Key Findindings, Root Causes, Business Recommendations & Actions

  =================================================================================== */



--===================================================================================
-- SECTION 1: DATABASE & SCHEMA SETUP
--===================================================================================

-- CREATING DATABASE
CREATE DATABASE CustomerSegmentDB;
GO


-- USING DATABASE
USE CustomerSegmentDB;
GO


-- CREATING SCHEMAS
CREATE SCHEMA original; -- To keep original & copied table
GO

CREATE SCHEMA datamodel; -- To keep all Fact & Dimension tables
GO

CREATE SCHEMA reporting; -- To keep all created views and stored procedures
GO


-- BEFORE IMPORTING DATASETS, IF DATASET IS IN .xlsx FORMAT WE HAVE TO CONVERT THAT INTO .csv FORMAT
-- So 'Online Retail' dataset is in xlsx format i am converting that into csv format using MS Excel

-- IMPORTING DATASET FILE USING IMPORT FLAT FILE OPTION


-- DESCRIBING TABLE 'ONLINERETAIL'
EXEC sp_help 'original.OnlineRetail';


-- DISPLAYING TABLE DETAILS 
SELECT * FROM original.OnlineRetail;
-- 541909 records


-- COUNTING NO. OF ROWS PRESENT IN TABLE
SELECT COUNT(*) FROM original.OnlineRetail; -- 541909 Records


-- CREATING DUPLICATE OF ORIGINAL TABLE (FOR CLEANING)
SELECT *
INTO original.OnlineRetailCopy
FROM original.OnlineRetail;


SELECT TOP(100) * FROM original.OnlineRetailCopy;

SELECT
    s.name AS [Schema Name],
    t.name AS [Table Name]
FROM sys.tables t
JOIN sys.schemas s
    ON t.schema_id = s.schema_id
WHERE s.name IN ('original', 'datamodel', 'reporting');



--===================================================================================
-- SECTION 2: DATA VALIDATION
--===================================================================================


-- ROW COUNT CHECK
SELECT
	COUNT(*) AS [Total Rows]
FROM original.OnlineRetailCopy;
-- Total 541909 records


-- NULL/BLANK VALUES CHECK
SELECT
	SUM(CASE WHEN InvoiceNo IS NULL OR LTRIM(RTRIM(InvoiceNo)) = '' THEN 1 ELSE 0 END) AS [Blank InvoiceNo],
	SUM(CASE WHEN StockCode IS NULL OR LTRIM(RTRIM(StockCode)) = '' THEN 1 ELSE 0 END) AS [Blank StockCode],
	SUM(CASE WHEN Description IS NULL OR LTRIM(RTRIM(Description)) = '' THEN 1 ELSE 0 END) AS [Blank Description],
	SUM(CASE WHEN Quantity IS NULL OR LTRIM(RTRIM(Quantity)) = '' THEN 1 ELSE 0 END) AS [Blank Quantity],
	SUM(CASE WHEN InvoiceDate IS NULL OR LTRIM(RTRIM(InvoiceDate)) = '' THEN 1 ELSE 0 END) AS [Blank InvoiceDate],
	SUM(CASE WHEN UnitPrice IS NULL OR LTRIM(RTRIM(UnitPrice)) = '' THEN 1 ELSE 0 END) AS [Blank UnitPrice],
	SUM(CASE WHEN CustomerID IS NULL OR LTRIM(RTRIM(CustomerID)) = '' THEN 1 ELSE 0 END) AS [Blank CustomerID],
	SUM(CASE WHEN Country IS NULL OR LTRIM(RTRIM(Country)) = '' THEN 1 ELSE 0 END) AS [Blank Country]
FROM original.OnlineRetailCopy;
-- Description field has 1454 and CustomerID field has 135080 NULL Records
-- For the RFM Analysis we have to remove Records with no customerID



-- ZERO/NEGATIVE UNIT PRICE CHECK
SELECT
	COUNT(*) AS [Invalid UnitPrice]
FROM original.OnlineRetailCopy
WHERE TRY_CAST(UnitPrice AS DECIMAL(10,2)) <=0; -- Currently UnitPrice is in TEXT format so I converted to decimal.
-- There are 2521 unit price records with zero or negative values



-- CHECKING CANCELLED/RETURN ORDERS
SELECT
	COUNT(*) AS [Cancelled Orders]
FROM original.OnlineRetailCopy
WHERE InvoiceNo LIKE 'C%';
-- There are 9288 orders are cancelled



-- ZERO/NEGATIVE QUANTITY CHECK
SELECT
	COUNT(*) AS [Invalid Quantities]
FROM originaL.OnlineRetailCopy
WHERE TRY_CAST(Quantity AS INT) <=0; -- Currently Quantity column is in TEXT format so I converted to integer.
-- There are 10624 quantities are negative



-- DUPLICATE ROW CHECK
SELECT
	SUM(Duplicate_Rows - 1) AS [Total Duplicates]
FROM(
	SELECT
		COUNT(*) AS [Duplicate Rows]
	FROM original.OnlineRetailCopy
	GROUP BY InvoiceNo, StockCode, Description, Quantity, InvoiceDate, UnitPrice, CustomerID, Country
	HAVING COUNT(*) > 1
	) AS Duplicate_Check;
-- Here we found 5268 rows were duplicate



-- DISTINCT COUNT
SELECT
	COUNT(DISTINCT InvoiceNo) AS [Unique InvoiceNo],
	COUNT(DISTINCT StockCode) AS [Unique StockCode],
	COUNT(DISTINCT CustomerID) AS [Unique CustomerID],
	COUNT(DISTINCT Country) AS [Unique Countries]
FROM original.OnlineRetailCopy;
-- Unique values are InvoiceNo = 25900, StockCode = 3958, CustomerID = 4372, Country = 38




--===================================================================================
-- SECTION 3: DATA CLEANING
--===================================================================================


-- STEP 01 : REMOVING NULL/BLANK CUSTOMERID
SELECT *
INTO original.OnlineRetailClean
FROM original.OnlineRetailCopy
WHERE CustomerID IS NOT NULL;

SELECT
	COUNT(*) AS Total_Rows
FROM original.OnlineRetailClean;
-- 406829 Records



--STEP 02 : REMOVING DUPLICATE RECORDS

-- CHECKING DUPLICATE RECORDS
SELECT
	SUM(Dup_Rows - 1) AS Total_Duplicates
FROM(
	SELECT
		COUNT(*) AS Dup_Rows
	FROM original.OnlineRetailClean
	GROUP BY InvoiceNo, StockCode, Description, Quantity, InvoiceDate, UnitPrice, CustomerID, Country
	HAVING COUNT(*) > 1
	) AS Dup_Check;
-- After removing NULL CustomerID I found 5225 Duplicate records



-- REMOVING DUPLICATE RECORDS
WITH Dup_Rank AS (
	SELECT
		*,
		ROW_NUMBER() OVER(PARTITION BY InvoiceNo, StockCode, Description, Quantity,
							InvoiceDate, UnitPrice, CustomerID, Country
							ORDER BY (SELECT NULL)) AS Rnk
	FROM original.OnlineRetailClean
)
DELETE FROM Dup_Rank
WHERE Rnk > 1;


SELECT
	COUNT(*) AS Total_Rows
FROM original.OnlineRetailClean;
-- 401604 Records



-- STEP 03 : CHANGING COLUMN DATA TYPES

-- Before changing data types directly i would like check the correctness
SELECT
	COUNT(*) AS [Error Occured]
FROM original.OnlineRetailClean
WHERE Quantity IS NOT NULL AND TRY_CAST(Quantity AS INT) IS NULL;

SELECT
	COUNT(*) AS [Error Occured]
FROM original.OnlineRetailClean
WHERE InvoiceDate IS NOT NULL AND TRY_CAST(InvoiceDate AS DATETIME) IS NULL;

SELECT
	COUNT(*) AS [Error Occured]
FROM original.OnlineRetailClean
WHERE UnitPrice IS NOT NULL AND TRY_CAST(UnitPrice AS DECIMAL(14,2)) IS NULL;

SELECT
	COUNT(*) AS [Error Occured]
FROM original.OnlineRetailClean
WHERE CustomerID IS NOT NULL AND TRY_CAST(CustomerID AS INT) IS NULL;


-- CHANGING COLUMN DATA TYPES
ALTER TABLE Original.OnlineRetailClean ALTER COLUMN Quantity INT;

ALTER TABLE Original.OnlineRetailClean ALTER COLUMN InvoiceDate DATETIME;
-- For this I am getting error so i like do some steps
-- Checking the Date format of the InvoiceDate (UK Dataset - DD-MM-YYYY)

SELECT InvoiceDate FROM original.OnlineRetailClean ORDER BY InvoiceDate;
-- Confirms that it is in DD-MM-YYYY hh:mm Format eg: 31-10-2011 08:56
-- So I am checking System Date Format

SELECT @@LANGUAGE AS Cur_Lang;
-- us_english

-- NOW I AM CHECKING THE ERRORS USING TRY_CONVERT COMMAND WITH STYLE CODE 105- (DD-MM-YYYY), 101- USA(MM/DD/YYYY), 103- UK(DD/MM/YYYY)
SELECT
	COUNT(*) AS [Error Occured]
FROM original.OnlineRetailClean
WHERE InvoiceDate IS NOT NULL AND TRY_CONVERT(DATETIME, InvoiceDate, 105) IS NULL;
-- Zero errors 

-- NOW WE HAVE TO SET DAT FORMAT TO DD-MM-YY

-- SQL SERVER HAS SESSION LEVEL COMMAND TO CONVERT DATE FORMAT i.e. "SET DATEFORMAT"
SET DATEFORMAT dmy;

ALTER TABLE Original.OnlineRetailClean ALTER COLUMN InvoiceDate DATETIME;

ALTER TABLE Original.OnlineRetailClean ALTER COLUMN UnitPrice DECIMAL(14, 2);

ALTER TABLE Original.OnlineRetailClean ALTER COLUMN CustomerID INT;



-- DESCRIBE THE TABLE
EXEC sp_help 'original.OnlineRetailClean';



-- STEP 04 : ADDING FLAG COLUMNS

-- CHECKING CANCELLED/RETURN ORDERS
SELECT
	COUNT(*) AS [Cancelled Orders]
FROM original.OnlineRetailClean
WHERE InvoiceNo LIKE 'C%';
-- 8872 Records


-- ADDING FLAG COLUMN 'ISRETURN' WITH 'YES' & 'NO' VALUES
ALTER TABLE Original.OnlineRetailClean ADD IsReturn NVARCHAR(5);


UPDATE Original.OnlineRetailClean
SET IsReturn = CASE WHEN InvoiceNo LIKE 'C%' THEN 'Yes' ELSE 'No' END;


SELECT COUNT(*) FROM Original.OnlineRetailClean;
-- 401604 Records


-- VERIFY
SELECT 
	IsReturn,
	COUNT(*) AS Orders
FROM original.OnlineRetailClean
GROUP BY IsReturn;



-- STEP 05 : FILTERING INVALID UNIT PRICES

-- ZERO/NEGATIVE UNIT PRICE CHECK
SELECT
	COUNT(*) AS [Invalid UnitPrice]
FROM original.OnlineRetailClean
WHERE UnitPrice <= 0;
-- 44 Records


-- REMOVING THOSE 44 RECORDS
DELETE FROM original.OnlineRetailClean
WHERE UnitPrice <= 0;


-- AFTER CLEANING FINAL NO. OF RECORDS
SELECT COUNT(*) FROM original.OnlineRetailClean;
-- 401560 Total Records



-- STEP 06 : ADDING 'SALES AMOUNT' COLUMN
ALTER TABLE original.OnlineRetailClean
ADD SalesAmount DECIMAL(14, 2);


UPDATE original.OnlineRetailClean
SET SalesAmount = UnitPrice * Quantity;


SELECT * FROM original.OnlineRetailClean ORDER BY Quantity ASC;


-- AFTER CLEANING DISTINCT COUNT
SELECT
	COUNT(DISTINCT InvoiceNo) AS [Unique InvoiceNo],
	COUNT(DISTINCT StockCode) AS [Unique StockCode],
	COUNT(DISTINCT CustomerID) AS [Unique CustomerID],
	COUNT(DISTINCT Country) AS [Unique Countries]
FROM original.OnlineRetailClean;
-- Unique values are InvoiceNo = 22186, StockCode = 3683, CustomerID = 4371, Country = 37



--===================================================================================
-- SECTION 4: CREATING DIMENSION TABLES
--===================================================================================


-- CREATING PRODUCT DIMENSION TABLE
SELECT COUNT(DISTINCT StockCode) AS [Total Products] FROM original.OnlineRetailClean;


-- DIM_PRODUCTS TABLE
WITH ProductCount AS (
	SELECT
		StockCode,
		Description,
		COUNT(*) AS [Desc Count]
	FROM original.OnlineRetailClean
	GROUP BY StockCode, Description
),
ProdRank AS (
	SELECT
		StockCode,
		Description,
		ROW_NUMBER() OVER(PARTITION BY StockCode ORDER BY [Desc Count] DESC) AS Rnk
	FROM ProductCount
)
SELECT
	StockCode,
	Description
INTO datamodel.Dim_Products
FROM ProdRank
WHERE Rnk = 1;


SELECT COUNT(*) AS [Total Products] FROM datamodel.Dim_Products;
-- 3683 Products



--CREATING CUSTOMERS DIMENSION TABLE
SELECT COUNT(DISTINCT CustomerID) AS [Total Customers] FROM original.OnlineRetailClean;


-- DIM_CUSTOMERS TABLE
WITH CountryGroup AS (
	SELECT
		CustomerID,
		Country,
		COUNT(*) AS [Country Count]
	FROM original.OnlineRetailClean
	GROUP BY CustomerID, Country
),
CountryRanked AS (
	SELECT
		CustomerID,
		Country,
		ROW_NUMBER() OVER(PARTITION BY CustomerID ORDER BY [Country Count] DESC) AS Rnk
	FROM CountryGroup
)
SELECT
	CustomerID,
	Country AS CustomersCountry
INTO datamodel.Dim_Customers
FROM CountryRanked
WHERE Rnk = 1;


SELECT COUNT(*) AS [Total Customers] FROM datamodel.Dim_Customers;
-- 4371 Customers



-- CREATING DIMENSION TABLE FOR COUNTRIES
SELECT COUNT(DISTINCT Country) AS [Total Countries] FROM original.OnlineRetailClean;


-- DIM_COUNTRY TABLE
SELECT
	DISTINCT Country
INTO datamodel.Dim_Countries
FROM original.OnlineRetailClean;

SELECT COUNT(*) AS [Total Countries] FROM datamodel.Dim_Countries;
-- 37 Countries



-- CREATING ONE FACT TABLE THAT INCLUDES ONLY MEASURES COLUMN
SELECT
	InvoiceNo,
	StockCode,
	Quantity,
	InvoiceDate,
	UnitPrice,
	CustomerID,
	IsReturn,
	SalesAmount,
	CAST(InvoiceDate AS DATE) AS OrderDate
INTO datamodel.Fact_Sales
FROM original.OnlineRetailClean;

SELECT COUNT(*) AS [Total Records] FROM datamodel.Fact_Sales;
-- 401560 Records



-- CREATING ONE DIMENSION TABLE THAT CONTAINS ONLY DATE RELATED DATA
WITH Calendar AS (
SELECT
    CAST(MIN(InvoiceDate) AS DATE) AS StartDate,
    CAST(MAX(InvoiceDate) AS DATE) AS EndDate
FROM original.OnlineRetailClean
),
Date_Build AS (
	SELECT StartDate
	FROM Calendar

	UNION ALL

	SELECT
		DATEADD(DAY, 1, StartDate)
	FROM Date_Build
	WHERE StartDate < ( SELECT
							EndDate
						FROM Calendar)
)
SELECT
	StartDate AS [Date],
	YEAR(StartDate) AS [Year],
	MONTH(StartDate) AS [Month],
	DATENAME(MONTH, StartDate) AS [MonthName],
	DATEPART(QUARTER, StartDate) AS [Quarter],
	DAY(StartDate) AS [Day],
	DATENAME(WEEKDAY, StartDate) AS [DayName],
	CASE
		WHEN DATENAME(WEEKDAY, StartDate) IN ('Saturday', 'Sunday') THEN 'Yes' ELSE 'No'
	END AS [IsWeekend]
INTO datamodel.Dim_Date
FROM Date_Build
OPTION (MAXRECURSION 0); -- Since we are using recursive CTE it limit 100 levels, so this option removes default 100 limit
-- For a daily calendar, one level equals one day, so the default would stop after 100 days.


SELECT * FROM datamodel.Dim_Date;
-- 374 DAYS


SELECT
	MIN(Date) AS MinDate,
	MAX(Date) AS MaxDate
FROM datamodel.Dim_Date;


SELECT
	CAST(MIN(InvoiceDate) AS DATE) AS MinDate,
	CAST(MAX(InvoiceDate) AS DATE) AS MaxDate
FROM original.OnlineRetailClean;






--===================================================================================
-- SECTION 5: SETTING PRIMARY AND FOREIGN KEYS TO DIMENSION TABLES
--===================================================================================


ALTER TABLE datamodel.Dim_Products
ALTER COLUMN StockCode NVARCHAR(50) NOT NULL;

ALTER TABLE datamodel.Dim_Products
ADD CONSTRAINT PK_Dim_Products_StockCode PRIMARY KEY (StockCode);

ALTER TABLE datamodel.Fact_Sales
ADD CONSTRAINT FK_Fact_Sales_StockCode FOREIGN KEY (StockCode) REFERENCES datamodel.Dim_Products(StockCode);


ALTER TABLE datamodel.Dim_Customers
ALTER COLUMN CustomerID INT NOT NULL;

ALTER TABLE datamodel.Dim_Customers
ADD CONSTRAINT PK_Dim_Customers_CustomerID PRIMARY KEY (CustomerID);

ALTER TABLE datamodel.Fact_Sales
ADD CONSTRAINT FK_Fact_Sales_CustomerID FOREIGN KEY (CustomerID) REFERENCES datamodel.Dim_Customers(CustomerID);


ALTER TABLE datamodel.Dim_Date
ALTER COLUMN Date DATE NOT NULL;

ALTER TABLE datamodel.Dim_Date
ADD CONSTRAINT PK_Dim_Date_Date PRIMARY KEY (Date);

ALTER TABLE datamodel.Fact_Sales
ADD CONSTRAINT FK_Fact_Sales_OrderDate FOREIGN KEY (OrderDate) REFERENCES datamodel.Dim_Date(Date);


ALTER TABLE datamodel.Dim_Countries
ALTER COLUMN Country NVARCHAR(100) NOT NULL;

ALTER TABLE datamodel.Dim_Countries
ADD CONSTRAINT PK_Dim_Countries_Country PRIMARY KEY (Country)

ALTER TABLE datamodel.Dim_Customers
ADD CONSTRAINT FK_Dim_Customers_CustomerCountry FOREIGN KEY (CustomersCountry) REFERENCES datamodel.Dim_Countries(Country)





--===================================================================================
-- SECTION 6: RFM CALCULATIONS
--===================================================================================


/* ==========================================================
	Recency — days since their last purchase (lower = better)
	Frequency — number of distinct orders placed (higher = better)
	Monetary — total amount spent (higher = better)
  ===========================================================*/

-- ADDING MONETARY COLUMN
ALTER TABLE datamodel.Dim_Customers ADD Monetary DECIMAL(14, 2);


UPDATE datamodel.Dim_Customers
SET Monetary = COALESCE(fs.Total_Sales, 0)
FROM datamodel.Dim_Customers dc
LEFT JOIN (
    SELECT
        CustomerID,
        SUM(SalesAmount) AS [Total Sales]
    FROM datamodel.Fact_Sales
    WHERE IsReturn = 'No'
    GROUP BY CustomerID
) AS fs
ON dc.CustomerID = fs.CustomerID;


SELECT COUNT(*) AS [Customers With No Monetary]
FROM datamodel.Dim_Customers
WHERE Monetary IS NULL;
-- 0 Records



-- ADDING RECENCY COLUMN
-- Recency = Number of days since the customer's last purchase
ALTER TABLE datamodel.Dim_Customers ADD Recency INT;

WITH SomeDates AS (
	SELECT
		DATEADD(DAY, 1, MAX(OrderDate)) AS TodayDate,
		DATEDIFF(DAY, MIN(OrderDate), DATEADD(DAY, 1, MAX(OrderDate))) AS MaxDate
	FROM datamodel.Fact_Sales
),
LastPurchasing AS (
	SELECT
		CustomerID,
		MAX(OrderDate) AS LastPurchase
	FROM datamodel.Fact_Sales
	WHERE IsReturn = 'No'
	GROUP BY CustomerID
)
UPDATE datamodel.Dim_Customers
SET Recency = COALESCE ( DATEDIFF(DAY, lp.LastPurchase, sd.TodayDate), sd.MaxDate)
FROM datamodel.Dim_Customers dc
LEFT JOIN LastPurchasing lp
	ON dc.CustomerID = lp.CustomerID
CROSS JOIN SomeDates sd;



SELECT COUNT(*) AS Customers_With_No_Recency
FROM datamodel.Dim_Customers
WHERE Recency IS NULL;
-- 0 Records

SELECT
    MIN(Recency) AS MinRecency, -- 1
    MAX(Recency) AS MaxRecency, -- 374
    ROUND(AVG(CAST(Recency AS FLOAT)),2) AS AvgRecency -- 95.18
FROM datamodel.Dim_Customers;



-- ADDING FREQUENCY COLUMN
-- Frequency — number of distinct orders placed
ALTER TABLE datamodel.Dim_Customers ADD Frequency INT;


UPDATE datamodel.Dim_Customers
SET Frequency = COALESCE(fs.Total_Orders, 0)
FROM datamodel.Dim_Customers dc
LEFT JOIN (
    SELECT
        CustomerID,
        COUNT(DISTINCT InvoiceNo) AS Total_Orders
    FROM datamodel.Fact_Sales
    WHERE IsReturn = 'No'
    GROUP BY CustomerID
) AS fs
ON dc.CustomerID = fs.CustomerID;


SELECT COUNT(*) AS [Customers With No Frequency]
FROM datamodel.Dim_Customers
WHERE Frequency IS NULL;
-- 0 Records


SELECT
    MIN(Frequency) AS MinOrders, -- 1
    MAX(Frequency) AS MaxOrders, -- 209
    ROUND(AVG(CAST(Frequency AS FLOAT)),2) AS AvgFrequency --4.24
FROM datamodel.Dim_Customers;



-- ADDING R_SCORE, F_SCORE, M_SCORE COLUMNS
ALTER TABLE datamodel.Dim_Customers ADD R_Score INT;
ALTER TABLE datamodel.Dim_Customers ADD F_Score INT;
ALTER TABLE datamodel.Dim_Customers ADD M_Score INT;

WITH RFMScored AS (
	SELECT
		CustomerID,
		NTILE(5) OVER(ORDER BY Recency DESC) AS R_Score,
		NTILE(5) OVER(ORDER BY Frequency ASC) AS F_Score,
		NTILE(5) OVER(ORDER BY Monetary ASC) AS M_Score
	FROM datamodel.Dim_Customers
)
UPDATE dc
SET 
	dc.R_Score = sc.R_Score,
	dc.F_Score = sc.F_Score,
	dc.M_Score = sc.M_Score
FROM datamodel.Dim_Customers dc
INNER JOIN RFMScored sc
	ON dc.CustomerID = sc.CustomerID;

SELECT * FROM datamodel.Dim_Customers;



-- ADDING RFM_SCORE COLUMN
ALTER TABLE datamodel.Dim_Customers ADD RFM_Score INT;

UPDATE datamodel.Dim_Customers 
SET RFM_Score = (R_Score * 100) + (F_Score * 10) + M_Score;



-- ADDING SEGMENT COLUMN
ALTER TABLE datamodel.Dim_Customers ADD Segment NVARCHAR(50);

UPDATE datamodel.Dim_Customers 
SET Segment = 
	CASE
		WHEN R_Score >= 4 AND F_Score >= 4 AND M_Score >= 4 THEN 'Champions'
		WHEN R_Score >= 4 THEN 'Potential Loyalists'
		WHEN R_Score <= 2 AND F_Score <= 2 AND M_Score <=2 THEN 'Lost Customers'
		WHEN R_Score <= 2 THEN 'At Risk'
		ELSE 'Loyal Customers'
	END;


SELECT
    Segment,
    COUNT(*) AS CustomerCount,
    SUM(Monetary) AS SegmentRevenue
FROM datamodel.Dim_Customers
GROUP BY Segment
ORDER BY CustomerCount DESC;




--===================================================================================
-- SECTION 7: DATA ANALYSIS
--===================================================================================


--===================================================================================
-- SECTION 7A: EXECUTIVE OVERVIEW
--===================================================================================

-- EXECUTIVE KPI SCORECARD
CREATE VIEW reporting.vw_ExcecutiveKPIs AS
SELECT
	(SELECT SUM(SalesAmount) FROM datamodel.Fact_Sales WHERE IsReturn = 'No') AS 'Total Revenue',
	(SELECT COUNT(DISTINCT InvoiceNo) FROM datamodel.Fact_Sales WHERE IsReturn = 'No') AS 'Total Orders',
	(SELECT COUNT(DISTINCT CustomerID) FROM datamodel.Fact_Sales) AS 'Total Customers',
	(SELECT SUM(Quantity) FROM datamodel.Fact_Sales WHERE IsReturn = 'No') AS 'Total Quantity Sold',
	(SELECT COUNT(DISTINCT StockCode) FROM datamodel.Fact_Sales) AS 'Total Products',
	(SELECT COUNT(DISTINCT CustomerID) FROM datamodel.Dim_Customers WHERE Recency > 90) AS 'Churned Customers',
	(SELECT
		CAST(
			100.00 * COUNT(DISTINCT CASE WHEN Recency > 90 THEN CustomerID END) / COUNT(*)
		AS DECIMAL(10, 2))
	FROM datamodel.Dim_Customers) AS 'Churn Rate %',
	(SELECT CAST(AVG(Monetary) AS DECIMAL(10, 2)) FROM datamodel.Dim_Customers) AS 'Avg CLV',
	(SELECT COUNT(DISTINCT InvoiceNo) FROM datamodel.Fact_Sales WHERE IsReturn = 'Yes') AS 'Returned Orders',
	(SELECT ABS(SUM(Quantity)) FROM datamodel.Fact_Sales WHERE IsReturn = 'Yes') AS 'Returned Quantities',
	(SELECT 
		CAST(
			SUM(CASE WHEN IsReturn = 'No' THEN SalesAmount END)
			/ NULLIF(COUNT(DISTINCT CASE WHEN IsReturn = 'No' THEN InvoiceNo END), 0)
		AS DECIMAL(10,2)) FROM datamodel.Fact_Sales)  AS 'Average Order Value';


-- KPIs
SELECT * FROM reporting.vw_ExcecutiveKPIs;



-- REVENUE TREND BY MONTH
CREATE VIEW reporting.vw_RevenueTrendbyMonth AS
SELECT
	dd.Year AS 'Year',
	dd.Month AS 'Month',
	dd.MonthName AS 'Month Name',
	SUM(CASE WHEN IsReturn = 'No' THEN fs.salesamount END) AS 'Total Revenue'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Date dd
	ON fs.OrderDate = dd.Date
GROUP BY dd.year, dd.MonthName, dd.Month;



-- REVENUE TREND BY MONTH
SELECT
	Year,
	[Month Name],
	[Total Revenue]
FROM reporting.vw_RevenueTrendbyMonth
ORDER BY Year, Month;



-- TOP COUNTRIES BY REVENUE
CREATE VIEW reporting.vw_TopCountriesbyRevenue AS
SELECT
	dc.CustomersCountry AS 'Country',
	COUNT(DISTINCT dc.CustomerID) AS 'Total Customers',
	SUM(fs.SalesAmount) AS 'Total Revenue'
FROM datamodel.Fact_Sales fs
LEFT JOIN datamodel.Dim_Customers dc
	ON fs.CustomerID = dc.CustomerID
GROUP BY dc.CustomersCountry;



-- TOP 10 COUNTRIES BY REVENUE
SELECT TOP(10)
	Country,
	[Total Revenue]
FROM reporting.vw_TopCountriesbyRevenue
ORDER BY [Total Revenue] DESC;



--BOTTOM 10 COUNTRIES BY REVENUE
SELECT TOP(10)
	Country,
	[Total Revenue]
FROM reporting.vw_TopCountriesbyRevenue
ORDER BY [Total Revenue] ASC;



-- HIGHEST REVENUE GENERATED COUNTRY
SELECT TOP(1)
	Country,
	[Total Revenue]
FROM reporting.vw_TopCountriesbyRevenue
ORDER BY [Total Revenue] DESC;
/* The United Kingdom is the highest revenue-generating country, contributing £6.75M in revenue. */


-- LOWEST REVENUE GENERATED COUNTRY
SELECT TOP(1)
	Country,
	[Total Revenue]
FROM reporting.vw_TopCountriesbyRevenue
ORDER BY [Total Revenue] ASC;
/* Saudi Arabia is the lowest revenue-generating country, contributing only £131.17 in total revenue. */



-- REVENUE RETURN RATE BY MONTH
CREATE VIEW reporting.vw_RevenueReturnRatebyMonth AS
SELECT
	dd.Year AS 'Year',
	dd.Month AS 'Month',
	dd.MonthName AS 'Month Name',
	SUM(CASE WHEN fs.IsReturn = 'Yes' THEN ABS(fs.SalesAmount) END) AS 'Return Revenue',
	SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END) AS 'Total Revenue',
	CAST(
		100 * SUM(CASE WHEN fs.IsReturn = 'Yes' THEN ABS(fs.SalesAmount) END)
		/ 
		NULLIF(SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END), 0)
	AS DECIMAL(10,2)) AS 'Return Rate %'
FROM datamodel.Fact_Sales fs
LEFT JOIN datamodel.Dim_Date dd
	ON fs.OrderDate = dd.Date
GROUP BY dd.Year, dd.Month, dd.MonthName;



-- REVENUE RETURN RATE BY MONTH RATE BY MONTH
SELECT
	Year,
	[Month Name],
	[Return Rate %]
FROM reporting.vw_RevenueReturnRatebyMonth
ORDER BY Year, Month;
/* Revenue return rates remained below 10% for most months, but December 2011 recorded the highest return rate at 33.96%,
	September 2011 had the lowest return rate (2.24%). */



-- REVENUE BY SEGMENT
CREATE VIEW reporting.vw_RevenuebySegment AS
SELECT
	Segment,
	COUNT(DISTINCT CustomerID) AS 'Total Customers',
	SUM(Monetary) AS 'Total Revenue',
	CAST(100 * SUM(Monetary) / SUM(SUM(Monetary)) OVER() AS DECIMAL(10, 2)) AS 'Revenue %'
FROM datamodel.Dim_Customers
WHERE Segment IS NOT NULL
GROUP BY Segment;


SELECT * FROM reporting.vw_RevenueBySegment;



-- HIGHEST REVENUE GENERATED CUSTOMER SEGMENT
SELECT TOP(1)
	Segment,
	[Total Revenue],
	[Revenue %]
FROM reporting.vw_RevenueBySegment
ORDER BY [Total Revenue] DESC;
/* Champions is the highest revenue-generating customer segment, contributing £5.85M and accounting for 65.84% of total revenue. */



-- LOWEST REVENUE GENERATED CUSTOMER SEGMENT
SELECT TOP(1)
	Segment,
	[Total Revenue],
	[Revenue %]
FROM reporting.vw_RevenueBySegment
ORDER BY [Total Revenue] ASC;
/* Lost Customers is the lowest revenue-generating segment, contributing £175.17K and accounting for only 1.97% of total revenue. */



-- ORDERS BY WEEKDAY VS WEEKEND
CREATE VIEW reporting.vw_OrdersbyWeekdayWeekend AS
SELECT
	CASE WHEN dd.IsWeekend = 'Yes' THEN 'Weekend' ELSE 'Weekday' END AS 'Day Type',
	COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo END) AS 'Total Orders',
	COUNT(DISTINCT CASE WHEN fs.IsReturn = 'Yes' THEN fs.InvoiceNo END) AS 'Returned Orders'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Date dd
	ON fs.OrderDate = dd.Date
GROUP BY CASE WHEN dd.IsWeekend = 'Yes' THEN 'Weekend' ELSE 'Weekday' END;



-- ORDERS BY WEEKDAY VS WEEKEND
SELECT
	[Day Type],
	[Total Orders]
FROM reporting.vw_OrdersbyWeekdayWeekend;
/* Weekdays accounted for the majority of orders (16,363),
	indicating significantly higher customer purchasing activity compared to weekends (2,169). */



-- RETURN ORDERS BY WEEKDAY VS WEEKEND
SELECT
	[Day Type],
	[Returned Orders]
FROM reporting.vw_OrdersbyWeekdayWeekend;
/* Weekdays recorded the highest number of returned orders (3,481), compared to just 173 returned orders on weekends. */



-- GROSS REVENUE VS NET REVENUE (IMPACT OF RETURNS)
SELECT
    SUM(CASE WHEN IsReturn = 'No' THEN SalesAmount END) AS 'Gross Revenue',
    SUM(CASE WHEN IsReturn = 'Yes' THEN ABS(SalesAmount) END) AS 'Revenue Lost to Returns',
    SUM(SalesAmount) AS 'Net Revenue',
    CAST(
        100.0 * SUM(CASE WHEN IsReturn = 'Yes' THEN ABS(SalesAmount) END)
        / NULLIF(SUM(CASE WHEN IsReturn = 'No' THEN SalesAmount END), 0)
    AS DECIMAL(10,2)) AS 'Revenue Return Rate %'
FROM datamodel.Fact_Sales;
/* Product returns resulted in a revenue loss of £608.69K, reducing gross revenue from £8.89M to a net revenue of £8.28M,
	with an overall revenue return rate of 6.85%. */




--===================================================================================
-- SECTION 7B: SEGMENTATION ANALYSIS
--===================================================================================


-- CUSTOMER SEGMENTATION KPI SCORECARD
CREATE VIEW reporting.vw_SegmentationKPIs AS
SELECT
	COUNT(DISTINCT CustomerID) AS 'Total Customers',
	COUNT(DISTINCT CASE WHEN Segment = 'Champions' THEN CustomerID END) AS 'Champion Customers',
	COUNT(DISTINCT CASE WHEN Segment = 'At Risk' THEN CustomerID END) AS 'At Risk customers',
	COUNT(DISTINCT CASE WHEN Segment = 'Lost Customers' THEN CustomerID END) AS 'Lost Customers',
	CAST(
		100 * SUM(CASE WHEN Segment = 'Champions' THEN Monetary END)
		/ SUM(Monetary)
	AS DECIMAL(10, 2)) AS 'Champions Revenue %'
FROM datamodel.Dim_Customers;


SELECT * FROM reporting.vw_SegmentationKPIs;



-- SEGMENT DETAILS - TOTAL CUSTOMERS, ORDERS, REVENUE %, AVERAGE RECENCY
CREATE VIEW reporting.vw_SegmentationDetail AS
WITH RecencybySegment AS (
	SELECT
		Segment,
		CAST(AVG(CAST(Recency AS FLOAT)) AS DECIMAL(10, 2)) AS 'Avg Recency'
	FROM datamodel.Dim_Customers
	GROUP BY Segment
)
SELECT
	dc.Segment,
	COUNT(DISTINCT dc.CustomerID) AS 'Total Customers',
	COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo END) AS 'Total Orders',
	CAST(
		100 * SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END)
		/ SUM(SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END)) OVER()
	AS DECIMAL(10,2)) AS 'Revenue %',
	rs.[Avg Recency] AS 'Avg days since last purchase'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Customers dc
	ON fs.CustomerID = dc.CustomerID
INNER JOIN RecencybySegment rs
	ON rs.Segment = dc.Segment
GROUP BY dc.Segment, rs.[Avg Recency];


SELECT * FROM reporting.vw_SegmentationDetail ORDER BY [Total Customers] DESC;



-- SEGEMENT DISTRIBUTION
CREATE VIEW reporting.vw_SegmentDistribution AS
SELECT
	Segment,
	COUNT(CustomerID) AS 'Total Customers',
	SUM(Monetary) AS 'Total Revenue',
	CAST(AVG(CAST(R_Score AS FLOAT)) AS DECIMAL(5, 2)) AS 'Avg R_Score',
	CAST(AVG(CAST(F_Score AS FLOAT)) AS DECIMAL(5, 2)) AS 'Avg F_Score',
	CAST(AVG(CAST(M_Score AS FLOAT)) AS DECIMAL(5, 2)) AS 'Avg M_Score'
FROM datamodel.Dim_Customers
GROUP BY Segment;


SELECT * FROM reporting.vw_SegmentDistribution;



-- RFM HEATMAP (RECENCY * FREQUENCY SCORE)
SELECT
	R_Score,
	F_Score,
	COUNT(CustomerID) AS 'Total Customers'
FROM datamodel.Dim_Customers
GROUP BY R_Score, F_Score;



-- CUSTOMERS VS REVENUE BY SEGMENT
SELECT *
FROM reporting.vw_RevenuebySegment
ORDER BY [Total Revenue] DESC;



-- AVERAGE ORDER VALUE BY SEGMENT
SELECT
    dc.Segment,
    COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo END) AS 'Total Orders',
    SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END) AS 'Total Revenue',
    CAST(
        SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END)
        / NULLIF(COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo END), 0)
    AS DECIMAL(10,2)) AS 'Avg Order Value'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Customers dc
	ON fs.CustomerID = dc.CustomerID
GROUP BY dc.Segment
ORDER BY [Avg Order Value] DESC;
/* Champions had the highest average order value (£543.24), while Lost Customers had the lowest (£223.71). */



-- ORDER VOLUME SHARE BY SEGMENT
SELECT
    dc.Segment,
    COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo END) AS 'Total Orders',
    CAST(
        100.0 * COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo END)
        / SUM(COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo END)) OVER()
    AS DECIMAL(10,2)) AS 'Order %'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Customers dc
	ON fs.CustomerID = dc.CustomerID
GROUP BY dc.Segment
ORDER BY [Order %] DESC;
/* Champions placed the highest number of orders (10,772), accounting for 58.13% of all orders,
	while Lost Customers contributed the lowest share at 4.23%. */



-- RFM SCORE DISTRIBUTION (CUSTOMER COUNT PER COMBINED RFM SCORE)
SELECT
    RFM_Score,
    COUNT(CustomerID) AS 'Total Customers'
FROM datamodel.Dim_Customers
GROUP BY RFM_Score
ORDER BY RFM_Score DESC;
/* Most customers fall into either the highest-value segment (RFM 555) or 
	the lowest-value segment (RFM 111), with over 350 customers in each group. */




--===================================================================================
-- SECTION 7C: CUSTOMERS ANALYSIS
--===================================================================================


-- CUSTOMERS ANALYSIS KPI SCORECARD
CREATE VIEW reporting.vw_CustomersKPIs AS
SELECT
	COUNT(CustomerID) AS 'Total Customers',
	COUNT(CASE WHEN Frequency > 1 THEN CustomerID END) AS 'Repeat Customers',
	CAST(
		100 * COUNT(CASE WHEN Frequency > 1 THEN CustomerID END)
		/ COUNT(CustomerID)
	AS DECIMAL(10, 2)) AS 'Repeat Rate %',
	CAST(AVG(CAST(Recency AS FLOAT)) AS DECIMAL(10, 2)) AS 'Avg Recency',
	CAST(AVG(CAST(Frequency AS FLOAT)) AS DECIMAL(10, 2)) AS 'Avg Frequency',
	CAST(AVG(Monetary) AS DECIMAL(10, 2)) AS 'Avg CLV'
FROM datamodel.Dim_Customers;


SELECT * FROM reporting.vw_CustomersKPIs;


-- TOP CUSTOMERS BY REVENUE
CREATE VIEW reporting.vw_TopCustomersbyRevenue AS
SELECT
	CustomerID,
	Segment,
	Monetary AS 'Total Revenue',
	Recency AS 'Days Since Last Purchase',
	Frequency AS 'Total Orders'
FROM datamodel.Dim_Customers
WHERE Monetary <> 0;

SELECT * FROM reporting.vw_TopCustomersbyRevenue;

/* SELECT
	fs.CustomerID,
	dc.Segment,
	SUM(fs.SalesAmount) AS 'Total Revenue',
	COUNT(DISTINCT fs.InvoiceNo) AS 'Total Orderss',
	dc.Recency AS 'Days Since Last Purchase'
FROM datamodel.Fact_Sales fs
LEFT JOIN datamodel.Dim_Customers dc
ON fs.CustomerID = dc.CustomerID
WHERE fs.IsReturn = 'No'
GROUP BY fs.CustomerID, dc.Segment,dc.Monetary, dc.Recency, dc.Frequency
HAVING SUM(fs.SalesAmount) <> 0
ORDER BY SUM(fs.SalesAmount) DESC;*/


-- TOP 10 CUSTOMERS BY REVENUE
SELECT TOP(10)
	CustomerID,
	[Total Revenue]
FROM reporting.vw_TopCustomersbyRevenue
ORDER BY [Total Revenue] DESC;
/* Customer 14646 is the top revenue-generating customer, contributing £280.21K in total revenue. */



-- BOTTOM 10 CUSTOMERS BY REVENUE
SELECT TOP(10)
	CustomerID,
	[Total Revenue]
FROM reporting.vw_TopCustomersbyRevenue
ORDER BY [Total Revenue] ASC;
/* Customer 16738 generated the lowest revenue (£3.75), while all bottom 10 customers contributed less than £21 in total revenue. */



-- TOP 10 CUSTOMERS BY ORDER VOLUME
SELECT TOP(10)
	CustomerID,
	[Total Orders]
FROM reporting.vw_TopCustomersbyRevenue
ORDER BY [Total Orders] DESC;
/* Customer 12748 is the most active customer, placing 209 orders in total. */



-- BOTTOM 10 CUSTOMERS BY ORDER VOLUME
SELECT TOP(10)
	CustomerID,
	[Total Orders]
FROM reporting.vw_TopCustomersbyRevenue
ORDER BY [Total Orders] ASC;



-- RECENCY DISTRIBUTION (DAYS VS CUSTOMERS)
CREATE VIEW reporting.vwRecencyDistribution AS
WITH RecencyBin AS (
	SELECT
		CustomerID,
		CASE
			WHEN Recency <= 30 THEN '0-30 days'
			WHEN Recency <= 60 THEN '31-60 days'
			WHEN Recency <= 90 THEN '61-90 days'
			ELSE '90+ days'
		END AS 'Recency Group',
			CASE
			WHEN Recency <= 30 THEN 1
			WHEN Recency <= 60 THEN 2
			WHEN Recency <= 90 THEN 3
			ELSE 4
		END AS 'Sort Recency Group'
	FROM datamodel.Dim_Customers
)
SELECT
	[Sort Recency Group],
	[Recency Group],
	COUNT(CustomerID) AS 'Total Customers'
FROM RecencyBin
GROUP BY [Sort Recency Group], [Recency Group];


-- DAYS VS TOTAL CUSTOMERS
SELECT
	[Recency Group],
	[Total Customers]
FROM reporting.vwRecencyDistribution
ORDER BY [Sort Recency Group] ASC;
/* Most customers (1,628) made a purchase within the last 30 days, while 1,482 customers have not purchased for more than 90 days. */



-- CLV DISTRIBUTION
CREATE VIEW reporting.CLVDistribution AS
WITH CLVBin AS (
	SELECT
		CustomerID,
		CASE
			WHEN Monetary <= 500 THEN '£0-£500'
			WHEN Monetary <= 1500 THEN '£501-£1.5K'
			WHEN Monetary <= 3000 THEN '£1.5K-£3K'
			WHEN Monetary <= 6000 THEN '£3K-£6K'
			ELSE '£6K+'
		END AS 'CLV Group',
			CASE
			WHEN Monetary <= 500 THEN 1
			WHEN Monetary <= 1500 THEN 2
			WHEN Monetary <= 3000 THEN 3
			WHEN Monetary <= 6000 THEN 4
			ELSE 5
		END AS 'Sort CLV Group'
	FROM datamodel.Dim_Customers
)
SELECT
	[Sort CLV Group],
	[CLV Group],
	COUNT(CustomerID) AS 'Total Customers'
FROM CLVBin
GROUP BY [Sort CLV Group], [CLV Group];


-- CLV VS TOTAL CUSTOMERS
SELECT
	[CLV Group],
	[Total Customers]
FROM reporting.CLVDistribution
ORDER BY [Sort CLV Group];
/* Most customers (1,798) have a CLV between £0 and £500, while only 209 customers have a CLV greater than £6K. */



-- CUSTOMERS REVENUE 80/20 PARETO ANALYSIS (I am using this view in next query)
CREATE VIEW reporting.CustRevParetoAnalysis AS
SELECT
	CustomerID,
	Monetary,
	RANK() OVER(ORDER BY Monetary DESC) AS 'Revenue Rank',
	SUM(Monetary) OVER (ORDER BY Monetary DESC ROWS UNBOUNDED PRECEDING) AS 'Cumulative Revenue',
	CAST(
		100 * SUM(Monetary) OVER (ORDER BY Monetary DESC ROWS UNBOUNDED PRECEDING)
		/ SUM(Monetary) OVER()
	AS DECIMAL(10, 2)) AS 'Cumulative Revenue %'
FROM datamodel.Dim_Customers;



-- TOP 20% VS REMAINING 80% CUSTOMERS REVENUE ANALYSIS
WITH CustomerCounts AS
(
    SELECT
		COUNT(CustomerID) AS 'Total_Customers'
    FROM reporting.CustRevParetoAnalysis
)
SELECT
    CASE
		WHEN [Revenue Rank] <= CEILING(cc.Total_Customers * 0.20)
		THEN 'Top 20% Customers'
		ELSE 'Remaining 80% Customers'
	END AS 'Customer Group',
    COUNT(CustomerID) AS 'Total Customers',
    SUM(Monetary) AS 'Total Revenue',
    CAST(
        100 *SUM(Monetary)
		/ (SELECT SUM(Monetary) FROM reporting.CustRevParetoAnalysis)
    AS DECIMAL(10,2)) AS 'Revenue %'
FROM reporting.CustRevParetoAnalysis p
CROSS JOIN CustomerCounts cc
GROUP BY 
	CASE
		WHEN [Revenue Rank] <= CEILING(cc.Total_Customers * 0.20)
		THEN 'Top 20% Customers'
		ELSE 'Remaining 80% Customers'
	END;
/* IThe top 20% of customers (875 customers) generated 74.85% of total revenue (£6.65M),
	while the remaining 80% contributed only 25.15%. */



-- ONE-TIME VS REPEAT CUSTOMER REVENUE ANALYSIS
SELECT
    CASE WHEN Frequency = 1 THEN 'One-Time Customers' ELSE 'Repeat Customers' END AS 'Customer Type',
    COUNT(CustomerID) AS 'Total Customers',
    SUM(Monetary) AS 'Total Revenue',
    CAST(
		100.0 * SUM(Monetary)
		/ SUM(SUM(Monetary)) OVER()
	AS DECIMAL(10,2)) AS 'Revenue %'
FROM datamodel.Dim_Customers
GROUP BY CASE WHEN Frequency = 1 THEN 'One-Time Customers' ELSE 'Repeat Customers' END;
/* Repeat customers contributed the majority of revenue (£8.27M, 93.09%), while one-time customers accounted for only 6.91%. */



-- ONE-TIME BUYERS RATE %
SELECT
    CAST(
        100.0 * COUNT(CASE WHEN Frequency = 1 THEN CustomerID END)
		/ COUNT(CustomerID)
    AS DECIMAL(10,2)) AS 'One-Time Buyers Rate %'
FROM datamodel.Dim_Customers;
/* 34.16% of customers made only one purchase and did not return for additional orders. */



-- REVENUE CONTRIBUTION OF TOP 1% CUSTOMERS
WITH RankedCustomers AS (
    SELECT
        CustomerID,
        Monetary,
        NTILE(100) OVER(ORDER BY Monetary DESC) AS Percentile
    FROM datamodel.Dim_Customers
)
SELECT
    COUNT(CustomerID) AS 'Top 1% Customer Count',
    SUM(Monetary) AS 'Their Revenue',
    CAST(
        100.0 * SUM(Monetary)
		/ (SELECT SUM(Monetary) FROM datamodel.Dim_Customers)
    AS DECIMAL(10,2)) AS 'Revenue %'
FROM RankedCustomers
WHERE Percentile = 1;
/* The top 1% of customers (44 customers) generated £2.85M in revenue, contributing 32.06% of total revenue. */




--===================================================================================
-- SECTION 7D: PRODUCT ANALYSIS
--===================================================================================


-- PRODUCT ANALYSIS KPI SCORECARD
CREATE VIEW reporting.vw_PrductsKPIs AS
SELECT
	COUNT(DISTINCT StockCode) AS 'Total Products',
	SUM(CASE WHEN IsReturn = 'No' THEN Quantity END) AS 'Total Quantity Sold',
	SUM(CASE WHEN IsReturn = 'No' THEN Quantity END)
	/ 
	COUNT(DISTINCT CASE WHEN IsReturn = 'No' THEN InvoiceNo END) AS 'Avg Items per Order',
	CAST(
		100.00 * SUM(CASE WHEN IsReturn = 'Yes' THEN ABS(Quantity) END)
		/ SUM(CASE WHEN IsReturn = 'No' THEN Quantity END)
	AS DECIMAL(10, 2)) AS 'Product Return Rate %'
FROM datamodel.Fact_Sales;


SELECT * FROM reporting.vw_PrductsKPIs;



-- TOP 10 PRODUCTS BY REVENUE
SELECT TOP (10)
	dp.Description AS 'Product Name',
	SUM(fs.SalesAmount) AS 'Total Revenue'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Products dp
	ON fs.StockCode = dp.StockCode
WHERE fs.IsReturn = 'No'
GROUP BY dp.Description
ORDER BY [Total Revenue] DESC;
/* PAPER CRAFT, LITTLE BIRDIE was the highest revenue-generating product, contributing £168.47K in total revenue. */



-- BOTTOM 10 PRODUCTS BY REVENUE
SELECT TOP (10)
	dp.Description AS 'Product Name',
	SUM(fs.SalesAmount) AS 'Total Revenue'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Products dp
	ON fs.StockCode = dp.StockCode
WHERE fs.IsReturn = 'No'
GROUP BY dp.Description
ORDER BY [Total Revenue] ASC;



-- TOP 10 PRODUCTS BY QUANTITY SOLD
SELECT TOP (10)
	dp.Description AS 'Product Name',
	SUM(fs.Quantity) AS 'Total Quantity Sold'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Products dp
	ON fs.StockCode = dp.StockCode
WHERE fs.IsReturn = 'No'
GROUP BY dp.Description
ORDER BY [Total Quantity Sold] DESC;
/* PAPER CRAFT, LITTLE BIRDIE was the most purchased product, with 80,995 units sold,
	making it the top-selling product by volume. */



-- TOP 10 MOST RETURNED PRODUCTS
SELECT TOP (10)
	dp.Description AS 'Product Name',
	SUM(ABS(fs.Quantity)) AS 'Total Quantity Sold'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Products dp
	ON fs.StockCode = dp.StockCode
WHERE fs.IsReturn = 'Yes'
GROUP BY dp.Description
ORDER BY [Total Quantity Sold] DESC;
/* PAPER CRAFT, LITTLE BIRDIE recorded the highest number of returns, with 80,995 units returned. */



-- PRODUCT PERFORMANCE DETAILS
CREATE VIEW reporting.vw_ProductPerformance AS
SELECT
	fs.StockCode,
	dp.Description,
	SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount ELSE 0 END) AS 'Total Revenue',
	SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.Quantity ELSE 0 END) AS 'Total Quantity Sold',
	SUM(CASE WHEN fs.IsReturn = 'Yes' THEN ABS(fs.Quantity) ELSE 0 END) AS 'Return Quantity',
	COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.CustomerID END) AS 'How Many Customers Buy this?',
	COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo END) AS 'Total Orders',

	CAST(
		SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount ELSE 0 END)
		/ NULLIF(COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo END), 0)
	AS DECIMAL(10, 2)) AS 'Avg Revenue per Order',

	SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.Quantity ELSE 0 END)
	/ NULLIF(COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo ELSE 0 END), 0) AS 'Avg Quantity per Order',

	CAST(
		100.00 * SUM(CASE WHEN fs.IsReturn = 'Yes' THEN ABS(fs.Quantity) ELSE 0 END)
		/ NULLIF(SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.Quantity ELSE 0 END), 0)
	AS DECIMAL(10, 2)) AS 'Product Return Rate %',

	RANK() OVER(ORDER BY SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount ELSE 0 END) DESC) AS 'Revenue Rank',

	CAST(
		100.00 * SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount ELSE 0 END)
		/ SUM(SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount ELSE 0 END)) OVER()
	AS DECIMAL(10, 2)) AS 'Revenue Contribution %'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Products dp
	ON fs.StockCode = dp.StockCode
GROUP BY fs.StockCode, dp.Description;


SELECT * FROM reporting.vw_ProductPerformance;



-- TOP 20% VS REMAINING 80% PRODUCT REVENUE ANALYSIS
WITH ProductCount AS
(
    SELECT
		COUNT(StockCode) AS 'Total_Products'
    FROM reporting.vw_ProductPerformance
)
SELECT
    CASE
		WHEN [Revenue Rank] <= CEILING(pc.Total_Products * 0.20)
		THEN 'Top 20% Products'
		ELSE 'Remaining 80% Products'
	END AS 'Product Group',
    COUNT(StockCode) AS 'Total Products',
    SUM([Total Revenue]) AS 'Total Revenue',
    CAST(
        100 *SUM([Total Revenue])
		/ (SELECT SUM([Total Revenue]) FROM reporting.vw_ProductPerformance)
    AS DECIMAL(10,2)) AS 'Revenue %'
FROM reporting.vw_ProductPerformance pp
CROSS JOIN ProductCount pc
GROUP BY 
	CASE
		WHEN [Revenue Rank] <= CEILING(pc.Total_Products * 0.20)
		THEN 'Top 20% Products'
		ELSE 'Remaining 80% Products'
	END;
/* The top 20% of products (737 products) generated 78.80% of total revenue (£7.00M),
	while the remaining 80% contributed only 21.20%. */



-- MARKET BASKET ANALYSIS: TOP 10 PRODUCT PAIRS FREQUENTLY PURCHASED TOGETHER
SELECT TOP(10)
    p1.Description AS 'Product A',
    p2.Description AS 'Product B',
    COUNT(DISTINCT a.InvoiceNo) AS 'Times Bought Together'
FROM datamodel.Fact_Sales a
INNER JOIN datamodel.Fact_Sales b
    ON a.InvoiceNo = b.InvoiceNo AND a.StockCode < b.StockCode
INNER JOIN datamodel.Dim_Products p1
	ON a.StockCode = p1.StockCode
INNER JOIN datamodel.Dim_Products p2
	ON b.StockCode = p2.StockCode
WHERE a.IsReturn = 'No' AND b.IsReturn = 'No'
GROUP BY p1.Description, p2.Description
ORDER BY [Times Bought Together] DESC;
/* JUMBO BAG PINK POLKADOT and JUMBO BAG RED RETROSPOT were the most frequently purchased together,
	appearing in 546 orders, suggesting a strong cross-selling opportunity. */



-- PRODUCT RETURN RATE ANALYSIS
WITH ReturnGroup AS (
	SELECT
		*,
		CASE
			WHEN [Product Return Rate %] = 0 THEN 'No Returns'
			WHEN [Product Return Rate %] <= 5 THEN 'Low Returns (<=5%)'
			WHEN [Product Return Rate %] <= 15 THEN 'Moderate Returns (5-15%)'
			ELSE 'High Returns (>15%)'
		END AS 'Return Type',
		CASE
			WHEN [Product Return Rate %] = 0 THEN 1
			WHEN [Product Return Rate %] <= 5 THEN 2
			WHEN [Product Return Rate %] <= 15 THEN 3
			ELSE 4
		END AS 'Sort Return Type'
	FROM reporting.vw_ProductPerformance
)
SELECT
	[Return Type],
	COUNT(StockCode) AS 'Total Products',
    SUM([Total Revenue]) AS 'Total Revenue'
FROM ReturnGroup
GROUP BY [Sort Return Type], [Return Type]
ORDER BY [Sort Return Type] ASC;
/* The majority of revenue (£6.55M) comes from products with low return rates (≤5%),
	while only 177 products fall into the high-return category (>15%). */



-- PRODUCT PERFORMANCE BY PRICE GROUP/TIER
WITH ProductPriceGroup AS (
SELECT
	*,
    CASE
        WHEN UnitPrice < 1 THEN 'Under £1'
        WHEN UnitPrice < 5 THEN '£1 - £5'
        WHEN UnitPrice < 10 THEN '£5 - £10'
        WHEN UnitPrice < 20 THEN '£10 - £20'
        ELSE '£20+'
    END AS 'Price Group',
	CASE
        WHEN UnitPrice < 1 THEN 1
        WHEN UnitPrice < 5 THEN 2
        WHEN UnitPrice < 10 THEN 3
        WHEN UnitPrice < 20 THEN 4
        ELSE 5
    END AS 'Sort Price Group'
FROM datamodel.Fact_Sales
)
SELECT
	[Price Group],
	COUNT(DISTINCT StockCode) AS 'Total Products',
    SUM(CASE WHEN IsReturn = 'No' THEN SalesAmount END) AS 'Total Revenue',
    SUM(CASE WHEN IsReturn = 'No' THEN Quantity END) AS 'Total Quantity Sold'
FROM ProductPriceGroup
GROUP BY [Sort Price Group], [Price Group]
HAVING SUM(CASE WHEN IsReturn = 'No' THEN SalesAmount END) > 0
ORDER BY [Sort Price Group] ASC;
/* Products priced between £1 and £5 generated the highest revenue (£6.04M) and
	ordrs volume (2.92M units), making this the strongest-performing price group. */




--===================================================================================
-- SECTION 7E: GEOGRAPHICAL ANALYSIS
--===================================================================================


-- GEOGRAPHICAL PERFORMANCE ANALYSIS
CREATE VIEW reporting.vw_CountryPerformance AS
SELECT
	dc.CustomersCountry,
	SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount ELSE 0 END) AS 'Total Revenue',
	CAST(
		100.00 * SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount ELSE 0 END)
		/ SUM(SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount ELSE 0 END)) OVER()
	AS DECIMAL(10, 2)) AS 'Revenue Contribution %',

	RANK() OVER(ORDER BY SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount ELSE 0 END) DESC) AS 'Revenue Rank',

	COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.CustomerID END) AS 'Total Customers',
	COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo END) AS 'Total Orders',
	SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.Quantity END) AS 'Total Quantity Sold',
	SUM(CASE WHEN fs.IsReturn = 'Yes' THEN ABS(fs.Quantity) ELSE 0 END) AS 'Return Quantity',

	CAST(
		100.00 * SUM(CASE WHEN fs.IsReturn = 'Yes' THEN ABS(fs.Quantity) END)
		/ NULLIF(SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.Quantity END), 0)
	AS DECIMAL(10, 2)) AS 'Return Rate %',

	
	CAST(
		SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount ELSE 0 END)
		/ NULLIF(COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.CustomerID END), 0)
	AS DECIMAL(10, 2)) AS 'Avg Revenue per Customer',

	CAST(
		SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount ELSE 0 END)
		/ NULLIF(COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo END), 0)
	AS DECIMAL(10, 2)) AS 'Avg Order Value'
	
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Customers dc
	ON fs.CustomerID = dc.CustomerID
GROUP BY dc.CustomersCountry;


SELECT * FROM reporting.vw_CountryPerformance;



-- TOP 10 COUNTRIES BY REVENUE
SELECT TOP(10)
	CustomersCountry AS 'Country',
	[Total Revenue]
FROM reporting.vw_CountryPerformance
ORDER BY [Total Revenue] DESC;
/* The United Kingdom was the top-performing country with £7.29M in revenue. */




-- BOTTOM 10 COUNTRIES BY REVENUE
SELECT TOP(10)
	CustomersCountry AS 'Country',
	[Total Revenue]
FROM reporting.vw_CountryPerformance
ORDER BY [Total Revenue] ASC;
/* Saudi Arabia was the lowest revenue-generating country, contributing only £145.92 in total revenue. */



-- TOP 10 COUNTRIES BY QUANTITY SOLD
SELECT TOP(10)
	CustomersCountry AS 'Country',
	[Total Quantity Sold]
FROM reporting.vw_CountryPerformance
ORDER BY [Total Quantity Sold] DESC;
/* The United Kingdom was the largest market by quantity sold, accounting for 4.24M units,
	while the Netherlands ranked second with 200K units. */



-- BOTTOM 10 COUNTRIES BY QUANTITY SOLD
SELECT TOP(10)
	CustomersCountry AS 'Country',
	[Total Quantity Sold]
FROM reporting.vw_CountryPerformance
ORDER BY [Total Quantity Sold] ASC;
/* Saudi Arabia recorded the lowest sales volume with just 80 units sold,
	while all bottom 10 countries sold fewer than 1,000 units each. */



-- TOP 10 COUNTRIES BY TOTAL CUSTOMERS
SELECT TOP(10)
	CustomersCountry AS 'Country',
	[Total Customers]
FROM reporting.vw_CountryPerformance
ORDER BY [Total Customers] DESC;
/* The United Kingdom dominates the customer base with 3,920 customers,
	while Germany ranks second with only 94 customers. */



-- BOTTOM 10 COUNTRIES BY TOTAL CUSTOMERS
SELECT TOP(10)
	CustomersCountry AS 'Country',
	[Total Customers]
FROM reporting.vw_CountryPerformance
ORDER BY [Total Customers] ASC;
/* Brazil, RSA, Iceland, Lithuania, European Community, Lebanon, Saudi Arabia,
	Singapore, and Czech Republic each had only one customer. */



-- TOP 10 COUNTRIES BY TOTAL ORDERS
SELECT TOP(10)
	CustomersCountry AS 'Country',
	[Total Orders]
FROM reporting.vw_CountryPerformance
ORDER BY [Total Orders] DESC;
/* The United Kingdom led all countries with 16,646 orders, while Germany ranked second with 457 orders. */



-- BOTTOM 10 COUNTRIES BY TOTAL ORDERS
SELECT TOP(10)
	CustomersCountry AS 'Country',
	[Total Orders]
FROM reporting.vw_CountryPerformance
ORDER BY [Total Orders] ASC;
/* Brazil, Lebanon, RSA, and Saudi Arabia recorded only one order each,
	while all bottom 10 countries had five or fewer orders, indicating very low order activity in these markets. */



-- TOP 10 COUNTRIES WERE MOST QUANTITIES ARE RETURNED
SELECT TOP(10)
	CustomersCountry AS 'Country',
	[Return Quantity]
FROM reporting.vw_CountryPerformance
ORDER BY [Return Quantity] DESC;
/* The United Kingdom had the highest number of returned units (259,167),
	while EIRE ranked a distant second with 4,196 returned units. */



-- TOP 10 COUNTRIES BY QUANTITY RETURN RATE
SELECT TOP(10)
	CustomersCountry AS 'Country',
	[Return Rate %]
FROM reporting.vw_CountryPerformance
ORDER BY [Return Rate %] DESC;
/* The USA had the highest return rate at 57.93%, while most major markets maintained return rates below 12%. */



-- TOP 20% VS REMAINING 80% COUNTRIES REVENUE ANALYSIS
WITH CountryCount AS
(
    SELECT
		COUNT(CustomersCountry) AS 'Total_Countries'
    FROM reporting.vw_CountryPerformance
)
SELECT
    CASE
		WHEN [Revenue Rank] <= CEILING(cc.Total_Countries * 0.20)
		THEN 'Top 20% Countries'
		ELSE 'Remaining 80% Countries'
	END AS 'Country Group',
    COUNT(CustomersCountry) AS 'Total Countries',
    SUM([Total Revenue]) AS 'Total Revenue',
    CAST(
        100 *SUM([Total Revenue])
		/ (SELECT SUM([Total Revenue]) FROM reporting.vw_CountryPerformance)
    AS DECIMAL(10,2)) AS 'Revenue %'
FROM reporting.vw_CountryPerformance cp
CROSS JOIN CountryCount cc
GROUP BY 
	CASE
		WHEN [Revenue Rank] <= CEILING(cc.Total_Countries * 0.20)
		THEN 'Top 20% Countries'
		ELSE 'Remaining 80% Countries'
	END
ORDER BY [Total Revenue] DESC;
/* Revenue is highly concentrated across countries, with the top 20% of countries (8 countries) generating 96.01% of total revenue (£8.53M),
	while the remaining 29 countries contributed only 3.99%. */



-- DOMESTIC (UK) VS FOREIGN REVENUE ANALYSIS
-- (Here we considered UK as a domestic country.)
SELECT
    CASE WHEN dc.CustomersCountry = 'United Kingdom' THEN 'Domestic (UK)' ELSE 'Foreign' END AS 'Market Type ',
    COUNT(DISTINCT dc.CustomerID) AS 'Total Customers',
    SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END) AS 'Total Revenue',
    CAST(
        100.0 * SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END)
        / SUM(SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END)) OVER()
    AS DECIMAL(10,2)) AS 'Revenue %'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Customers dc
	ON fs.CustomerID = dc.CustomerID
GROUP BY CASE WHEN dc.CustomersCountry = 'United Kingdom' THEN 'Domestic (UK)' ELSE 'Foreign' END;
/* The UK market contributed 81.97% of total revenue (£7.29M), while all foreign markets combined accounted for only 18.03%. */



-- COUNTRY-WISE REPEAT CUSTOMER RATE
SELECT
    CustomersCountry AS 'Country',
    COUNT(CustomerID) AS 'Total Customers',
    COUNT(CASE WHEN Frequency > 1 THEN CustomerID END) AS 'Repeat Customers',
    CAST(
        100.0 * COUNT(CASE WHEN Frequency > 1 THEN CustomerID END) / COUNT(CustomerID)
    AS DECIMAL(10,2)) AS 'Repeat Rate %'
FROM datamodel.Dim_Customers
GROUP BY CustomersCountry
HAVING COUNT(CustomerID) >= 10  -- Excludes countries with fewer than 10 customers to ensure meaningful repeat-rate analysis.
ORDER BY [Repeat Rate %] DESC;
/* Switzerland and Belgium recorded the highest repeat customer rate (75%),
	indicating stronger customer retention compared to other major markets. */
GO




--===================================================================================
-- SECTION 7F: TIME ANALYSIS
--===================================================================================


-- TIME BASED SALES PERFORMANCE ANALYSIS
CREATE VIEW reporting.vw_TimePerformance AS
SELECT
	dd.Year,
	dd.Quarter,
	dd.Month,
	dd.MonthName,
	SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END) AS 'Total Revenue',
	CAST(
		100.00 * SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END)
		/ SUM(SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END)) OVER()
	AS DECIMAL(10, 2)) AS 'Revenue %',

	COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.CustomerID END) AS 'Total Customers',
	COUNT(DISTINCT fs.CustomerID) AS 'Total Customers(incl. Return Customers)',
	COUNT(DISTINCT CASE WHEN fs.IsReturn = 'No' THEN fs.InvoiceNo END) AS 'Total Orders',
	SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.Quantity END) AS 'Total Quantity Sold',
	SUM(CASE WHEN fs.IsReturn = 'Yes' THEN ABS(fs.Quantity) END) AS 'Returned Quantity',
	CAST(
		100.00 * SUM(CASE WHEN fs.IsReturn = 'Yes' THEN ABS(fs.Quantity) END)
		/ NULLIF(SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.Quantity END), 0)
	AS DECIMAL(10, 2)) AS 'Return Rate %'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Date dd
ON fs.OrderDate = dd.Date
GROUP BY dd.Year, dd.Quarter, dd.Month, dd.MonthName;
GO


SELECT * FROM reporting.vw_TimePerformance;


-- TOTAL CUSTOMERS BY MONTH
SELECT
	Year,
	Month,
	MonthName,
	[Total Customers]
FROM reporting.vw_TimePerformance
ORDER BY YEAR ASC, Month ASC;


-- TOP REVENUE MONTH
SELECT TOP(1)
	Year,
	MonthName AS 'Month',
	[Total Revenue]
FROM reporting.vw_TimePerformance
ORDER BY [Total Revenue] DESC;
/* November 2011 was the highest revenue-generating month, contributing £1.16M in total revenue. */



-- LOWEST REVENUE MONTH
SELECT TOP(1)
	Year,
	MonthName AS 'Month',
	[Total Revenue]
FROM reporting.vw_TimePerformance
ORDER BY [Total Revenue] ASC;
/* Revenue ranged from £446.08K in February 2011 to £1.16M in November 2011,
	highlighting significant seasonal variation in sales performance. */



-- TOP REVENUE QUARTER
SELECT TOP(1)
    Year,
    Quarter,
    SUM([Total Revenue]) AS [Total Revenue]
FROM reporting.vw_TimePerformance
GROUP BY Year, Quarter
ORDER BY [Total Revenue] DESC;
/* Revenue peaked in Q4 2011, reaching £2.71M, making it the best-performing quarter in the analysis period. */



-- MONTHLY ORDER VOLUMN
SELECT
	Year,
	Month,
	MonthName,
	[Total Orders],
	CAST([Total Revenue] / [Total Orders] AS DECIMAL(10, 2)) AS 'Avg Order Value'
FROM reporting.vw_TimePerformance;
/* Order volume peaked in November 2011 (2,657 orders),
	whereas December 2011 generated the highest average order value (£664.77). */



-- NEW CUSTOMERS BY MONTH
WITH FirstPurchase AS (	
	SELECT
		CustomerID,
		MIN(OrderDate) AS 'First Order Date'
	FROM datamodel.Fact_Sales
	WHERE IsReturn = 'No'
	GROUP BY CustomerID
)
SELECT
	dd.Year,
	dd.Month,
	dd.MonthName,
	COUNT(*) AS 'New Customers'
FROM FirstPurchase fp
INNER JOIN datamodel.Dim_Date dd
	ON fp.[First Order Date] = dd.Date
GROUP BY dd.Year, dd.Month, dd.MonthName
ORDER BY dd.Year, dd.Month;
/* Customer acquisition was highest in December 2010 with 885 new customers,
	while only 41 new customers were acquired in December 2011. */



-- ORDERS BY DAYS
SELECT
	dd.IsWeekend,
	dd.DayName,
	COUNT(DISTINCT InvoiceNo) AS 'Total Orders',
	SUM(fs.SalesAmount) AS 'Total Revenue'
FROM datamodel.Fact_Sales fs
INNER JOIN datamodel.Dim_Date dd
	ON fs.OrderDate = dd.Date
WHERE fs.IsReturn = 'No'
GROUP BY dd.IsWeekend, dd.DayName
ORDER BY [Total Orders] DESC;
/* Sales activity was concentrated on weekdays, with Thursday leading in both order volume (4,032 orders) and revenue (£1.97M),
	while Sunday recorded the lowest sales performance. */



-- MONTH OVER MONTH REVENUE GROWTH %
WITH MonthlyRevenue AS (
    SELECT
        dd.Year,
		dd.Month,
        dd.MonthName,
        SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END) AS Revenue
    FROM datamodel.Fact_Sales fs
    INNER JOIN datamodel.Dim_Date dd
		ON fs.OrderDate = dd.Date
    GROUP BY dd.Year, dd.Month, dd.MonthName
)
SELECT
    Year,
    MonthName,
    Revenue AS 'Total Revenue',
    LAG(Revenue) OVER(ORDER BY Year, Month) AS 'Previous Month Revenue',
    CAST(
        100.0 * (Revenue - LAG(Revenue) OVER(ORDER BY Year, Month))
        / NULLIF(LAG(Revenue) OVER(ORDER BY Year, Month), 0)
    AS DECIMAL(10,2)) AS 'MoM Growth %'
FROM MonthlyRevenue
ORDER BY Year, Month;
/* Revenue growth peaked in September 2011 with a 47.61% month-over-month increase,
	while December 2011 recorded the largest decline at -55.27%. */



-- QUARTER OVER QUARTER REVENUE GROWTH %
WITH QuarterlyRevenue AS (
    SELECT
        dd.Year,
        dd.Quarter,
        SUM(CASE WHEN fs.IsReturn = 'No' THEN fs.SalesAmount END) AS Revenue
    FROM datamodel.Fact_Sales fs
    INNER JOIN datamodel.Dim_Date dd
		ON fs.OrderDate = dd.Date
    GROUP BY dd.Year, dd.Quarter
)
SELECT
    Year,
    Quarter,
    Revenue AS 'Total Revenue',
    LAG(Revenue) OVER(ORDER BY Year, Quarter) AS 'Previous Quarter Revenue',
    CAST(
        100.0 * (Revenue - LAG(Revenue) OVER(ORDER BY Year, Quarter))
        / NULLIF(LAG(Revenue) OVER(ORDER BY Year, Quarter), 0)
    AS DECIMAL(10,2)) AS 'QoQ Growth %'
FROM QuarterlyRevenue
ORDER BY Year, Quarter;
/* Revenue increased every quarter throughout 2011, reaching £2.71M in Q4,
	with the strongest growth of 23.49% recorded between Q3 and Q4. */


-- ADVANCED TIME ANALYSIS
-- HOUR OFDAY ORDER PATTERN ANALYSIS
SELECT
    DATEPART(HOUR, InvoiceDate) AS 'Hour of Day',
    COUNT(DISTINCT InvoiceNo) AS 'Total Orders',
    SUM(CASE WHEN IsReturn = 'No' THEN SalesAmount END) AS 'Total Revenue'
FROM datamodel.Fact_Sales
WHERE IsReturn = 'No'
GROUP BY DATEPART(HOUR, InvoiceDate)
ORDER BY [Hour of Day];
/* Order activity peaked between 10 AM and 3 PM, with 12 PM recording the highest order volume (3,130 orders)
	and highest revenue (£1.37M), indicating that midday was the most active sales period. */



-- CUSTOMER REPEAT PURCHASE ANALYSIS (% OF FIRST TIME CUSTOMERS WHO WENT ON TO REPURCHASE AGAIN)
WITH FirstPurchase AS (
    SELECT
        CustomerID,
        MIN(OrderDate) AS FirstOrderDate
    FROM datamodel.Fact_Sales
    WHERE IsReturn = 'No'
    GROUP BY CustomerID
),
RepurchaseMonth AS (
    SELECT
        fp.CustomerID,
        dd.Year AS PurchaseYear,
        dd.Month AS PurchaseMonth,
        dd.MonthName AS PurchaseMonthName
    FROM FirstPurchase fp
    INNER JOIN datamodel.Dim_Date dd
		ON fp.FirstOrderDate = dd.Date
)
SELECT
    rm.PurchaseYear AS 'Purchase Year',
    rm.PurchaseMonthName AS 'Purchase Month',
    COUNT(DISTINCT rm.CustomerID) AS 'New Customers',
    COUNT(DISTINCT CASE WHEN dc.Frequency > 1 THEN rm.CustomerID END) AS 'Repeat Buyers (to date)',
    CAST(
        100.0 * COUNT(DISTINCT CASE WHEN dc.Frequency > 1 THEN rm.CustomerID END)
        / COUNT(DISTINCT rm.CustomerID)
    AS DECIMAL(10,2)) AS 'Repeat Rate %'
FROM RepurchaseMonth rm
INNER JOIN datamodel.Dim_Customers dc
	ON rm.CustomerID = dc.CustomerID
GROUP BY rm.PurchaseYear, rm.PurchaseMonth, rm.PurchaseMonthName
ORDER BY rm.PurchaseYear, rm.PurchaseMonth;




--===================================================================================
-- SECTION 7G: SOME IMPORTANT INSIGHTS
--===================================================================================


-- REVENUE AT RISK VS SAFE REVENUE VS POTENTIAL RECOVERY VALUE
SELECT
    SUM(CASE WHEN Segment IN ('At Risk', 'Lost Customers') THEN Monetary ELSE 0 END) AS 'Revenue At Risk',
    SUM(CASE WHEN Segment IN ('Champions', 'Loyal Customers') THEN Monetary ELSE 0 END) AS 'Safe Revenue',
    CAST(
        SUM(CASE WHEN Segment IN ('At Risk', 'Lost Customers') THEN Monetary ELSE 0 END) * 0.20
    AS DECIMAL(12,2)) AS 'Potential Recovery Value'
FROM datamodel.Dim_Customers;
/* £1.33M of revenue is at risk from At Risk and Lost Customers, while Champions and Loyal Customers contribute £6.99M in safe revenue;
	recovering just 20% of the at-risk revenue could generate an additional £265.73K. */



-- AT RISK CUSTOMERS
SELECT
    CustomerID,
	CustomersCountry,
	Segment,
	Monetary,
	Frequency,
	Recency
FROM datamodel.Dim_Customers
WHERE Segment IN ('At Risk', 'Lost Customers');





--===================================================================================
-- SECTION 14: STORED PROCEDURES
--===================================================================================


-- CUSTOMER DETAILS AND TRANSACTION HISTORY
CREATE PROCEDURE reporting.sp_GetCustomerDetails @CustomerID INT AS
BEGIN
    SELECT
        CustomerID,
		CustomersCountry AS 'Country',
		Segment,
        Recency AS 'Days Since Last Purchase',
		Frequency AS '# of Orders',
		Monetary AS 'Revenue',
        R_Score,
		F_Score,
		M_Score
    FROM datamodel.Dim_Customers
    WHERE CustomerID = @CustomerID;

    SELECT
        fs.InvoiceNo,
		fs.OrderDate,
		dp.Description AS 'Product Name',
        fs.Quantity,
		fs.UnitPrice,
		fs.SalesAmount,
		fs.IsReturn
    FROM datamodel.Fact_Sales fs
    JOIN datamodel.Dim_Products dp
		ON fs.StockCode = dp.StockCode
    WHERE fs.CustomerID = @CustomerID
    ORDER BY fs.OrderDate DESC;
END;


-- CUSTOMER DETAILS LOOKUP BY CUSTOMER ID
EXEC reporting.sp_GetCustomerDetails @CustomerID = 14646;



-- CUSTOMER DETAILS BY SEGMENT
CREATE PROCEDURE reporting.sp_GetCustomerSegmentDetails @Segment NVARCHAR(30) AS
BEGIN
    SELECT
        CustomerID,
		CustomersCountry AS 'Country',
		Recency AS 'Days Since Last Purchase',
		Frequency AS '# of Orders',
		Monetary AS 'Revenue',
		RFM_Score
    FROM datamodel.Dim_Customers
    WHERE Segment = @Segment
    ORDER BY Monetary DESC;
END;


-- CUSTOMER SEGMENT DETAILS LOOKUP
EXEC reporting.sp_GetCustomerSegmentDetails @Segment = 'Champions';




/*
===================================================================================
-- SECTION 8: KEY FINDINGS, ROOT CAUSES, BUSINESS INSIGHTS & RECOMMENDATIONS
===================================================================================

===================================================================================
-- SECTION 8A:  DATA CLEANING SUMMARY (BEFORE vs AFTER)
===================================================================================

    Raw table (original.OnlineRetail)                          541,909
    Step 01 - Removed blank / NULL CustomerID   (-135,080)     406,829
    Step 02 - Removed exact duplicate rows      (-5,225)       401,604
    Step 05 - Removed zero / negative UnitPrice (-44)          401,560   <-- final total records

    Total removed : 140,349 rows (25.90%)
    Total retained: 401,560 rows (74.10%)

Total records and unique values, before vs after cleaning

    Metric               Before Cleaning    After Cleaning    Change
    -------------------  -----------------  ----------------  ---------
    Total Records        541,909            401,560           -140,349
    Unique InvoiceNo     25,900             22,186            -3,714
    Unique StockCode     3,958              3,683             -275
    Unique CustomerID    4,372              4,371             -1
    Unique Countries     38                 37                -1


What this tells us
  -- The data cleaning process removed many transaction records, but only one customer was lost.
	 This means the customer base used for RFM analysis (4,371 customers) remains almost unchanged.
  -- Most of the removed records (135,080 rows) did not have a Customer ID. Since these transactions
     cannot be linked to a specific customer, they cannot be included in Recency, Frequency, and Monetary (RFM) analysis.
	 Therefore, the analysis focuses only on identified customers.
  -- Return and cancelled transactions (8,872 rows) were intentionally kept in the dataset and marked as `IsReturn = 'Yes'`.
     These transactions are excluded from RFM calculations but are analyzed separately to understand return behavior and
	 its impact on the business.


===================================================================================
-- SECTION 8B: KEY FINDINGS
===================================================================================

  REVENUE & OPERATIONS

  1. Gross revenue across the full period was £8.89M, reduced to a net revenue of £8.28M after £608.69K was lost to returns,
	 an overall revenue return rate of 6.85%.
  2. Weekdays drove the large majority of activity: 16,363 orders vs. 2,169 on weekends, and
	 3,481 returned orders on weekdays vs just 173 on weekends.
  3. Revenue was highly seasonal: monthly revenue ranged from £446.08K (Feb 2011) to £1.16M (Nov 2011),
	 with Q4 2011 the strongest quarter at £2.71M.
  4. Order activity peaked between 10 AM and 3 PM, with 12 PM the single busiest hour (3,130 orders, £1.37M revenue).
	 Thursday was the strongest day overall (4,032 orders, £1.97M revenue) & Sunday is the weakest day.


  CUSTOMER SEGMENTATION (RFM)

  5. Champions (1,009 customers, 23.1% of the base) generated £5.85M -- 65.84% of total revenue and
	 carried the highest average order value (£543.24) and the most orders placed (10,772, or 58.13% of all orders).
  6. Lost Customers (816 customers, 18.7% of the base) generated only £175.17K -- 1.97% of revenue,
	 the lowest of any segment, with the lowest AOV (£223.71) and lowest share of orders (4.23%).
  7. At Risk (933 customers), Loyal Customers (874), Potential Loyalists (739) -- together with Champions and Lost Customers,
     accounting for all 4,371 customers.
  8. Customers cluster at the extremes of the RFM scale: both the top score band (RFM 555) and the bottom band (RFM 111)
	 each contain 350+ customers, with comparatively few customers in the middle bands.
  9. 1,482 customers (33.9%) have not purchased in over 90 days; 1,628 customers purchased within the last 30 days.


  CUSTOMER VALUE & CONCENTRATION

 10. The top 20% of customers by revenue (875 customers) generated 74.85% of total revenue (£6.65M);
	 the remaining 80% contributed only 25.15%.
 11. The top 1% of customers (44 customers) alone generated £2.85M -- 32.06% of total revenue.
 12. Repeat customers generated £8.27M (93.09% of revenue); one-time customers contributed just 6.91%.
	 Even so, 34.16% of all customers made only one purchase and never returned.
 13. The single highest-value customer (ID 14646) generated £280.21K in total revenue;
	 the most active customer (ID 12748) placed 209 orders.
 14. 1,798 customers have a lifetime value between £0-£500; only 209 customers exceed £6K in lifetime value.


  PRODUCT PERFORMANCE

 15. "PAPER CRAFT, LITTLE BIRDIE" ranks number one (#1) by both revenue (£168.47K) and units sold (80,995)
	 (but see Root Cause #3 below before treating this as a genuine top-seller.)
 16. The top 20% of products (737 SKUs) generated 78.80% of total product revenue (£7.00M);
	 the remaining 80% contributed only 21.20%.
 17. Products priced £1-£5 generated the highest revenue (£6.04M) and highest quantity sold (2.92M units) of any price band.
 18. £6.55M in revenue comes from products with a return rate of 5% or lower;
     only 177 products fall into the high-return category (>15%).
 19. "JUMBO BAG PINK POLKADOT" and "JUMBO BAG RED RETROSPOT" were purchased together in 546 orders,
	 the strongest cross-sell pairing identified.


  GEOGRAPHYCAL PERFORMANCE

 20. The United Kingdom dominates every dimension measured: £7.29M in revenue (81.97% of the total), 3,920 customers,
	 16,646 orders, and 4.24M units sold.
 21. The top 20% of countries (8 of 38) generated 96.01% of total revenue (£8.53M);
	 the remaining 29 countries contributed only 3.99% combined.
 22. Saudi Arabia was the smallest market on every measure (£145.92 revenue, 80 units sold).
	 Nine countries had only a single customer each (Brazil, RSA, Iceland, Lithuania, European Community, Lebanon,
	 Saudi Arabia, Singapore, Czech Republic).
 23. The USA showed an anomalously high return rate of 57.93%, versus under 12% for most major markets
	 (See Root Cause #4.)
 24. Switzerland and Belgium recorded the highest repeat customer rates (75%) of any market with meaningful volume.


  REVENUE RISK

 25. £1.33M in revenue sits with At Risk and Lost Customers, against £6.99M classified as "safe" revenue (Champions + Loyal Customers).
	 Recovering just 20% of the at-risk revenue would recapture an estimated £265.73K.



===================================================================================
-- SECTION 8C: ROOT CAUSES
===================================================================================

  1. REVENUE IS HIGHLY CONCENTRATED AMONG A SMALL GROUP OF CUSTOMERS.

     The same 80/20-style pattern repeats at every level of this analysis -- 20% of customers drive ~75% of revenue,
	 20% of products drive ~79% of revenue, and just 8 countries drive 96% of revenue. This isn't three
     unrelated findings; it's one underlying pattern showing up from three angles. The root cause is a naturally
	 winner take most retail demand curve, reinforced by the business's near total dependence on a single home
     market (the UK).


  2. DECEMBER 2011 FIGURES ARE MISLEADING BECAUSE THE MONTH IS INCOMPLETE.

     The dataset ends on 9 December 2011 -- December is a partial month (roughly 9 days of data), not a full one.
	 This fully explains several findings that would otherwise look like alarming declines: the -55.27%
     month-over-month revenue drop in December, the sharp fall to only 41 new customers that month, and
	 December's elevated 33.96% return rate (a small number of returns against a small number of completed sales
     produces an exaggerated percentage). None of these are genuine seasonal or business health signals 
	 they are artifacts of the data collection window ending mid-month, and should be excluded or footnoted in any
     forward-looking trend analysis.


  3. THE #1 PRODUCT RANKING IS DRIVEN BY A SINGLE CUSTOMER'S BULK ORDER, NOT REGULAR CUSTOMER DEMAND.

     Customer 16446 placed one order for 80,995 units of "PAPER CRAFT, LITTLE BIRDIE" and later fully cancelled the entire order
	 on a separate invoice. Because the original sale and the cancellation are recorded as two distinct invoices,
	 the product still ranks #1 by both revenue and units sold under a same invoice only view of the data, despite this single
     customer relationship ending in a complete reversal.
	 A very similar pattern can be seen with Customer 12346, who placed a bulk order of ~74,215 units of
	 "MEDIUM CERAMIC TOP STORAGE JAR" and later cancelled the entire order. These unusual transactions likely contribute a 
	 significant portion of the £608.69K revenue lost due to returns (Key Finding #1). 
	 The high ranking of this SKU does not reflect actual customer demand. Instead, it is mainly driven by two large one time
	 transactions from two individual customers, both of which proved to be unstable and were ultimately cancelled.


  4. THE USA'S HIGH RETURN RATE IS DRIVEN BY LOW SALES VOLUME, NOT NECESSARILY A QUALITY ISSUE.

	 The USA's 57.93% return rate looks very high compared to the UK's low single digit return rate. 
	 However, the USA has a very small sales volume of only 2,458 units, while the UK sold 4.24M units. 
	 Because the USA has so few transactions, even a small number of returned orders can increase the return rate significantly. 
	 This is a small denominator effect and does not necessarily indicate a product or service problem in the USA market. 
	 However, since the market size is very small, the returned transactions should still be reviewed manually for further investigation.


  5. A LARGE NUMBER OF LOST CUSTOMERS INDICATES A RETENTION ISSUE.

	 18.7% of customers (816 customers) are classified as Lost, and another 21.3% (933 customers) are in the At Risk segment.
	 Together, these groups represent a large portion of the customer base that may no longer be actively engaged.
	 At the same time, repeat customers generate 93.09% of total revenue (Key Finding #12), showing that retained customers
	 create significant long-term value. However, 34.16% of customers make only one purchase and never return,
	 suggesting that the main challenge is not attracting customers but encouraging them to come back.
	 This indicates a need for stronger post purchase engagement, retention programs, and repeat purchase strategies.



===================================================================================
-- SECTION 8D: BUSINESS INSIGHTS
===================================================================================

  1. The business is heavily dependent on the UK market. The UK accounts for 81.97% of total revenue and
	 also contributes the majority of customers, orders, and sales volume. Although the company operates in international markets,
	 their contribution remains relatively small.

  2. A small group of customers drives most of the business revenue. Champions and Loyal Customers together contribute £6.99M in revenue,
	 while Champions alone (1,009 customers) generate nearly two-thirds of total revenue.
	 This shows that the company's performance depends heavily on keeping these high-value customers satisfied and engaged.
	 Losing even a portion of this group could have a noticeable impact on overall revenue.

  3. Getting customers to make a second purchase is the biggest opportunity for growth. More than one-third of customers
	 never return after their first order, while repeat customers generate the vast majority of revenue.
	 This suggests that the business would benefit more from increasing repeat purchases among existing customers
	 than from focusing solely on acquiring new customers. Converting more first-time buyers into repeat customers could have
	 a significant impact on long term revenue growth.

  4. Top product and country rankings should be interpreted carefully. Some rankings in this dataset are affected by large one time orders
	 that were later cancelled (Root Cause #3) and by small markets with very low transaction volumes (Root Cause #4).
	 As a result, raw rankings may not always reflect true customer demand or market performance.
	 Before using these rankings for inventory planning, marketing campaigns, or expansion decisions,
	 they should be validated against transaction volume and return/cancellation patterns.

  5. Returns are concentrated in specific products and transactions, making them easier to address.
	 The business already generates £6.55M in revenue from products with low return rates (≤5%),
	 showing that most of the product catalog performs well. The £608.69K in revenue lost due to returns is concentrated within
	 a smaller group of products and a limited number of unusual transactions. This suggests that a targeted approach (Root Cause #3)
	 would be more effective than making broad changes across the entire product catalog.



===================================================================================
-- SECTION 8E: BUSINESS RECOMMENDATIONS & ACTIONS
===================================================================================

  1. PROTECT THE CHAMPIONS SEGMENT WITH A FORMAL RETENTION PROGRAM.

	 Champions generate 65.84% of total revenue, making them the most valuable customer segment. To retain these customers,
	 the business should consider a dedicated loyalty program that offers benefits such as early access to new products,
	 priority support, or personalized rewards. Since even a small decline in this segment could significantly affect revenue,
	 Champion customers should be monitored regularly using the vw_SegmentationKPIs and vw_RevenueBySegment views,
	 rather than being treated the same as all other customers in marketing campaigns.


  2. BUILD A SECOND-PURCHASE CAMPAIGN TARGETING FIRST-TIME BUYERS.

	 Since 34.16% of customers do not make a second purchase, the business should focus on encouraging repeat orders.
	 A follow up campaign, such as an email, discount offer, or product recommendation, can be sent shortly after a
	 customer's first purchase to increase the likelihood of a second order. Recent first time customers can be identified using
	 the vw_NewCustomersByMonth logic already developed, allowing the business to target customers before they become one time buyers.


  3. RE-ENGAGE AT-RISK CUSTOMERS WITH TARGETED CAMPAIGNS.
   
	 At Risk and Lost Customers represent £1.33M in revenue, making them an important retention opportunity.
	 Using the reporting.sp_GetCustomerSegmentDetails procedure with @Segment = 'At Risk', the business can create a prioritized list
	 of high value customers for targeted outreach. Rather than treating all customers in this segment the same, efforts should focus
	 first on those with the highest revenue potential. A realistic initial objective would be to recover the estimated 
	 £265.73K in revenue from this group.


  4. AUDIT AND REMOVE CANCELLED BULK ORDERS FROM PRODUCT PERFORMANCE ANALYSIS.
     
	 Before using Top 10 product rankings for inventory or marketing decisions, unusual bulk transactions
	 should be reviewed and filtered where appropriate. Large one-time orders with exceptionally high quantities,
	 especially those that were later fully cancelled, can distort product rankings. Applying checks for unusually large orders or
	 matching cancellations (such as the cases involving Customer 16446 and Customer 12346) will help ensure that 
	 product performance reports reflect actual customer demand rather than isolated transactions.


  5. INVESTIGATE THE USA MARKET MANUALLY BEFORE ACTING ON ITS RETURN RATE.

     The USA's 57.93% return rate should be investigated before drawing conclusions. Since the market has a very low sales
	 volume of only 2,458 units sold, a small number of returned orders can have a large impact on the overall return rate.
	 Instead of implementing broad market level changes, the business should first review the underlying transactions to
	 determine whether the issue is driven by a few specific orders.


  7. USE PRODUCT CROSS-SELL OPPORTUNITIES TO INCREASE SALES.

     "JUMBO BAG PINK POLKADOT" and "JUMBO BAG RED RETROSPOT" are frequently purchased together, appearing in 546 orders.
	 The business can use this existing buying pattern to create bundle offers or "Customers Also Bought" recommendations
	 during checkout. This is a low cost opportunity to increase average order value by encouraging customers to purchase 
	 products that are already commonly bought together.


  8. EXPAND INTERNATIONAL SALES BY FOCUSING ON PROVEN MARKETS.

     Switzerland and Belgium show strong customer loyalty, with repeat-customer rates of 75%, despite having relatively
	 low sales volume. Because these markets have already demonstrated positive customer behavior,
	 they may be better candidates for future international growth initiatives than completely new or untested markets.
	 Expanding in countries where customer loyalty is already proven can help reduce the risks associated with international expansion.


*/














































































