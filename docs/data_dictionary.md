# Data dictionary

All data lives in Excel tables (ListObjects). Table names are what the VBA, Power Query and formulas refer to.

## Raw export files (`data/raw/sales_YYYY_MM.csv`)

One file per month, one line per product on an order. Same 16 columns as `tblRaw` below (minus the first two). Columns are matched by header name, so the order doesn't matter.

## tblRaw (sheet `Raw_Data`)

Rows exactly as imported. Every column is stored as text so nothing gets auto-converted by Excel.

| Column | Description |
|---|---|
| Raw Row ID | Running number assigned on import. Used everywhere to trace a row back to its source. |
| Source File | CSV file name, or `Manual Entry` for rows added through the Data Entry form |
| Order ID | e.g. `SO-101169`. Lines from one order share the ID. |
| Order Date / Ship Date | `YYYY-MM-DD` or `MM/DD/YYYY` (some store terminals) |
| Customer ID / Customer Name | `CUS-10175` / person or business name |
| Segment | Consumer, Corporate, Small Business |
| City / Province | Province in any spelling (code, full name, abbreviation) |
| Product ID / Product Name | `PRD-3001` / name as printed by the till |
| Quantity / Unit Price / Discount | Text as exported (`$1,234.50`, `10%`, `0.1`, blank...) |
| Sales Channel | In-Store / Online, in various spellings |
| Sales Rep / Payment Method | |

## tblStaging (sheet `Staging`)

Output of the Power Query query `SalesStaging`: typed, cleaned, de-duplicated and enriched. Values that couldn't be fixed are left blank for the validation step.

| Column | Type | Notes |
|---|---|---|
| RawRowID, SourceFile | number, text | carried over from tblRaw |
| OrderID, CustomerID, ProductID | text | upper-cased, trimmed |
| OrderDate, ShipDate | date | blank if the text wasn't a real date |
| CustomerName, Segment, City | text | spacing cleaned, each word capitalised |
| ProvinceCode, Province, Region | text | via `tblProvinceAlias` and `tblProvinces`; blank if unrecognised |
| ProductName, Category, SubCategory, ListPrice, UnitCost | | from `tblProducts`; blank if the product isn't in the master |
| Quantity, UnitPrice, Discount | number | Discount as a fraction (0.15 = 15%) |
| Channel | text | `In-Store` / `Online`, blank if unknown |
| SalesRep, PaymentMethod | text | |

## tblSales (sheet `Sales_Data`)

Rows that passed every Error rule, plus calculated columns. Source for all pivots, charts and KPIs.

| Column | Calculation |
|---|---|
| LineID | = RawRowID |
| Year, Quarter, MonthNum, MonthName | from OrderDate (`Q1`..`Q4`, `Jan`..`Dec`) |
| ShipDays | ShipDate - OrderDate |
| GrossSales | Quantity x UnitPrice |
| DiscountAmt | GrossSales x Discount |
| Revenue | GrossSales - DiscountAmt (net revenue) |
| COGS | Quantity x UnitCost |
| Profit | Revenue - COGS |
| MarginPct | Profit / Revenue |
| OrderFlag | 1 on the first line of each order, 0 on the rest. Summing it counts orders without the Data Model. |
| QualityFlag | `OK`, or `Warning` if the row failed a warning rule |

All money values are rounded to 2 decimals (half away from zero).

## tblExceptions (sheet `Exceptions`)

One row per failed rule: RawRowID, SourceFile, OrderID, ProductID, RuleID, Severity, Field, Issue (rule description), Value (the bad value), LoggedOn.

## Analysis tables

| Table | Grain | Columns |
|---|---|---|
| tblRFM (`Customer_RFM`) | customer | CustomerID, CustomerName, Segment, Region, FirstOrder, LastOrder, Orders, Revenue, Profit, RecencyDays, R, F, M, RFMScore, RFMSegment |
| tblABC (`Product_ABC`) | product | Rank, ProductID, ProductName, Category, Units, Revenue, Profit, MarginPct, RevenueShare, CumulativeShare, ABCClass |
| tblCommission (`Commission`) | rep x quarter | Year, Quarter, SalesRep, Region, Revenue, Quota, Attainment, Commission, EffectiveRate |

## Configuration tables

| Table | Sheet | Used by |
|---|---|---|
| tblRules | Validation_Rules | validation engine |
| tblProducts | Lookups | Power Query merge, Lookup rule, data entry dropdown, list price |
| tblProvinces, tblProvinceAlias | Lookups | Power Query province mapping, data entry dropdown |
| tblReps, tblCommissionTiers | Lookups | commission calculation, dropdowns |
| tblSegments, tblChannels, tblPayments | Lookups | data entry dropdowns, upload template |

## Logs

| Table | Columns |
|---|---|
| tblImportLog | FileName, Rows, ImportedOn, Status |
| tblRunLog | RunID, Started, Finished, Seconds, RowsImported, RowsValidated, RowsLoaded, RowsRejected, Warnings, Accuracy, CleanMethod, Status |
