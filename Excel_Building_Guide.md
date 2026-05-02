# ══════════════════════════════════════════════════════════════
#  COFFEE SALES BUSINESS ANALYSIS — COMPLETE EXCEL GUIDE
#  Author : Suraj Vaidya | GitHub: surajvaidya20
#  Dataset : Coffee Orders Data (Kaggle)
#  Tool    : Microsoft Excel (2016 / 2019 / 365)
# ══════════════════════════════════════════════════════════════

DATASET LINK : https://www.kaggle.com/datasets/mohammadkaiftahir/coffee-orders-data
RECORDS      : ~1,000 orders across 3 tables — Orders, Customers, Products
EXCEL SKILLS : VLOOKUP, XLOOKUP, IF, Pivot Tables, Pivot Charts, Slicers,
               Timeline, Conditional Formatting, Power Query, Dashboard

══════════════════════════════════════════════════════════════
STEP 1 — DOWNLOAD & UNDERSTAND THE DATA
══════════════════════════════════════════════════════════════

1. Go to: https://www.kaggle.com/datasets/mohammadkaiftahir/coffee-orders-data
2. Download the file — it comes as coffee_orders_data.xlsx
3. Open in Excel. You will see 3 sheets:

   ORDERS TABLE (main table):
   - order_id, order_date, customer_id, product_id, quantity

   CUSTOMERS TABLE:
   - customer_id, customer_name, email, phone, address, city,
     country, postcode, loyalty_card

   PRODUCTS TABLE:
   - product_id, coffee_type, roast_type, size, unit_price,
     price_per_100g, profit

══════════════════════════════════════════════════════════════
STEP 2 — DATA CLEANING (Working Sheet)
══════════════════════════════════════════════════════════════

A) CONVERT ALL 3 SHEETS INTO EXCEL TABLES
   - Click anywhere inside Orders data
   - Press Ctrl + T → check "My table has headers" → OK
   - Name the table: "Orders" (Table Design tab → Table Name)
   - Repeat for Customers table → name "Customers"
   - Repeat for Products table  → name "Products"

B) CHECK FOR ISSUES IN ORDERS SHEET
   - Select all → Data tab → Remove Duplicates
   - Check each column for blanks: Home → Find & Select → Go To Special → Blanks
   - Check order_date column: ensure format is DATE (not text)
     → Select column → Ctrl+1 → Number → Date → dd-mmm-yyyy

C) STANDARDIZE DATE FORMAT
   - Click on order_date column header → Format Cells → Date
   - Choose: 14-Mar-2012 style (e.g. 05-Jan-2019)

══════════════════════════════════════════════════════════════
STEP 3 — DATA ENRICHMENT USING VLOOKUP / XLOOKUP / IF
══════════════════════════════════════════════════════════════
(Add these columns to the Orders sheet — after quantity column)

─────────────────────────────────────────────
COLUMN 1: Customer Name  (using XLOOKUP)
─────────────────────────────────────────────
=XLOOKUP(C2, Customers[Customer ID], Customers[Customer Name],,0)

Explanation: Look up the customer_id (C2) in Customers table,
return their name. The 0 at end means exact match.

─────────────────────────────────────────────
COLUMN 2: Email  (using XLOOKUP)
─────────────────────────────────────────────
=IF(XLOOKUP(C2, Customers[Customer ID], Customers[Email],,0)=0,
   "Not Provided",
   XLOOKUP(C2, Customers[Customer ID], Customers[Email],,0))

Explanation: Gets email. If blank, shows "Not Provided" instead.

─────────────────────────────────────────────
COLUMN 3: Country  (using XLOOKUP)
─────────────────────────────────────────────
=XLOOKUP(C2, Customers[Customer ID], Customers[Country],,0)

─────────────────────────────────────────────
COLUMN 4: Coffee Type  (using XLOOKUP from Products)
─────────────────────────────────────────────
=XLOOKUP(D2, Products[Product ID], Products[Coffee Type],,0)

─────────────────────────────────────────────
COLUMN 5: Roast Type  (using XLOOKUP)
─────────────────────────────────────────────
=XLOOKUP(D2, Products[Product ID], Products[Roast Type],,0)

─────────────────────────────────────────────
COLUMN 6: Size  (using VLOOKUP — practice both!)
─────────────────────────────────────────────
=VLOOKUP(D2, Products[#All], 4, FALSE)

Explanation: D2 = Product ID, Products[#All] = full Products table,
4 = 4th column (Size), FALSE = exact match.

─────────────────────────────────────────────
COLUMN 7: Unit Price  (using VLOOKUP)
─────────────────────────────────────────────
=VLOOKUP(D2, Products[#All], 5, FALSE)

─────────────────────────────────────────────
COLUMN 8: Sales  (Calculated column)
─────────────────────────────────────────────
= [@Quantity] * [@[Unit Price]]

─────────────────────────────────────────────
COLUMN 9: Coffee Type Full Name  (using IF)
─────────────────────────────────────────────
=IF([@[Coffee Type]]="Rob","Robusta",
  IF([@[Coffee Type]]="Exc","Excelsa",
    IF([@[Coffee Type]]="Ara","Arabica",
      IF([@[Coffee Type]]="Lib","Liberica",""))))

─────────────────────────────────────────────
COLUMN 10: Roast Type Full Name  (using IF)
─────────────────────────────────────────────
=IF([@[Roast Type]]="M","Medium",
  IF([@[Roast Type]]="L","Light",
    IF([@[Roast Type]]="D","Dark","")))

─────────────────────────────────────────────
COLUMN 11: Loyalty Card  (from Customers)
─────────────────────────────────────────────
=XLOOKUP(C2, Customers[Customer ID], Customers[Loyalty Card],,0)

══════════════════════════════════════════════════════════════
STEP 4 — CREATE PIVOT TABLES (Insert → PivotTable)
══════════════════════════════════════════════════════════════
Create each on a new sheet called "PivotTables"

─────────────────────────────────────────────
PIVOT TABLE 1: Total Sales Over Time
─────────────────────────────────────────────
- Rows    : order_date (group by Month → right-click → Group → Months & Years)
- Columns : Coffee Type Full Name
- Values  : Sum of Sales (format as Currency $)

→ Creates a line chart showing sales trend by coffee type over time

─────────────────────────────────────────────
PIVOT TABLE 2: Sales by Country
─────────────────────────────────────────────
- Rows    : Country
- Values  : Sum of Sales

→ Creates a bar chart — which country buys the most?

─────────────────────────────────────────────
PIVOT TABLE 3: Top 5 Customers by Sales
─────────────────────────────────────────────
- Rows    : Customer Name
- Values  : Sum of Sales
- Filter  : Top 5 (Value Filters → Top 10 → change to 5)

→ Creates a horizontal bar chart of top customers

─────────────────────────────────────────────
PIVOT TABLE 4: Sales by Roast Type
─────────────────────────────────────────────
- Rows    : Roast Type Full Name
- Values  : Sum of Sales, Count of order_id

→ Creates a donut/pie chart

─────────────────────────────────────────────
PIVOT TABLE 5: Sales by Size
─────────────────────────────────────────────
- Rows    : Size
- Values  : Sum of Sales

→ Creates a column chart

══════════════════════════════════════════════════════════════
STEP 5 — BUILD THE DASHBOARD SHEET
══════════════════════════════════════════════════════════════
Create a new sheet → rename to "Dashboard"

LAYOUT:
┌─────────────────────────────────────────────────────────┐
│   ☕ Coffee Sales Business Dashboard   [Logo optional]   │
├──────────────┬──────────────┬───────────────────────────┤
│  Total Sales │ Total Orders │ Total Customers  Avg Order │
│  $xxxxxxx    │    1,000     │      xxx         $xx.xx   │
├──────────────────────────┬──────────────────────────────┤
│  Sales Over Time (Line)  │   Sales by Country (Bar)     │
│                          │                              │
├──────────────────────────┼──────────────────────────────┤
│  Top 5 Customers (Bar)   │  Sales by Roast Type (Donut) │
│                          │                              │
├──────────────────────────┴──────────────────────────────┤
│  [ Timeline ] [ Slicer: Roast ] [ Slicer: Size ]        │
│              [ Slicer: Loyalty Card ]                   │
└─────────────────────────────────────────────────────────┘

HOW TO BUILD:
1. Copy each Pivot Chart (right-click → Copy) → paste on Dashboard sheet
2. Resize charts to fit the layout above
3. Remove gridlines: View → uncheck Gridlines
4. Add background color: Select all (Ctrl+A) → Fill Color → light gray or white
5. Add title text box: Insert → Text Box → type "☕ Coffee Sales Business Dashboard"
   → Font size 18, Bold, Dark color

KPI CARDS (manual — use text boxes + formulas):
- Total Sales:    =SUM(Orders[Sales])         → format as currency
- Total Orders:   =COUNTA(Orders[Order ID])
- Avg Order Size: =AVERAGE(Orders[Sales])
- Total Customers: =COUNTA(UNIQUE(Orders[Customer Name]))

Place each in a colored rectangle (Insert → Shapes → Rectangle)
Use colors: Dark Blue header, White text

SLICERS:
1. Click on any Pivot Chart → PivotChart Analyze → Insert Slicer
2. Add: Roast Type Name, Size, Loyalty Card
3. Right-click each slicer → Report Connections → select ALL pivot tables
   (so one click filters every chart)

TIMELINE:
1. Click any Pivot Chart → PivotChart Analyze → Insert Timeline
2. Select order_date → OK
3. Report Connections → connect to all pivot tables

══════════════════════════════════════════════════════════════
STEP 6 — CONDITIONAL FORMATTING
══════════════════════════════════════════════════════════════
On the Orders working sheet:

A) Highlight top 10% sales rows:
   - Select Sales column → Home → Conditional Formatting
   → Top/Bottom Rules → Top 10% → Green fill

B) Highlight blank emails:
   - Select Email column → Conditional Formatting
   → Highlight Cells Rules → Equal To → "Not Provided" → Red fill

C) Data Bars on Sales column:
   - Select Sales column → Conditional Formatting → Data Bars → Blue

══════════════════════════════════════════════════════════════
STEP 7 — FINAL CHECKS & SCREENSHOT
══════════════════════════════════════════════════════════════

□ Dashboard has title and date
□ All charts have clear titles (double-click chart title to edit)
□ KPI cards show correct numbers
□ Slicers are connected to all charts
□ Timeline works (click on a month — all charts update)
□ No gridlines on Dashboard sheet
□ File saved as .xlsx

TO SCREENSHOT:
- Press Windows + Shift + S → select dashboard area
- Save as: dashboard_screenshot.png
- Upload to GitHub

══════════════════════════════════════════════════════════════
TIME ESTIMATE
══════════════════════════════════════════════════════════════
Step 1-2  (Download + Clean):          30 minutes
Step 3    (VLOOKUP/XLOOKUP/IF):        1.5 hours
Step 4    (5 Pivot Tables):            1 hour
Step 5    (Dashboard):                 2 hours
Step 6    (Conditional Formatting):    30 minutes
Step 7    (Final + Screenshot):        30 minutes

TOTAL: ~6 hours across 1-2 days
══════════════════════════════════════════════════════════════
