# ☕ Coffee Sales Business Analysis
### Excel · VLOOKUP · XLOOKUP · Pivot Tables · MySQL · SQL Window Functions

[![Excel](https://img.shields.io/badge/Excel-Data%20Cleaning%20%26%20Dashboard-217346?logo=microsoftexcel&logoColor=white)](https://microsoft.com/excel)
[![MySQL](https://img.shields.io/badge/MySQL-KPI%20Queries-4479A1?logo=mysql&logoColor=white)](https://mysql.com)
[![Dataset](https://img.shields.io/badge/Dataset-Kaggle%20Coffee%20Orders-20BEFF?logo=kaggle)](https://www.kaggle.com/datasets/mohammadkaiftahir/coffee-orders-data)
[![GitHub](https://img.shields.io/badge/GitHub-surajvaidya20-black?logo=github)](https://github.com/surajvaidya20)

---

## 📌 Project Overview

An end-to-end **Excel + SQL business analysis** of a coffee sales company — covering raw data cleaning, formula-based data enrichment, pivot table analysis, interactive dashboard design, and advanced SQL KPI queries.

This project demonstrates the core skills tested by companies like **TCS, Wipro, Infosys, and Accenture** in data analyst interviews: Excel formulas, Pivot Tables, dashboard design, and SQL aggregations.

> **Dataset:** Coffee Orders Data — Orders, Customers & Products across 4 years and 3 countries  
> **Source:** [Kaggle — mohammadkaiftahir/coffee-orders-data](https://www.kaggle.com/datasets/mohammadkaiftahir/coffee-orders-data)

---

## 🎯 Business Questions Answered

| # | Business Question | Tool Used |
|---|-------------------|-----------|
| 1 | What is the total revenue, order count, and average order value? | Excel KPI Cards + SQL |
| 2 | Which coffee type generates the most revenue? | Pivot Table + SQL GROUP BY |
| 3 | Which country has the highest sales? | Pivot Chart + SQL |
| 4 | Who are our top 5 customers by spending? | Pivot Table + SQL TOP N |
| 5 | How has revenue grown year-over-year? | Line Chart + SQL YoY CTE |
| 6 | Do loyalty card holders spend more? | Pivot Table + SQL |
| 7 | Which month is the peak sales month per coffee type? | SQL Window Function |
| 8 | What is the cumulative revenue over time? | SQL Running Total |

---

## 🛠️ Excel Skills Demonstrated

| Skill | Where Used |
|-------|-----------|
| **VLOOKUP** | Pull product size and unit price from Products table |
| **XLOOKUP** | Retrieve customer name, email, country, loyalty card |
| **IF (Nested)** | Expand coffee type codes → full names (e.g. Rob → Robusta) |
| **Pivot Tables** | Sales by time, country, roast type, product, customer |
| **Pivot Charts** | Line, bar, donut, column charts from pivot tables |
| **Slicers** | Dynamic filters for roast type, size, loyalty card |
| **Timeline** | Date range filter connected to all pivot charts |
| **Conditional Formatting** | Highlight top sales, missing data |
| **Power Query** | Data type conversion and duplicate removal |

---

## 🛢️ SQL Skills Demonstrated

| Technique | Query |
|-----------|-------|
| `GROUP BY` + `ORDER BY` | Revenue by coffee type, country, roast |
| `ROUND`, `AVG`, `SUM` | KPI calculations |
| `DATE_FORMAT` | Monthly trend grouping |
| `CTE` (WITH clause) | YoY growth, peak month analysis |
| `LAG()` Window Function | Month-over-month and year-over-year growth |
| `RANK()` Window Function | Best-selling product per country |
| `SUM() OVER()` | Cumulative running total revenue |
| `CASE WHEN` | Customer segmentation (High/Mid/Low Value) |
| `NULLIF` | Safe division to avoid divide-by-zero |

---

## 📊 Excel Dashboard Preview

![Coffee Sales Dashboard](Screenshots/dashboard.png)

**Dashboard includes:**
- 4 KPI metric cards — Total Revenue, Orders, Customers, Avg Order Value
- Line chart — Total sales over time by coffee type
- Bar chart — Revenue by country
- Bar chart — Top 5 customers
- Donut chart — Revenue by roast type
- Slicers — Roast Type, Size, Loyalty Card
- Timeline — Dynamic date range filter

---

## 📂 Repository Structure

```
Coffee-Sales-Excel-SQL-Analysis/
│
├── Dataset/
│   └── coffee_orders_data.xlsx       ← Raw dataset from Kaggle (3 tables)
│
├── Excel/
│   └── Coffee_Sales_Dashboard.xlsx   ← Cleaned data + Pivot Tables + Dashboard
│
├── SQL/
│   └── coffee_sales_sql_queries.sql  ← 15 SQL queries (KPIs + Window Functions)
│
├── Screenshots/
│   └── dashboard.png                 ← Dashboard screenshot for portfolio
│
├── Excel_Building_Guide.md           ← Step-by-step Excel build instructions
└── README.md
```

---

## 🚀 How to Reproduce

### Excel Part
1. Download dataset from [Kaggle](https://www.kaggle.com/datasets/mohammadkaiftahir/coffee-orders-data)
2. Open `coffee_orders_data.xlsx` in Excel
3. Follow `Excel_Building_Guide.md` step by step
4. Final output: Interactive dashboard with slicers and timeline

### SQL Part
1. Export your cleaned Excel Orders sheet as CSV
2. Import into MySQL: Right-click database → Table Data Import Wizard
3. Open `coffee_sales_sql_queries.sql` in MySQL Workbench
4. Run queries section by section

---

## 📈 Key Business Insights

- **Liberica coffee** generated the highest revenue despite lower order volume — premium pricing effect
- **United States** accounts for 79% of total revenue across all years
- **Dark roast** is the most popular by order count, but **Light roast** has higher average order value
- **Loyalty card holders** spend on average 18% more per order than non-members
- Revenue peaked in **Q3 (July–September)** consistently across all 4 years — seasonal pattern
- Top 5 customers alone contribute over **21% of total revenue** — high customer concentration risk

---

## 💼 Business Recommendations

1. **Expand loyalty programme** — loyalty card holders spend 18% more; incentivise sign-ups
2. **Focus marketing on US market** — 79% revenue share; deepen penetration before expanding
3. **Promote Liberica in international markets** — high revenue per order but low volume outside US
4. **Stock up for Q3** — consistent seasonal peak in July–September; plan inventory 8 weeks ahead
5. **Reduce customer concentration risk** — top 5 customers = 21% revenue; diversify customer base

---

## 👤 Author

**Suraj Vaidya**  
B.E. Student — Data Analytics & AI/ML  
📧 surajvaidya11@gmail.com  
🔗 [GitHub](https://github.com/surajvaidya20) | [LinkedIn](https://linkedin.com/in/surajvaidya)  
📍 Chandrapur, Maharashtra, India

---

## 📄 License

This project is open-source under the [MIT License](LICENSE).

---

> ⭐ Star this repo if it helped you build your own portfolio project!
