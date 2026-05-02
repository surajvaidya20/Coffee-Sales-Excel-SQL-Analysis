-- ══════════════════════════════════════════════════════════════
-- COFFEE SALES BUSINESS ANALYSIS — SQL KPI QUERIES
-- Author  : Suraj Vaidya | GitHub: surajvaidya20
-- Database: MySQL  |  Table: coffee_orders
-- Dataset : https://www.kaggle.com/datasets/mohammadkaiftahir/coffee-orders-data
-- ══════════════════════════════════════════════════════════════
-- HOW TO USE:
-- 1. Export your cleaned Excel file as CSV
-- 2. Import into MySQL using: Table Data Import Wizard
-- 3. Run queries below in MySQL Workbench
-- ══════════════════════════════════════════════════════════════


-- ─────────────────────────────────────────────────────────────
-- SECTION 1: OVERALL BUSINESS KPIs
-- ─────────────────────────────────────────────────────────────

-- Q1. Core KPI Summary — Revenue, Orders, Customers, Avg Order
SELECT
    COUNT(DISTINCT order_id)                        AS total_orders,
    COUNT(DISTINCT customer_name)                   AS unique_customers,
    ROUND(SUM(sales), 2)                            AS total_revenue,
    ROUND(AVG(sales), 2)                            AS avg_order_value,
    ROUND(SUM(sales) / COUNT(DISTINCT customer_name), 2) AS revenue_per_customer
FROM coffee_orders;


-- Q2. Total Revenue by Year
SELECT
    YEAR(order_date)          AS order_year,
    ROUND(SUM(sales), 2)      AS total_revenue,
    COUNT(DISTINCT order_id)  AS total_orders
FROM coffee_orders
GROUP BY YEAR(order_date)
ORDER BY order_year;


-- Q3. Monthly Revenue Trend (all years)
SELECT
    DATE_FORMAT(order_date, '%Y-%m') AS month,
    ROUND(SUM(sales), 2)             AS monthly_revenue,
    COUNT(order_id)                  AS total_orders
FROM coffee_orders
GROUP BY DATE_FORMAT(order_date, '%Y-%m')
ORDER BY month;


-- ─────────────────────────────────────────────────────────────
-- SECTION 2: PRODUCT ANALYSIS
-- ─────────────────────────────────────────────────────────────

-- Q4. Revenue by Coffee Type
SELECT
    coffee_type_name,
    ROUND(SUM(sales), 2)      AS total_revenue,
    COUNT(order_id)           AS total_orders,
    ROUND(AVG(sales), 2)      AS avg_order_value,
    ROUND(SUM(sales) * 100.0 /
        (SELECT SUM(sales) FROM coffee_orders), 2) AS revenue_share_pct
FROM coffee_orders
GROUP BY coffee_type_name
ORDER BY total_revenue DESC;


-- Q5. Revenue by Roast Type
SELECT
    roast_type_name,
    ROUND(SUM(sales), 2)    AS total_revenue,
    COUNT(order_id)         AS total_orders,
    ROUND(AVG(unit_price), 2) AS avg_unit_price
FROM coffee_orders
GROUP BY roast_type_name
ORDER BY total_revenue DESC;


-- Q6. Revenue by Product Size
SELECT
    size,
    ROUND(SUM(sales), 2)    AS total_revenue,
    COUNT(order_id)         AS total_orders,
    SUM(quantity)           AS total_units_sold
FROM coffee_orders
GROUP BY size
ORDER BY total_revenue DESC;


-- Q7. Best-Selling Coffee Type per Country (Window Function)
WITH country_product_sales AS (
    SELECT
        country,
        coffee_type_name,
        ROUND(SUM(sales), 2) AS revenue
    FROM coffee_orders
    GROUP BY country, coffee_type_name
),
ranked AS (
    SELECT *,
        RANK() OVER (PARTITION BY country ORDER BY revenue DESC) AS rn
    FROM country_product_sales
)
SELECT country, coffee_type_name, revenue
FROM ranked
WHERE rn = 1
ORDER BY revenue DESC;


-- ─────────────────────────────────────────────────────────────
-- SECTION 3: CUSTOMER ANALYSIS
-- ─────────────────────────────────────────────────────────────

-- Q8. Top 10 Customers by Total Revenue
SELECT
    customer_name,
    country,
    ROUND(SUM(sales), 2)   AS total_spent,
    COUNT(order_id)        AS total_orders,
    loyalty_card
FROM coffee_orders
GROUP BY customer_name, country, loyalty_card
ORDER BY total_spent DESC
LIMIT 10;


-- Q9. Loyalty Card vs Non-Loyalty Card Spending
SELECT
    loyalty_card,
    COUNT(DISTINCT customer_name)  AS customer_count,
    ROUND(SUM(sales), 2)           AS total_revenue,
    ROUND(AVG(sales), 2)           AS avg_order_value,
    ROUND(SUM(sales) * 100.0 /
        (SELECT SUM(sales) FROM coffee_orders), 2) AS revenue_share_pct
FROM coffee_orders
GROUP BY loyalty_card;


-- Q10. Revenue by Country
SELECT
    country,
    COUNT(DISTINCT customer_name)  AS total_customers,
    ROUND(SUM(sales), 2)           AS total_revenue,
    ROUND(AVG(sales), 2)           AS avg_order_value
FROM coffee_orders
GROUP BY country
ORDER BY total_revenue DESC;


-- ─────────────────────────────────────────────────────────────
-- SECTION 4: ADVANCED ANALYSIS (Window Functions & CTEs)
-- ─────────────────────────────────────────────────────────────

-- Q11. Year-over-Year Revenue Growth
WITH yearly AS (
    SELECT
        YEAR(order_date)      AS yr,
        ROUND(SUM(sales), 2)  AS revenue
    FROM coffee_orders
    GROUP BY YEAR(order_date)
)
SELECT
    yr,
    revenue,
    LAG(revenue) OVER (ORDER BY yr)  AS prev_year_revenue,
    ROUND(
        (revenue - LAG(revenue) OVER (ORDER BY yr)) /
        NULLIF(LAG(revenue) OVER (ORDER BY yr), 0) * 100
    , 2) AS yoy_growth_pct
FROM yearly;


-- Q12. Running Total Revenue (Cumulative Sales over Time)
SELECT
    DATE_FORMAT(order_date, '%Y-%m') AS month,
    ROUND(SUM(sales), 2)             AS monthly_revenue,
    ROUND(SUM(SUM(sales)) OVER (ORDER BY DATE_FORMAT(order_date, '%Y-%m')), 2)
        AS cumulative_revenue
FROM coffee_orders
GROUP BY DATE_FORMAT(order_date, '%Y-%m')
ORDER BY month;


-- Q13. Customer Purchase Frequency Segments (RFM-lite)
WITH customer_stats AS (
    SELECT
        customer_name,
        COUNT(order_id)        AS order_count,
        ROUND(SUM(sales), 2)   AS total_spent,
        MAX(order_date)        AS last_purchase
    FROM coffee_orders
    GROUP BY customer_name
)
SELECT
    customer_name,
    order_count,
    total_spent,
    last_purchase,
    CASE
        WHEN order_count >= 5 AND total_spent > 100  THEN 'High Value'
        WHEN order_count >= 3 OR  total_spent > 50   THEN 'Mid Value'
        ELSE 'Low Value'
    END AS customer_segment
FROM customer_stats
ORDER BY total_spent DESC;


-- Q14. Peak Sales Month for Each Coffee Type
WITH monthly_product AS (
    SELECT
        coffee_type_name,
        DATE_FORMAT(order_date, '%Y-%m') AS month,
        ROUND(SUM(sales), 2)             AS revenue
    FROM coffee_orders
    GROUP BY coffee_type_name, DATE_FORMAT(order_date, '%Y-%m')
),
ranked AS (
    SELECT *,
        RANK() OVER (PARTITION BY coffee_type_name ORDER BY revenue DESC) AS rn
    FROM monthly_product
)
SELECT coffee_type_name, month AS peak_month, revenue AS peak_revenue
FROM ranked
WHERE rn = 1;


-- Q15. Month-over-Month Growth Rate
WITH monthly AS (
    SELECT
        DATE_FORMAT(order_date, '%Y-%m') AS month,
        ROUND(SUM(sales), 2)             AS revenue
    FROM coffee_orders
    GROUP BY DATE_FORMAT(order_date, '%Y-%m')
)
SELECT
    month,
    revenue,
    LAG(revenue) OVER (ORDER BY month) AS prev_month,
    ROUND(
        (revenue - LAG(revenue) OVER (ORDER BY month)) /
        NULLIF(LAG(revenue) OVER (ORDER BY month), 0) * 100
    , 2) AS mom_growth_pct
FROM monthly
ORDER BY month;


-- ─────────────────────────────────────────────────────────────
-- BONUS: Quick Data Quality Check
-- ─────────────────────────────────────────────────────────────

-- Check total rows and column nulls
SELECT
    COUNT(*)                            AS total_rows,
    SUM(CASE WHEN order_id     IS NULL THEN 1 ELSE 0 END) AS null_order_id,
    SUM(CASE WHEN customer_name IS NULL THEN 1 ELSE 0 END) AS null_customer,
    SUM(CASE WHEN sales         IS NULL THEN 1 ELSE 0 END) AS null_sales,
    SUM(CASE WHEN order_date    IS NULL THEN 1 ELSE 0 END) AS null_date
FROM coffee_orders;
