
# 🛒 E-Commerce Sales Analysis Project

## 📊 Overview
A comprehensive Excel-based analysis project focused on e-commerce sales data to derive actionable business insights. This project uncovers sales trends, product performance, discount impacts, and profitability, enabling strategic decision-making for business growth.

---

## 🎯 Project Goals
- Analyze e-commerce sales data across regions, markets, and categories.
- Understand the effect of discounts on revenue and profitability.
- Identify top and bottom-performing products.
- Offer insights and recommendations to improve overall sales performance.

---

## 🛠 Tools & Technologies
- **Microsoft Excel**
  - Data Cleaning
  - Formula-based analysis
  - Pivot Tables
  - Dashboard Creation

---

## 🗃️ Dataset Description
The dataset includes:
- **Row ID, Order ID, Order Date**
- **Customer ID, Segment**
- **City, State, Country**
- **Region, Market**
- **Category, Subcategory**
- **Product Name**
- **Quantity, Sales, Discount, Profit**

---

## 🧹 Data Cleaning & Preparation
- Removed duplicate entries and irrelevant columns.
- Standardized date format to DD-MM-YYYY.
- Validated and corrected numeric fields (Sales, Discount, Profit).
- Ensured there were no missing or invalid values.

---

## 🔍 Functional Excel Analysis

### 1. **Basic Aggregates**
- **Total Sales:**  
  `=SUM(Dataset!O2:O51189)`
- **Total Quantity Sold:**  
  `=SUM(Dataset!N2:N51189)`
- **Total Profit:**  
  `=SUM(Dataset!Q2:Q51189)`
- **Average Discount:**  
  `=AVERAGE(Dataset!P2:P51189)`
- **Average Order Value (AOV):**  
  `=Total Sales / Number of Orders`

### 2. **Profit Margin**
- **Formula:**  
  `=Total Profit / Total Sales`

### 3. **Top-Selling Product**
- **Formula:**  
  `=INDEX(I4:I3580, MATCH(LARGE(J4:J3580, ROW(A1)), J4:J3580, 0))`

### 4. **Lowest-Selling Product**
- **Formula:**  
  `=INDEX(I5:I3580, MATCH(SMALL(J5:J3580, ROW(A1)), J5:J3580, 0))`

### 5. **Product-Wise Sales & Profit**
- **Sales per Product:**  
  `=SUMIF(Dataset!L2:L51189, S3, Dataset!O2:O51189)`
- **Profit per Product:**  
  `=SUMIF(Dataset!L2:L51189, S3, Dataset!Q2:Q51189)`
- **Product Count per Category:**  
  `=COUNTIF(Dataset!L2:L51189, S3)`

---


## 📊 Dashboard Highlights
- Total Sales, Profit, Units Sold, Avg Order Value
- Region- and City-wise Sales Trends
- Monthly Sales Trendline (Peak in December)
- Discount vs Profitability Relationship
- Best and Worst Performing Products
- Orders by Marketplace (Asia Pacific, USCA)
  
![image](https://github.com/user-attachments/assets/5fc84430-2b97-4b0c-9e6f-25d3681c666a)

![image](https://github.com/user-attachments/assets/9eddb24c-3d75-45d0-a9bd-92917199a9ff)






## 📌 Key Business Insights

### 📈 General Performance
- **Total Sales:** $33.04M  
- **Total Units Sold:** 277K  
- **Total Customers:** 17,415  
- **Average Order Value:** $127  
- **Total Profit:** $1.064M  
- **Profit Margin:** 16.36%

### 🏙️ Top Markets & Segments
- **Top Cities:** NYC ($117.3K), LA ($107.5K)
- **Top Market:** Africa (28.15%)
- **Top Segment:** Corporate ($3.83M)

### 🛍️ Product Performance
- **Top Product:** Herbal Essences Bio ($65.4K)
- **Lowest Product:** Dove Shea Butter Body Wash (3 units)
- **Most Profitable Category:** Body Care ($590K)
- **Loss-Making Category:** Home & Accessories (-$56K)

---



## 💡 Recommendations

## Scale High-Performing Market
- Increase inventory and promotions in Africa, LATAM, and US cities like NYC and LA.
- Expand partnerships with retailers in top regions.

## Revamp or Phase Out Low-Selling Products
- Review poor performers (e.g., Stila Eyeshadow, Burt’s Bees Lemon) for delisting or repositioning.
- Focus R&D on expanding best-performing lines like shampoos and body care.

## Segment-Based Targeting
- Design loyalty programs and bundles tailored to Corporate buyers.
- Run targeted campaigns to uplift Consumer and Home Office segments.

## Seasonal Promotion Planning
- Capitalize on Q4 peak sales with pre-planned campaigns from October onwards.
- Launch holiday bundles and limited editions.

## Optimize Discounting Strategy
- Avoid blanket discounts; test with A/B pricing.
- Promote value-based pricing and loyalty rewards to improve profit margins.

---
## E-Commerce Summary Report (Automated with Excel VBA)

To streamline stakeholder reporting, I built an automated one-page performance report in Excel using VBA scripting.

✅ Key Metrics Included:

    Total Sales: $33.04M

    Profit & Margin: $1.064M (16.36%)

    Top Products & Regions

    Sales by City & Customer Type

This report is auto-generated via VBA, reducing manual effort and enabling quick, consistent insight sharing.

🔧 Tools Used: Excel, VBA (Macros), Pivot Tables

![image](https://github.com/user-attachments/assets/53780984-29e5-4710-8f11-bcec0777c83b)


## 📁 Files Included
- `E-Commerce Sales Analysis Report.pdf`: Full report with analysis, insights, and recommendations.
- `README.md`: GitHub documentation (this file).

---

## 👩‍💼 Author
**Neha Jade**

---

## 📬 Contact
If you'd like to connect, discuss the project, or collaborate, feel free to reach out via GitHub or email.

---

