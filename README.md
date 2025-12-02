
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

2. **Profit Margin** 
  `=Total Profit / Total Sales`
<img width="750" height="317" alt="Screenshot 2025-12-01 090444" src="https://github.com/user-attachments/assets/f514471d-fd22-40de-b614-be9b171dc877" />

### 3. **Top-Selling Product**
- **Formula:**  
  `=INDEX(I4:I3580, MATCH(LARGE(J4:J3580, ROW(A1)), J4:J3580, 0))`

### 4. **Lowest-Selling Product**
- **Formula:**  
  `=INDEX(I5:I3580, MATCH(SMALL(J5:J3580, ROW(A1)), J5:J3580, 0))`
<img width="1275" height="224" alt="Screenshot 2025-12-01 090647" src="https://github.com/user-attachments/assets/5a429e4a-3a4d-49a5-9666-8904f877149e" />

### 5. **Product-Wise Sales & Profit**
- **Sales per Product:**  
  `=SUMIF(Dataset!L2:L51189, S3, Dataset!O2:O51189)`
- **Profit per Product:**  
  `=SUMIF(Dataset!L2:L51189, S3, Dataset!Q2:Q51189)`
- **Product Count per Category:**  
  `=COUNTIF(Dataset!L2:L51189, S3)`
  
<img width="739" height="211" alt="Screenshot 2025-12-01 090655" src="https://github.com/user-attachments/assets/f9c3a315-6a98-44a0-bc88-e791be3d7e37" />

---
# 📊 Lookup Analysis (Excel Formula Preview)

This project includes a dedicated **Lookup Analysis** section that demonstrates how Excel lookup formulas extract insights from the dataset.

---

## 🔎 VLOOKUP Analysis
| Product | Category (via VLOOKUP) |
|--------|--------------------------|
| L'Oréal Paris Preference – Light Warm Brown | `=VLOOKUP(A4, A3:G51183, 3, FALSE)` |

---

## 🔎 HLOOKUP Analysis
| Product | Sales (via HLOOKUP) |
|--------|-----------------------|
| MAC 210 Precise Eye Liner Brush | `=HLOOKUP("Sales", A2:G51118, 3, FALSE)` |

---

## 🔎 XLOOKUP Analysis
| Product | Profit (via XLOOKUP) |
|--------|------------------------|
| L'Oréal Paris Preference – Light Warm Brown | `=XLOOKUP("L'Oréal Paris Preference - Light Warm Brown", A2:A51183, G2:G51183)` |

---

## 📘 Excel Formulas Used
```excel
=VLOOKUP(I4, A3:G51183, 3, FALSE)
=HLOOKUP("Sales", A2:G51118, 3, FALSE)
=XLOOKUP("L'Oréal Paris Preference - Light Warm Brown", A2:A51183, G2:G51183)
=VLOOKUP(L4, A3:G51183, 2, FALSE)

<img width="1075" height="361" alt="image" src="https://github.com/user-attachments/assets/fe41f148-37e3-4368-bb8c-ef03277cedbb" />


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
# 📊 Automated Sales MIS Report (Excel VBA)

Automates generation of an **E-Commerce MIS (Management Information System) report** in Excel using VBA.  
Provide your transaction dataset, click a button, and get a clean Daily / Weekly / Monthly sales analysis.

---

## ✅ Features

- Generate **Daily**, **Weekly**, and **Monthly** reports  
- Permanent buttons on a dedicated sheet — do not overlap the report  
- Clean summary output (no raw transaction dump)  
- Calculates these KPIs and sections:
  - **Sales Overview**
    - Total Sales  
    - Total Quantity Sold  
    - Total Customers  
    - Average Discount
  - **Market Performance**
    - Market Sales Distribution (Top Markets)  
    - Top Cities by Sales  
    - Sales by Customer Type
  - **Product Performance**
    - Total Profit  
    - Profit Margin  
    - Average Order Value (AOV)  
    - Top-Selling Products  
    - Worst-Selling Products  
    - Top Product Categories

---
<img width="1048" height="832" alt="image" src="https://github.com/user-attachments/assets/27a46928-7607-47ab-8a2c-6df62f8b8a16" />



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

