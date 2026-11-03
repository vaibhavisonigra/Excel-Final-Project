# Excel-Final-Project
# Customer Transaction Analysis - Excel Project Documentation

## 📌 Project Overview
This project is an end-to-end **Customer Transaction Analysis System** built in Microsoft Excel. It processes raw transactional data, cleanses and filters dynamic subsets, extracts key business metrics, performs what-if scenario analyses, executes statistical regressions, and visualizes sales performance across various dimensions (Product, Region, Category, and Time).

---

## 📁 Workbook Architecture & Sheet Breakdown

The workbook contains multiple specialized sheets, categorized into Data Processing, Analytical Models, Reporting, and Dashboards:

---

### 1. 📄 `Raw Data` Sheet
* **Description:** Contains the primary source dataset recording transactional records.
* **Key Fields:**
  * `Transaction_ID` (e.g., TRX0184)
  * `Customer_ID` (e.g., CUST0416)
  * `Customer_Name` (e.g., Kathleen Bennett)
  * `Product_Code` / `Product_Name` (e.g., Smartphone, Laptop, Bookshelf)
  * `Category` (e.g., Electronics, Appliances, Furniture)
  * `Quantity`, `Unit_Price`, `Total_Amount`
  * `Payment_Method` (Credit Card, Debit Card, PayPal, Cash)
  * `Region` (North, South, East, West)
  * `Customer_Segment` (Basic, Standard, Premium)
  * `Timestamp` / `Date`

---

### 2. 🔍 `Data Analysis` Sheet
* **Description:** Dynamic workspace for filtered data analysis, lookup operations, and conditional extraction.
* **Key Formulas Applied:**
  * **Dynamic Array Filtering:**
    ```excel
    =FILTER(SelectData, SalesData[Category]="Electronics", "No Data Found")
    ```
  * **Search & Lookup Functions:** `VLOOKUP`, `INDEX-MATCH`, and array formulas for targeted dynamic analysis.

---

### 3. 📊 `Business Insights` Sheet
* **Description:** Summarizes core analytical outcomes and key performance parameters directly calculated from raw transactions.
* **Calculated Metrics:**
  * **Total Sales:** Aggregated sum of all orders (`SUM(Total_Amount)`).
  * **Total Transactions:** Count of total orders (`COUNTA` / `COUNT`).
  * **Average Transaction:** Mean order value (`AVERAGE`).
  * **Highest Transaction Value:** Maximum order purchase (`MAX`).
  * **Lowest Transaction Value:** Minimum order purchase (`MIN`).

---

### 4. 🎯 `Scenario Summary` & `Scenario Analysis` Sheets
* **Description:** What-If analysis tools used to evaluate sensitivity across changing sales environments (Low, Normal, High sales).
* **Components:**
  * **Scenario Manager:** Configured changing cells (`Quantity`, `Unit Price`) to project resulting output variables like `Total Sales`.
  * **Goal Seek Analysis:** Target value searching (e.g., finding required quantity or unit price to achieve specific revenue goals like $1,500).

---

### 5. 📉 `Data Analysis` (Regression Model) Sheet
* **Description:** Statistical regression output examining relationships between quantity, price, and sales revenue.
* **Key Output Sections:**
  * **Regression Statistics:** $R^2$, Adjusted $R^2$, Standard Error, Observations count ($N = 200$).
  * **ANOVA (Analysis of Variance):** Degrees of Freedom (df), Sum of Squares (SS), Mean Square (MS), F-statistic, Significance F.
  * **Coefficients Table:** Intercept, Quantity Coefficient, Unit Price Coefficient with standard errors, p-values, and 95% Confidence Intervals.

---

### 6. 🏆 Data Visualization & Pivot Table Sheets

* **`High Value Customers`:**
  * Summarizes high-value buyers sorting by `Sum of Total_Amount`.
  * **Chart:** High-value customer bar visual comparing individual spenders (e.g., Mark Carter, Edward Mitchell).

* **`Category Wise Sales`:**
  * Pivot summary analyzing revenue across `Electronics`, `Appliances`, and `Furniture`.
  * **Chart:** Column chart highlighting top-grossing category segments.

* **`Region Wise Sales Analysis`:**
  * Evaluates regional revenue distribution across `East`, `West`, `North`, and `South`.
  * **Chart:** Regional comparison horizontal/bar visual chart.

* **`Monthly Sales Trend`:**
  * Time-series analysis grouping transactions by Month (Jan - Dec).
  * **Chart:** Monthly trendline graph demonstrating seasonal spikes and transaction velocity over time.

* **`Product Wise Sales Analysis`:**
  * Breakdown of product-level revenue contribution (Smartphone, Laptop, Office Chair, Bookshelf, Coffee Maker, etc.).
  * **Chart:** Pie / Donut Chart visualizing percentage distribution by product.

---

### 7. 💻 `Dashboard` Sheet
* **Description:** Integrated interactive performance dashboard combining dynamic charts and interactive slicers.
* **Dashboard Elements:**
  * **Monthly Sales Trend Chart:** Line chart displaying timeline trends.
  * **Product-Wise Sales Contribution:** Donut/Pie chart highlighting product split.
  * **Interactive Slicers:** Dynamic filters for `Region` (East, West, North, South) and `Category` allowing real-time slice-and-dice interactivity across all visuals.

---

### 8. 🔤 Data Cleaning & Lookup Verification Sheets

* **`Product Frequency`:** Frequency counts and unit distribution by product items using `COUNTIF`.
* **`List Comparison`:** Cross-checks customer databases across two master lists using double dynamic lookup formulas:
  ```excel
  =INDEX(SalesData[Customer_Name], ROW(A5):A23)
  ```
* **`Customer Abbreviation`:** Cleans customer metadata, extracting initials (`First Initial`, `Last Initial`) and standardized regional/state abbreviation tags.

---

## 🛠️ How to Use This Project

1. **Exploring Raw Data:** Open `Raw Data` to review baseline structure.
2. **Filtering & Searching:** Use `Data Analysis` to see formula-driven filtering (`FILTER` & `INDEX/MATCH`).
3. **Running What-If Analysis:** Go to `Scenario Analysis` to view different sales volume models via Goal Seek & Scenario Manager.
4. **Interacting with Dashboard:** Open the `Dashboard` sheet, use the `Region` and `Category` slicers to filter all visualizations dynamically.
