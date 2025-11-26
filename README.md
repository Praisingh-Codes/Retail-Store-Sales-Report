## Retail Store Sales Analytics & Automated Reporting System

A fully automated Excel-based analytics system designed to help retail businesses track revenue, costs, profit, delivery efficiency, customer patterns, and product performance.

This project integrates data collection, cleaning, processing, analysis, VBA automation, and interactive dashboards into a single plug-and-play Excel solution.

#### Project Structure

Retail Store Sales Report/

├── Data/                          # Raw, cleaned, processed datasets

├── Outcomes/                      # Dashboard and Sales Form output images

│   ├── sales analytics dashboard outcome image.png

│   └── sales form outcome image.png

├── Resources/                     # Icons, UI components, design assets

└── Retail Store Sales Analytics Report.xlsm   # Main automated Excel workbook

#### Dashboard & Form Previews

📊 Sales Analytics Dashboard

![Sales Dashboard](Outcomes/sales%20analytics%20dashboard%20outcome%20image.png)

📝 Automated Sales Entry Form

![Sales Form](Outcomes/sales%20form%20outcome%20image.png)

#### Project Objectives

The client wanted insights into:

✔️ Top-performing products (revenue, cost, profit)

✔️ Monthly & daily trends in revenue, cost, and profit

✔️ Order status patterns (Completed, Pending, Returned)

✔️ Relationship between delivery time and returns

To deliver these insights, the project provides a dynamic Excel Analytics Engine that requires zero technical skills to operate.

#### Key Features

📝 Automated Data Collection (VBA)

- Custom Sales Entry Form built using Form Controls

- Protected layout (only input fields editable)

VBA script automatically:

- Captures and validates inputs

- Calculates revenue & delivery days

- Stores records in a database table

- Resets the form

- Displays success confirmation

Debugged VBA Script (paste-based):

    Sub Submit_PasteValues()
        On Error GoTo CleanFail

        Dim ws As Worksheet
        Dim lastRow As Long

        Set ws = ThisWorkbook.Worksheets("Retail Store Sales")

        If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
            lastRow = 2
            
        Else
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
            
        End If

        ws.Cells(lastRow, "A").PasteSpecial Paste:=xlPasteValues

    CleanExit:
        Application.CutCopyMode = False
        Exit Sub

    CleanFail:
        MsgBox "Nothing to paste or invalid clipboard content.", vbExclamation
        Resume CleanExit    
    End Sub

🧹 Data Cleaning

- Standardized date, text, and numeric formats

- Removed duplicates

- Corrected inconsistent spellings

- Imputed or removed missing values

- Validated spelling, email format, etc.

🔄 Data Processing

- Merged tables into a unified dataset

Added calculated fields:

- Revenue = Quantity × Unit Price

- Delivery Days = Delivered Date - Order Date

- Used VLOOKUP to connect tables

- Applied filtering & sorting for analytics

📈 Data Analysis

Descriptive Statistics

- Identified trends, averages, and distribution patterns

- Checked for outliers and anomalies

Hypothesis Testing (t-Test)

Question:

Do longer delivery times increase return likelihood?

- H0: Delivery time has no effect on returns

- H1: Longer delivery time increases return probability

Result:

✔️ p-value < 0.05 → Reject H0

✔️ Returned orders take ≈1.79 days longer on average

Business Insight:

Long delivery times significantly increase return rates.

📊 Interactive Dashboard

The dashboard includes:

KPIs

- Total Revenue

- Total Cost

- Net Profit

- Total Orders

- Completed vs Returned Orders %

Visualizations

- World map: Revenue by Country

- Bar charts: Category-wise Revenue, Cost, Profit

- Donut & Pie: Order Status & Payment Methods

- Line chart: Monthly Trends

- Column chart: Daily Revenue

Slicers for:

- Month

- Year

- Country

- Product Category


#### Insights & Recommendations

1️⃣ Longer delivery times → higher return rates

Monitor orders >7 days to reduce returns.

2️⃣ Top-performing product categories

Apparel & Electronics dominate margins → prioritize stock & promotions.

3️⃣ Payment method trends

Bank transfer & mobile money account for majority of orders.

4️⃣ Revenue peaks mid-year

Plan marketing and inventory ahead of seasonal spikes.

#### Tech Stack

- Microsoft Excel (Advanced Level)

- Pivot Tables & Pivot Charts

- VBA (Macros)

- Statistical Analysis (t-Test)

- Excel Form Controls

- Slicers, Maps, and Dynamic Charts

#### How to Use

1. Download the project folder.

2. Open Retail Store Sales Analytics Report.xlsm.

3. Enable Macros.

4. Add new records using the Sales Form.

5. Dashboard updates automatically.

6. Use slicers for filtering and analysis.

#### Contributions

Pull requests and improvement suggestions are welcome.
