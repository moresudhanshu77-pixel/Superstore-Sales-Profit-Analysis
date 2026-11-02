# Superstore Sales & Profit Analysis

**Project Type:** Data Analysis & Interactive Dashboard
**Tools Used:** Microsoft Excel, Microsoft Power BI
**Dataset:** Sample Superstore Sales Dataset

## 1. Project Overview

This project involved analyzing retail sales data using Microsoft Excel and Power BI. The objective was to clean and prepare the raw dataset, calculate key business performance metrics, and develop an interactive dashboard to explore sales, profit, orders, product categories, customer segments, regional performance, and monthly sales trends.

The project followed two main stages: data cleaning and preparation in Excel, followed by data analysis and visualization in Power BI.

## 2. Data Cleaning and Preparation in Microsoft Excel

The original Sample Superstore dataset was provided in a CSV format.

I imported the dataset into Microsoft Excel and prepared it for analysis.

### Cleaning steps performed

**1. Imported the raw dataset**

- Opened the original Sample Superstore CSV file in Excel.
- Reviewed the dataset and its column structure.
- Identified the fields required for sales and profit analysis.

**2. Checked and corrected date formatting**

- Identified inconsistent date formats in the Order Date and Ship Date columns.
- Used Excel's Text to Columns feature to convert the date values into a consistent date format.
- Standardized the displayed date format to DD-MM-YYYY.
- Changed the cell format to General to verify that Excel recognized the dates as actual date values rather than text.
- Confirmed that the dates displayed as Excel serial numbers when formatted as General.

**3. Saved the cleaned dataset**

- Saved a separate cleaned copy of the dataset in Excel format.
- Used the cleaned Excel file as the source for importing data into Power BI.

## 3. Data Analysis and DAX Measures in Power BI

After cleaning the dataset, I imported the Excel file into Power BI and created DAX measures to calculate important business metrics.

### DAX Measures Created

**1. Total Sales**

```dax
Total Sales = SUM('YourTableName'[Sales])
```

Purpose:

- Calculates the total revenue generated from all sales transactions.
- Used to display the Total Sales KPI card.

**2. Total Profit**

```dax
Total Profit = SUM('YourTableName'[Profit])
```

Purpose:

- Calculates the total profit generated from sales.
- Used to display the Total Profit KPI card and analyze profitability across regions.

**3. Total Orders**

```dax
Total Orders = DISTINCTCOUNT('YourTableName'[Order ID])
```

Purpose:

- Counts unique order IDs rather than counting every individual row.
- Provides the total number of distinct orders.
- Used to display the Total Orders KPI card.

*Note: Replace 'YourTableName' with the actual table name in your Power BI model.*

## 4. Dashboard Development in Power BI

Created an interactive sales and profit dashboard using Power BI visualizations.

### Key Performance Indicators (KPIs)

Developed three KPI cards:

- Total Sales
- Total Profit
- Total Orders

### Data Visualizations

**1. Monthly Sales Trend**

- Created a line chart to analyze sales performance over time.
- Used Order Date and Sales to display monthly sales trends.

**2. Sales by Category**

- Created a bar chart to compare sales across product categories.
- Used Category and Sales fields.

**3. Profit by Region**

- Created a horizontal bar chart to compare profit across different regions.
- Used Region and Profit fields.

**4. Sales by Customer Segment**

- Created a column chart to compare sales across customer segments.
- Used Segment and Sales fields.

### Interactive Slicers

Added three slicers to make the dashboard interactive:

- Region: Filter the dashboard by geographical region.
- Category: Filter the dashboard by product category.
- Order Date: Filter the dashboard by a selected date range.

Verified that the slicers interact with the KPI cards and charts, allowing users to explore different subsets of the data.

## 5. Dashboard Formatting and Design

Applied consistent formatting to improve readability and presentation:

- Used a professional blue colour theme for sales visualizations.
- Used green to distinguish profit-related visualizations.
- Applied consistent navy chart titles.
- Formatted KPI cards with clear labels and prominent values.
- Enabled data labels on charts for easier comparison.
- Organized KPI cards, charts, and slicers into a structured dashboard layout.

## 6. Project Outcome

Successfully transformed a raw retail sales dataset into a cleaned, analysis-ready dataset and developed an interactive Power BI dashboard.

The final dashboard enables users to monitor key performance indicators and explore sales and profit performance by month, region, product category, and customer segment using interactive filters.

## 7. Project Workflow Summary

**Raw CSV Dataset → Microsoft Excel → Date Cleaning and Standardization → Cleaned Excel Dataset → Power BI Import → DAX Measures → Data Visualizations → Interactive Dashboard**
