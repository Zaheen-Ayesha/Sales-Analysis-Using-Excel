# Sales-Analysis-Using-Excel

The primary objective of this sales analysis project is to analyze and evaluate the sales performance of a retail store and address the following key business questions posed by the client:
- Which products generate the highest revenue and profit, and what are their associated costs ?
- What are the trends in financial metrics such as revenue, cost, and profit over time ?
- How do order statuses(Completed, Returned) evolve, and what patterns can we identify?

The analysis aims to identify trends, patterns, and insights that can help in making data-driven decisions to improve sales, optimize inventory, and enhance customer satisfaction. 

The dataset contains detailed sales transactions for a retail store, including information such as customer names, product categories, product names, order dates, delivery dates, quantities, unit prices, order status, country, payment methods. The data spans multiple countries and product categories, including Electronics, Books, Apparel, Groceries, and Home Decor.

## Data Cleaning

1. <b>Standardizing Formats:</b>
  - The Order Date and Delivery Date columns were standardized to maintain uniform formatting, ensuring consistency for analysis and reporting.
2. <b>Removing Duplicates:</b>
  - Conditional formatting was applied to the Order ID column (the primary key) to highlight potential duplicate entries.
  - The Remove Duplicates feature was utilized by selecting all columns, successfully identifying and eliminating one duplicate record, thereby improving data accuracy.
3. <b>Handling Missing Values:</b>
  - One missing value in the Unit Price column was detected using conditional formatting to format the blank cells.
  - The missing value was imputed with the mean of the column, calculated through the Descriptive Statistics tool under Excel’s Data Analysis ribbon, ensuring data 
    completeness without introducing bias.
