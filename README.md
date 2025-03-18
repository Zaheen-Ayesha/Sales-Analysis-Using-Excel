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

## Data Processing

1. To enhance time-based analysis, key date components were extracted, and delivery time was calculated:
The TEXT function in Excel was used:
- `=TEXT(OrderDate, "YYYY") for Year`
- `=TEXT(OrderDate, "MMM") for Month`
- `=TEXT(OrderDate, "DDD") for Day`

Extracting date components enables time-series analysis, helping identify seasonal trends, peak order periods, and yearly/monthly performance variations.

The insights from time series analysis supports forecasting, resource allocation, and strategic planning, ultimately improving service delivery and customer 
satisfaction. 

2. To assess profitability, key financial metrics were computed:

- <b>Total Costs:</b> Calculated using `=ROUND([Unit Price] * [Quantity] * VLOOKUP([Product Name], Table2[#All], 2, FALSE), 0)`, incorporating unit price, quantity, and cost percentage.
- <b>Sales Revenue:</b> Derived as `Unit Price * Quantity`, representing total earnings before costs.
- <b>Net Profit:</b> Computed as `Sales Revenue - Total Costs`, reflecting actual profit.

Profitability analysis helps assess business performance by identifying high-margin products and cost-intensive items, enabling strategic pricing and cost control. Understanding cost structures supports data-driven decisions for pricing, promotions, and overall financial optimization.

## Data Analysis
1. <b>Descriptive Statistical Analysis:</b>A statistical summary of key variables—including Delivery Time, Total Cost, Sales Revenue, Net Profit, Quantity, and Unit Price—was generated using the Descriptive Statistics function within the Data Analysis Toolpak.

By using descriptive statistics, businesses can gain a clear overview of operational performance and identify areas for improvement before conducting deeper analysis.

2. <b>T-Test Analysis:</b>A t-test was conducted to examine the relationship between Delivery Time and Order Status, testing whether delivery time significantly impacts order completion.
   
- <b>Hypothesis Statement:</b>
  - Null Hypothesis (H₀): Delivery time does not influence whether an order is returned.
  - Alternative Hypothesis (H₁): Orders that take longer delivery time are more likely to be returned.

- A t-Test: Two-Sample Assuming Unequal Variances was conducted to examine whether delivery time influences order returns.
  
  <b>Key Results:</b>
  - <b>Mean Delivery Time:</b>
    - Completed Orders: 6.98 days
    - Returned Orders: 8.77 days
  - <b>Variance:</b>
    - Completed Orders: 12.70
    - Returned Orders: 16.07
  - <b>Observations:</b>
    - Completed Orders: 287
    - Returned Orders: 268
  - <b>t-Statistic:</b> -5.53
  - <b>p-value (two-tailed):</b> 4.96 × 10⁻⁸ (significant)
  - <b>Critical t-value (two-tailed):</b> 1.96

 <b>Interpretation:</b>
 - The p-value is much smaller than 0.05, indicating a statistically significant relationship between Delivery time and order status.
 - The negative t-statistic suggests that returned orders tend to have longer delivery times.
 -  t-statistic (-5.53) is less than -1.96, we reject the null hypothesis and conclude that longer delivery times are significantly associated with higher return rates.
 -  On average, returned orders take approximately 1.8 days longer to deliver compared to completed orders.
 -  A return rate of ~48% (268 out of 555 orders) is considerably high and indicates critical business challenges that need to be addressed.
   
    -  A high return rate can lead to:
       - Increased logistics costs (reverse shipping, restocking).
       - Lost revenue due to refund processing.
       - Higher operational burden (handling returns, quality checks). 

  
