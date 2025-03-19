# Sales Analysis Using Excel

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
1. <b>Descriptive Statistical Analysis:</b>A statistical summary of key variables—including Delivery Time, Total Cost, Sales Revenue, Net Profit, Quantity, and Unit 
   Price—was generated using the Descriptive Statistics function within the Data Analysis Toolpak.

   By using descriptive statistics, businesses can gain a clear overview of operational performance and identify areas for improvement before conducting deeper analysis.

2. <b>T-Test Analysis:</b>A t-test was conducted to examine the relationship between Delivery Time and Order Status, testing whether delivery time significantly impacts 
   order completion.
   
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

## Automated Data Entry Form Implementation
To improve operational efficiency, a macro-enabled data entry form was developed, allowing seamless entry of new sales records while maintaining data integrity.

### Development Process:

1️⃣ Form Layout Design:

- A structured form was designed with predefined input fields corresponding to key dataset attributes.
- Data validation techniques, including dropdown selections, were incorporated to minimize entry errors.

2️⃣ Macro-Driven Automation:

- The Record Macro feature was used to automate the process of appending new entries to the dataset.
- A Submit button was linked to the macro to facilitate data insertion

3️⃣ Error Handling & Confirmation Popup Message:

- An issue was initially encountered while appending data, which was resolved using the following VBA script:
  
   `Sheets("Retail Store Sales").Select
   Dim lastRow As Long
   lastRow = Sheets("Retail Store Sales").Cells(Rows.Count, "A").End(xlUp).Row + 1
   Range("A" & lastRow).PasteSpecial Paste:=xlPasteValues`

- To enhance user experience, a confirmation popup message was implemented using VBA, ensuring users receive immediate feedback upon successful data submission.
  
   `MsgBox "Submission Successful", vbInformation, "Confirmation"`

The automated data entry system significantly enhances efficiency, reducing manual effort and minimizing errors.

## KPI Calculation & Sales Performance Analysis Using Pivot Tables
A pivot table was created to compute:

1️⃣ <b>Revenue & Profit Analysis:</b>
- Total Revenue, Total Cost, and Net Profit calculations
- Filtering by Order Status
- Customer count analysis

2️⃣ <b>Order Status Analysis:</b>
- A pivot table displaying Completed vs. Returned Orders was created.
- The percentage of order status was calculated and visualized using a pie chart.

3️⃣ <b>Monthly Sales Trend Analysis:</b>
A Month Table Pivot Table was designed to analyze:
- Revenue, Total Cost, and Net Profit for Each Month
- A support column was added to include row numbers for reference.

4️⃣ <b>Calculation of Key Metrics for MoM Sales Analysis :</b>

To enhance financial analysis, the following additional measures were computed:

<b>Previous Month:</b> `=MATCH(B43,B26:B37,0)-1` (B43->Lookup value from pivot table with only month values, B26:B37-> Pivot table containing Monthly Sales Analysis)

<b>Previous Month Name:</b> `=IFERROR(VLOOKUP(C43,A26:E37,2,0),0)` (C43->Previous Month Value, A26:E37->Monthly Sales table Consisting of support Column,2->Column to lookup) 

<b>Current Value:</b> Summarized total revenue, total costs, net profit, total order for the selected period.

<b>Previous Month Value:</b> Extracted revenue, total costs, net profit, total order from the prior month.

<b>Value Difference:</b> `=Current Value - Previous Month Value`

<b>Percentage Difference:</b> (Value Difference / Previous Month Value) 

 - Formatted Using `=IF(H43>0,"+"&TEXT(H43,"0.0%"),TEXT(H43,"0.0%"))'(H43 corresponding to Percentage Difference)

<b>Final Value vs. Last Month (LM):</b> Comparison metric to highlight changes over time.

 - Calculated using `=H44&"|"&G44` (H44->Percentage Difference and G44->Value Difference)
 - Conditional formatting applied green for positive value and red for negative value.

## Interactive Dashboard Creation

The final step involved designing an interactive dashboard to visualize the key performance indicators effectively. The dashboard includes:
1. Dynamic charts displaying total revenue, net profit, and total cost, total orders trends over time.
2. Slicers and interactive elements to enhance usability.
3. Order Status Breakdown visualized using donut charts.
4. Revenue by Country mapped for global sales insights.
5. Orders by Payment Method presented in a pie chart.
6. Revenue, Cost and Profit by category visulized by stacked column chart.

## Observations & Results

1. <b>Product Performance:</b>
- Apparel and Electronics generate the highest sales revenue but also incur higher costs.
- Books and Groceries show consistent sales.

2. <b>Customer Demographics:</b>
- The majority of customers are from Nigeria, Australia, and the United States, with Nigeria contributing the highest number of orders.
- Customers from China and the United Kingdom also show significant sales activity.

3. <b>Payment Methods:</b>
- Credit Card and Mobile Money are the most popular payment methods, followed by Cash and Bank Transfer.
- The choice of payment method varies by country, with Mobile Money being more prevalent in Nigeria and Credit Card in Australia and the United States.

4. <b>Order Status:</b>
- Most orders are Completed, but there is a notable number of Returned orders, particularly in the Electronics and Apparel categories.
- The return rate is higher for high-value items like Smartphones and Headphones.

5. <b>Sales Revenue and Net Profit:</b>
- The highest sales revenue is generated from Electronics, particularly Smartphones and Laptops
- Despite high sales revenue, the net profit margin is lower for Electronics due to higher total costs (e.g., production, shipping).
- Apparel and Home Decor show healthier profit margins compared to Electronics.

<b>To improve sales performance and profitability, the following actions are recommended:</b>

<b>Optimize Inventory:</b> Focus on high-margin products like Apparel and Home Decor while managing costs for Electronics.

<b>Enhance Customer Experience:</b> Reduce return rates by improving product quality and ensuring accurate product descriptions.

<b>Streamline Delivery:</b> Improve delivery efficiency, especially for international orders, to enhance customer satisfaction.

<b>Targeted Marketing:</b> Leverage seasonal trends and customer preferences to run targeted marketing campaigns, particularly during peak sales months.

This report serves as a foundation for strategic planning and decision-making to achieve sustainable growth and profitability in the retail business.


  
