Excel Assignment 1 – Data Exploration

Description

This is my first assignment in the Data Analytics course Module 1.
I worked on a Product Dataset in Microsoft Excel that contains
information about 34 products including their Product ID, Product Name,
Brand Name, Price, Quantity, and Category.
I used basic Excel formulas and functions to explore and analyze
the data in a simple and organized way.

Dataset Details

- Total Products : 34
- Unique Products : 32 (2 duplicate entries found)
- Columns : Product ID, Product Name, Brand Name, 
            Price ($), Quantity, Category
- Categories : Electronics, Fashion, Kitchen, 
               Outdoor, Accessories

Tasks Completed

Task 1 – Basic Calculations
Used SUM, COUNT, AVERAGE to summarize the price data.

Task 2 – Min and Max
Used MIN and MAX to find the lowest and highest priced products.

Task 3 – IF Function
Added a new column called Price Range to label each product
as High Price (>=500) or Standard Price (<500).

Task 4 – SUMIF and COUNTIF
Used SUMIF to get total price of Electronics category.
Used COUNTIF to count products priced below $100.

Task 5 – Text Functions
Used LEFT, RIGHT and MID to extract Day, Country Code
and Month from the Product ID column.

Formulas Used

| Task | Formula |
|------|---------|
| Total Price | =SUM(D2:D35) |
| Count of Products | =COUNTA(B2:B35) |
| Average Price | =AVERAGE(D2:D35) |
| Minimum Price | =MIN(D2:D35) |
| Maximum Price | =MAX(D2:D35) |
| Price Range | =IF(D2>=500,"High Price","Standard Price") |
| Electronics Total | =SUMIF(F2:F35,"Electronics",D2:D35) |
| Products below $100 | =COUNTIF(D2:D35,"<100") |
| Day | =LEFT(A2,2) |
| Country Code | =RIGHT(A2,2) |
| Month | =MID(A2,4,3) |

Key Findings

- Total price of all products is $10,100
- Average price per product is $297.06
- Cheapest product is Adidas T-shirt at $30
- Most expensive product is Dell Laptop at $1,000
- Electronics category alone costs $8,050 out of $10,100
- 11 products are priced below $100
- 8 products are classified as High Price

Files in This Repository
Excel Assignment Module 1 - Data Analytics PDF
Excel Assignment Module 1 - Data Analytics xlsx

Author
Dhanalakshmi A
Data Analytics Student

License
This project is created for educational and assignment 
purposes only.

Acknowledgments

- Microsoft Excel
- GitHub
