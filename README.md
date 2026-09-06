# Sales Automation

## Project Overview

this project shows complete sales reporting process.

- Reads multiple excel files
- Cleans messy data
- Remove duplicates
- handles missing values
- Validate quantity and prices
- Generate KPIs
- Create charts
- Generates cleaned files and a consolidated master report


## Technologies Used

- python
- Pandas
- Numpy
- OpenPyXl
- Matplotlib
- OS
- Shutil
- Glob


## How to Run

1. Place Excel files in the `Data/Raw Files` folder.
2. Install the required libraries:
   ```bash
   pip install pandas numpy openpyxl matplotlib
   ```
3. Run the project:
   ```bash
  python Sales_automation.py
   ```

## Project Structure

```
Sales-Automation/
│
├── sales_Automation.ipynb
├── Sales_automation.py
├── README.md
├── requirements.txt
├── LICENSE
├── .gitignore
├── Data/
│   ├── Raw files/
│   ├── Cleaned files/
│   └── Archive/
│
└── Final Reports/
    ├── Chart_folder/
    │   ├── monthly_sales.png
    │   ├── Regional_sales.png
    │   ├── Payment_modes.png
    │   ├── age_group_max_purchased.png
    │   ├── Max_sales_by_Product.png
    │   ├── Maximum_discount.png
    │   ├── Order_stats.png
    │   └── total_sales_by_gender.png
    │
    ├── master_sales.xlsx
    └── Sales_Report.xlsx
```


## Output

The automation generates:

- Cleaned Excel files
- Archived processed input files
- A consolidated master_sales.xlsx file
- A Sales_Report.xlsx containing key sales KPIs
- Sales charts for different business metrics


## Generated KPIs

- Total Sales
- Total Orders
- Average Discount
- Average Order Value
- Highest Sale

## Generated Charts
- Monthly Sales
- Regional Sales
- Payment Methods
- Age Group vs Most Purchased Product
- Maximum Sales by Product
- Maximum Discount by Product
- Order Status
- Total Sales by Gender


## Automation Workflow
```
Monthly Excel Files
       ↓
Data/Raw files/
       ↓
Read Excel Files
       ↓
Data Cleaning
       ↓
Validation & Transformation
       ↓
Data/Cleaned files/
       ↓
Archive Original Files
       ↓
Combine Sales Data
       ↓
master_sales.xlsx
       ↓
KPIs + Charts
       ↓
Final Reports
```

## Future Improvements

- logging
- error handling
- configuration file
- database loading

## Author

Akshay










