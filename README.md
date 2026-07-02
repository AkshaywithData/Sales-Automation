# Sales Automation

## Project Overview

this project shows complete sales reporting process.

- Reads multiple excel files
- Clean messy data
- Remove duplicates
- handling missing values
- Validate quantity and prices
- Generate KPIs
- Create charts
- Saves cleaned combined report


## Technologies Used

- python
- Pandas
- Numpy
- Openyxl
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
  Sales_automation.py
   ```

## Project Structure

```text
Sales-Automation/
├── Sales_Automation.py
├── Sales_Automation.ipynb
├── README.md
├── requirements.txt
├── Data/
│   ├── Raw Files/
│   ├── Cleaned Files/
│   └── Archive/
└── Final Reports/
    ├── Chart_Folder/
    │   ├── Monthly_sales.png
    │   ├── Order_status.png
    │   └── Regional_sales.png
    ├── Master_Sales.xlsx
    └── Sales_Report.xlsx
```


## Author

Akshay










