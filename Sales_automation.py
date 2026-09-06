import pandas as pd
import numpy as np
import os
import shutil
import glob
from openpyxl import workbook

def clean_data(df):
    
    df["age"] = df["age"].fillna(
    df.groupby("gender")["age"].transform("median")
    )
    
    df["order_date"] = pd.to_datetime(df["order_date"])
    
    df["quantity"] = df["quantity"].fillna(
    df.groupby("product_name")["quantity"].transform("median"))
    
    df["quantity"] = df["quantity"].astype(int)
    
    df["discount_pct"] = df["discount_pct"].fillna(df.groupby("product_name")["discount_pct"].transform("median"))
    
    df["customer_satisfaction"] = df["customer_satisfaction"].fillna(df.groupby("product_name")
                                                                 ["customer_satisfaction"].transform("median"))
    
    df["days_to_ship"] = df["days_to_ship"].fillna(df.groupby("product_name")
                                                                 ["days_to_ship"].transform("median"))
    df["days_to_ship"] = df["days_to_ship"].astype(int)
    
    df.loc[df["days_to_ship"] <1, "days_to_ship"] = np.nan
    
    df.loc[df["days_to_ship"] > 20, "days_to_ship"] = np.nan
    
    df["gender"] = df["gender"].replace({"m" : "Male",
                                     "f" : "Female", 
                                     "M": "Male",
                                     "F" : "Female",
                                     "female" : "Female",
                                     "male" : "Male",
                                     "FEMALE" : "Female",
                                     "MALE" : "Male"})
    
    df.drop_duplicates(inplace=True)
    
    df.groupby("customer_id")["customer_name"].nunique().loc[lambda x: x > 1]
    
    name_conflict = (
    df.groupby("customer_id")["customer_name"]
      .transform("nunique") > 1)
    
    df["customer_name_conflict"] = name_conflict
    
    df["age"] = df["age"].astype(int)
    
    df.loc[(df["age"] > 100) | (df["age"] < 1), "age"] = np.nan
    
    df.loc[df["quantity"] < 0, "quantity"] = np.nan
    
    mask = df["quantity"] > 500 
    
    df.loc[df["quantity"] > 100, "quantity"] = np.nan
    
    df["quantity"] = df["quantity"].fillna(
    df.groupby("product_name")["quantity"].transform("median")
    )
    
    df["quantity"] = df["quantity"].astype(int)

    df.loc[mask, "sales_amount"] = (
    df.loc[mask, "quantity"] *
    df.loc[mask, "unit_price"] *
    (1 - df.loc[mask, "discount_pct"])
    )
    
    df["order_status"] = df["order_status"].replace({"delivered" :"Delivered",
                                                   "pending" : "Pending",
                                                   "shipped" : "Shipped",
                                                   "cancelled" : "Cancelled"})
    return df

def save_file(df, file):   
    
    os.makedirs("Data/Cleaned files", exist_ok = True)  
    
    new_file = os.path.basename(file).replace(".xlsx", "_new.xlsx")

    output_path = os.path.join("Data/Cleaned files", new_file)

    df.to_excel(output_path, index=False)
    
def archive_file(file):
    
    os.makedirs("Data/Archive", exist_ok= True)
    
    archive_path = os.path.join("Data/Archive", os.path.basename(file))
    
    shutil.move(file, archive_path)
    
def Update_master(data):

    master  = "Final Reports/master_sales.xlsx"
    
    try:
        finaldf = pd.concat(data, ignore_index=True)
        
        if not os.path.exists(master):
            
            finaldf.to_excel(master, index = False)
            
            print("master created")
        
        else:
            old = pd.read_excel(master)
            
            if not old.empty:
                finaldf = pd.concat([old, finaldf], ignore_index=True)
                
            finaldf.to_excel(master, index=False)
            print("Master updated")
                
    except ValueError:
        print("No new files")
        
        finaldf = pd.read_excel(master)   
    return finaldf     

def Generate_charts(finaldf):

    os.makedirs("Final Reports/Chart_folder", exist_ok=True)
    
    # monthly sales
    finaldf["month"] = finaldf["order_date"].dt.month

    monthly_sales = finaldf.groupby("month")["sales_amount"].sum()

    import matplotlib.pyplot as plt

    plt.bar(
        monthly_sales.index,
        monthly_sales.values
    )

    plt.title("Monthly_Sales")

    plt.savefig(
        os.path.join("Final Reports/Chart_folder", "monthly_sales.png"),
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

    #regional sales

    Regional_sales= finaldf.groupby("region")["sales_amount"].sum() 
        
    plt.bar(
        Regional_sales.index,
        Regional_sales.values
    )

    plt.title("Regional_sales")

    plt.savefig(
        os.path.join("Final Reports/Chart_folder", "Regional_sales.png"),
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

    #payment methods
    payments = finaldf["payment_method"].value_counts()

    plt.bar(
        payments.index,
        payments.values
    )
    plt.title("Paymnent_methods")
    plt.savefig(
        os.path.join("Final Reports/Chart_folder", "Payment_modes.png"),
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

    #age group max purchase
    bins = [0, 18, 25, 35, 45, 60, 100]
    labels = ["<18", "18-25", "26-35", "36-45", "46-60", "60+"]

    finaldf["age_group"] = pd.cut(finaldf["age"], bins=bins, labels=labels)

    finaldf.groupby(["age_group", "product_name"]).size()

    age_group_max = finaldf.groupby("age_group")["product_name"]\
    .agg(lambda x: x.value_counts().idxmax())
    
    plt.figure(figsize = (12,5))
    
    plt.plot(age_group_max.index, age_group_max.values)

    plt.title("age_group max purchased")
    
    plt.xlabel("Age Group")
    plt.ylabel("Products")
    plt.savefig(
       os.path.join("Final Reports/Chart_folder", "age_group max purchased.png"),
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()
    
    #max sales amount by a product
    product_max = finaldf.groupby("product_name")["sales_amount"] .max().sort_values(ascending=False)

    plt.figure(figsize=(12, 5))

    plt.plot(product_max.index, product_max.values)

    plt.title("Maximum Sales Amount by Product")
    plt.xlabel("Product Name")
    plt.ylabel("Maximum Sales Amount")

    plt.savefig(
        os.path.join("Final Reports/Chart_folder", "Max sales by Product.png"),
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

    #max discount desc
    discount_max =  finaldf.groupby("product_name")["discount_pct"].max().sort_values(ascending = False)
    plt.figure(figsize=(12, 5))
    plt.plot(
        discount_max.index,
        discount_max.values
    )
    plt.title("Discount_max percent")
    plt.savefig(
        os.path.join("Final Reports/Chart_folder", "Maximum discount.png"),
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

    # orders_types_counts
    orders_stat_info =  finaldf["order_status"].value_counts()

    plt.bar(orders_stat_info.index, orders_stat_info.values)
    plt.title("Order_Stats")
    plt.savefig(
        os.path.join("Final Reports/Chart_folder", "Order_stats.png"),
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

    #gender_specific sales
    sales_by_gender = finaldf.groupby("gender")["sales_amount"].sum().sort_values(ascending = False)

    plt.bar(sales_by_gender.index,
            sales_by_gender.values)

    plt.title("Sales")
    plt.savefig(
        os.path.join("Final Reports/Chart_folder", "total sales by gender.png"),
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()
    
def Generate_kpis(finaldf):
    #KPIS total sales, avg discounts, avg monthly orders, hisghest sale, total orders

    total_sales = finaldf["sales_amount"].sum()

    total_orders = finaldf["order_id"].nunique()

    average_discount = finaldf["discount_pct"].mean()

    average_order = finaldf["sales_amount"].mean()

    highest_sale = finaldf["sales_amount"].max()

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"

    kpis = [
        ("Total Sales", total_sales),
        ("Total Orders", total_orders),
        ("Average Discount", average_discount),
        ("Average Order Value", average_order),
        ("Highest Sale", highest_sale)
    ]

    ws["A1"] = "KPI"
    ws["B1"] = "Value"

    for row, (name, value) in enumerate(kpis, start=2):
        ws.cell(row=row, column=1).value = name
        ws.cell(row=row, column=2).value = value
        
    wb.save("Final Reports/Sales_Report.xlsx")
    
files = glob.glob("Data/Raw files/*.xlsx")

data = []
       
for file in files:
    
    df = pd.read_excel(file)
    
    df = clean_data(df)
    
    save_file(df, file)
    
    archive_file(file)
    
    data.append(df)
    
finaldf = Update_master(data)     

Generate_charts(finaldf)

Generate_kpis(finaldf)







 


    

    
    
    


