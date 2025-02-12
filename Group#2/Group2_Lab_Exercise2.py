import pandas as pd

# Load the Excel file
file_path_xlsx = "/content/sample_data/Sample - Superstore - Copy.xls"
xls = pd.ExcelFile(file_path_xlsx)

orders_df = pd.read_excel(xls, sheet_name="Orders")
returns_df = pd.read_excel(xls, sheet_name="Returns")

orders_df["Order Date"] = pd.to_datetime(orders_df["Order Date"])
orders_df["Ship Date"] = pd.to_datetime(orders_df["Ship Date"])

orders_df.fillna({"Customer Name": "Unknown", "Region": "Unknown"}, inplace=True)
orders_df["Sales"].fillna(orders_df["Sales"].median(), inplace=True)
orders_df["Quantity"].fillna(orders_df["Quantity"].median(), inplace=True)

orders_df.drop_duplicates(inplace=True)

orders_df = orders_df.merge(returns_df, on="Order ID", how="left")
orders_df["Returned"] = orders_df["Returned"].fillna("No")

orders_df["Profit Margin (%)"] = (orders_df["Profit"] / orders_df["Sales"]) * 100
orders_df["Discount Impact"] = orders_df["Discount"] * orders_df["Sales"]

orders_df = orders_df.sort_values(by="Order Date")
orders_df["Previous Sales"] = orders_df["Sales"].shift(1)
orders_df["Sales Growth (%)"] = ((orders_df["Sales"] - orders_df["Previous Sales"]) / orders_df["Previous Sales"]) * 100
orders_df["Sales Growth (%)"].fillna(0, inplace=True)

orders_df["Year-Month"] = orders_df["Order Date"].dt.strftime('%Y-%m')

total_orders = len(orders_df)
returned_orders = orders_df["Returned"].value_counts().get("Yes", 0)
orders_df["Return Rate (%)"] = (returned_orders / total_orders) * 100

orders_df["Average Order Value (AOV)"] = orders_df["Sales"].sum() / total_orders

operational_table = orders_df[[
    "Order Date", "Region", "City", "Category", "Sub-Category", "Sales", "Quantity", "Discount",
    "Profit Margin (%)", "Sales Growth (%)"
]].copy()

operational_table.rename(columns={
    "Order Date": "Date",
    "Category": "Product Category",
    "Sub-Category": "Sub-Category",
    "Sales": "Total Sales",
}, inplace=True)

# Select first 10 rows
operational_table_sample = operational_table.head(10)

top_performing_products = (
    orders_df.groupby(["Region", "Year-Month", "Sub-Category"])["Sales"]
    .sum()
    .reset_index()
    .sort_values(by=["Region", "Year-Month", "Sales"], ascending=[True, True, False])
)

top_products_dict = top_performing_products.groupby(["Region", "Year-Month"])["Sub-Category"].first().to_dict()

executive_report = orders_df.groupby(["Region", "Year-Month"]).agg(
    Total_Sales=("Sales", "sum"),
    Total_Profit=("Profit", "sum"),
    Profit_Margin_Percent=("Profit Margin (%)", "mean"),
    Sales_Growth_Percent=("Sales Growth (%)", "mean"),
).reset_index()

executive_report["Top-Performing Products"] = executive_report.apply(
    lambda row: top_products_dict.get((row["Region"], row["Year-Month"]), "N/A"), axis=1
)

executive_report["Discount Impact"] = orders_df["Discount Impact"].sum()
executive_report["Return Rate (%)"] = (returned_orders / total_orders) * 100
executive_report["Average Order Value (AOV)"] = orders_df["Sales"].sum() / total_orders

executive_report_sample = executive_report.head(100)

operational_output_file = "Group2_Operational_Report.xlsx"
executive_output_file = "Group2_Executive_Report.xlsx"

operational_table_sample.to_excel(operational_output_file, index=False, engine='openpyxl')
executive_report_sample.to_excel(executive_output_file, index=False, engine='openpyxl')

# Display output
print("Operational Table (First 10 Rows):")
print(operational_table_sample)

print("\nExecutive Report Table (First 10 Rows):")
print(executive_report_sample)