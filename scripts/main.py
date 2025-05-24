import pandas as pd 
import helpers
from xlsxwriter import Workbook

df = pd.read_excel("data/Online Retail.xlsx")
df["Total"]= df["Quantity"] * df["UnitPrice"]


with pd.ExcelWriter("reports/ecommerce_report.xlsx", engine="xlsxwriter") as writer: 
    pd.DataFrame({"Metric": ["Total Revenue"], "Value": [df["Total"].sum()]}).to_excel(writer, sheet_name="Revenue", index=False)

    # 💰 Top earning customers
    helpers.get_top_customers(df,10).to_frame(name="Total").to_excel(writer, sheet_name="Top_Customers")

    # ❌ Most cancelled items
    helpers.get_most_cancelled_items(df,10).to_frame(name="Cancelled Quantity").to_excel(writer, sheet_name="Cancelled_Items")

    # 📈 Top selling items by revenue
    helpers.get_top_selling_items(df,10).to_frame(name="Total Revenue").to_excel(writer, sheet_name="Top_Selling_Items")

    # 🚫 Most cancelled customers (quantity)
    helpers.get_most_cancelled_customers_qty(df,10).to_frame(name="Cancelled Qty").to_excel(writer, sheet_name="Cancelled_Customers_Qty")

    # 💸 Most cancelled customers (value)
    helpers.get_most_cancelled_customers_value(df,10).to_frame(name="Cancelled Value").to_excel(writer, sheet_name="Cancelled_Customers_Value")

   
    # 🔥 Most popular items (by quantity sold)
    helpers.get_most_popular_items(df,10).to_frame(name="Total Quantity Sold").to_excel(writer, sheet_name="Popular_Items")



