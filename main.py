import pandas as pd 

df = pd.read_excel("data/Online Retail.xlsx")
df["Total"]= df["Quantity"] * df["UnitPrice"]

def generate_report(df):
    #stock description according to stock ID
    stock_info = df[["StockCode","Description"]].drop_duplicates()
    stock_info.to_excel("stock_info.xlsx",index=False)
    # Convert necessary columns to numeric
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Total"] = pd.to_numeric(df["Total"], errors="coerce")

    # Cancelled data
    cancelled = df[df["InvoiceNo"].astype(str).str.startswith("C")]

    print("📉 Most Cancelled Items (by quantity):")
    print(
        cancelled.groupby("Description")["Quantity"]
        .sum()
        .sort_values()
        .head(10)
    )
    print("\n👥 Most Cancelled Customers (by quantity):")
    print(
        cancelled.groupby("CustomerID")["Quantity"]
        .sum()
        .sort_values()
        .head(10)
    )
    print("\n💸 Most Cancelled Customers (by value):")
    print(
        cancelled.groupby("CustomerID")["Total"]
        .sum()
        .sort_values()
        .head(10)
    )

    print("\n🔥 Top Selling Items (by quantity):")
    print(
        df.groupby("Description")["Quantity"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
    )

    print("\n💰 Top Earning Customers (by total purchase):")
    print(
        df.groupby("CustomerID")["Total"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
    )

generate_report(df)
