import pandas as pd

def revenue(df):
    return df["total"].sum()

# 1. Top Earning Customers
def get_top_customers(df, n):
    return df.groupby("CustomerID")["Total"].sum().sort_values(ascending=False).head(n)

# 2. Most Cancelled Items
def get_most_cancelled_items(df, n):
    cancelled = df[df["InvoiceNo"].astype(str).str.startswith("C")]
    return cancelled.groupby("Description")["Quantity"].sum().sort_values().head(n)

# 3. Top Selling Items (by revenue)
def get_top_selling_items(df, n):
    return df.groupby("Description")["Total"].sum().sort_values(ascending=False).head(n)

# 4a. Most Cancelled Customers (by quantity)
def get_most_cancelled_customers_qty(df, n):
    cancelled = df[df["InvoiceNo"].astype(str).str.startswith("C")]
    return cancelled.groupby("CustomerID")["Quantity"].sum().sort_values().head(n)

# 4b. Most Cancelled Customers (by value)
def get_most_cancelled_customers_value(df, n):
    cancelled = df[df["InvoiceNo"].astype(str).str.startswith("C")]
    return cancelled.groupby("CustomerID")["Total"].sum().sort_values().head(n)

# 5. Top Items Purchased by a Specific Customer
def get_top_items_by_customer(df, customer_id, n):
    customer_df = df[df["CustomerID"] == customer_id]
    return (
        customer_df.groupby("Description")[["Quantity", "Total"]]
        .sum()
        .sort_values("Total", ascending=False)
        .head(n)
    )

# 6. Most Popular Products (by total quantity sold)
def get_most_popular_items(df, n):
    return df.groupby("Description")["Quantity"].sum().sort_values(ascending=False).head(n)
