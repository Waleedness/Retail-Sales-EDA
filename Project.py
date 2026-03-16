import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# =========================================================================================================================================================================================================================
# 1. DATA LOADING
# =========================================================================================================================================================================================================================

dataframe = pd.read_excel(r"C:\Users\atifa\Desktop\Projects\Data set for Project 1.xlsx")

# =========================================================================================================================================================================================================================
# 2. REVENUE METRICS
# =========================================================================================================================================================================================================================

total_gross_revenue = dataframe["Gross Amount"].sum()
print(f"Total Gross Revenue: {total_gross_revenue}")

total_net_revenue = dataframe["Net Amount"].sum()
print(f"Total Net Revenue: {total_net_revenue}")

total_discount_amount = total_gross_revenue - total_net_revenue
print(f"Total Discounts Given: {total_discount_amount}")

# =========================================================================================================================================================================================================================
# 3. TRANSACTION METRICS
# =========================================================================================================================================================================================================================

total_transactions = dataframe["TID"].count()
print(f"Total Transactions: {total_transactions}")

# =========================================================================================================================================================================================================================
# 4. CUSTOMER METRICS
# =========================================================================================================================================================================================================================

total_unique_customers = dataframe["CID"].nunique()
print(f"Total Unique Customers: {total_unique_customers}")

average_revenue_per_customer = total_net_revenue / total_unique_customers
print(f"Average Revenue Per Customer: {average_revenue_per_customer}")

# =========================================================================================================================================================================================================================
# 5. PRODUCT & LOCATION OVERVIEW
# =========================================================================================================================================================================================================================

total_product_categories = dataframe["Product Category"].nunique()
print(f"Total Product Categories: {total_product_categories}")

total_store_locations = dataframe["Location"].nunique()
print(f"Total Store Locations: {total_store_locations}")

# =========================================================================================================================================================================================================================
# 6. CENTRAL TENDENCY METRICS
# =========================================================================================================================================================================================================================

median_net_amount = dataframe["Net Amount"].median()
print(f"Median Net Amount: {median_net_amount}")

max_value = dataframe["Net Amount"].max()
print(f"Largest Transaction: {max_value}")

min_value = dataframe["Net Amount"].min()
print(f"Smallest Transaction: {min_value}")

negative_transactions = (dataframe["Net Amount"] < 0).sum()
negative_percentage = (negative_transactions / len(dataframe)) * 100

print(f"Negative Transactions: {negative_transactions}")
print(f"Percentage of Negative Transactions: {negative_percentage:.2f}%")

# =========================================================================================================================================================================================================================
# 7. Revenue By Time
# =========================================================================================================================================================================================================================

df = dataframe.copy()

df["Purchase Date"] = pd.to_datetime(df["Purchase Date"], dayfirst=True)

df["Year"] = df["Purchase Date"].dt.year
df["YearMonth"] = df["Purchase Date"].dt.to_period("M").dt.to_timestamp()

net_amount_by_year = df.groupby("Year")["Net Amount"].sum().reset_index()
print(net_amount_by_year)

net_amount_by_month = df.groupby("YearMonth")["Net Amount"].sum().reset_index()
print(net_amount_by_month)

plt.ticklabel_format(style='plain', axis='y')
plt.plot(net_amount_by_year["Year"], net_amount_by_year["Net Amount"]/1000, marker="o")
plt.title("Net Amount By Year")
plt.xlabel("Year")
plt.ylabel("Net Amount (In Thousands)")
plt.grid(axis="y", linestyle="--", alpha=0.3)
plt.tight_layout()
plt.show()

plt.figure(figsize=(12,6))
plt.ticklabel_format(style='plain', axis='y')
plt.plot(net_amount_by_month["YearMonth"], net_amount_by_month["Net Amount"]/1000, marker="o")
plt.title("Net Amount By Month")
plt.xlabel("Date")
plt.ylabel("Net Amount (In Thousands)")
plt.grid(axis="y", linestyle="--", alpha=0.3)
plt.tight_layout()
plt.show()

# =========================================================================================================================================================================================================================
# 8. Revenue By Location
# =========================================================================================================================================================================================================================

net_amount_by_location = df.groupby("Location")["Net Amount"].sum().reset_index()

sorted_net_amount_by_location = net_amount_by_location.sort_values(by="Net Amount", ascending=False)
print(sorted_net_amount_by_location)

plt.figure(figsize=(17,8))
plt.ticklabel_format(style='plain', axis='y')
plt.bar(sorted_net_amount_by_location["Location"][::-1], sorted_net_amount_by_location["Net Amount"][::-1]/1000)
plt.title("Net Amount By Location")
plt.xlabel("Location")
plt.ylabel("Net Amount (In Thousands)")
plt.xticks(rotation=45)
plt.grid(axis="y", linestyle="--", alpha=0.4)
plt.tight_layout()
plt.show()

# =========================================================================================================================================================================================================================
# 9. Revenue By Product Category
# =========================================================================================================================================================================================================================

net_amount_by_product_category = df.groupby("Product Category")["Net Amount"].sum().reset_index()

sorted_net_amount_by_product_category = net_amount_by_product_category.sort_values(by="Net Amount", ascending=False)
print(sorted_net_amount_by_product_category)

plt.figure(figsize=(17,8))
plt.ticklabel_format(style='plain', axis='y')
plt.bar(sorted_net_amount_by_product_category["Product Category"][::-1], sorted_net_amount_by_product_category["Net Amount"][::-1]/1000)
plt.title("Net Amount By Product Category")
plt.xlabel("Product Category")
plt.ylabel("Net Amount (In Thousands)")
plt.xticks(rotation=45)
plt.grid(axis="y", linestyle="--", alpha=0.4)
plt.tight_layout()
plt.show()

# =========================================================================================================================================================================================================================
# 10. Revenue Distribution
# =========================================================================================================================================================================================================================

plt.figure(figsize=(17,10))
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(17,10), gridspec_kw={'height_ratios':[3,1]})

ax1.hist(dataframe["Net Amount"], bins=40, rwidth=0.9, edgecolor="black")
ax1.set_title("Distribution of Net Transaction Amounts")
ax1.set_xlabel("Net Amount")
ax1.set_ylabel("Number of Transactions")
ax1.grid(axis="y", linestyle="--", alpha=0.4)

ax2.boxplot(
    dataframe["Net Amount"],
    vert=False,
    patch_artist=True,
    boxprops=dict(facecolor="yellow", color="black"),
    medianprops=dict(color="black"),
    flierprops=dict(marker='x', color='green', markersize=8)
)

ax2.set_xlabel("Net Amount")
ax2.set_yticks([])
ax2.grid(axis="x", linestyle="--", alpha=0.4)

plt.tight_layout()
plt.show()

# =========================================================================================================================================================================================================================
# 11. Top Customers By Revenue
# =========================================================================================================================================================================================================================

revenue_by_customer = df.groupby("CID")["Net Amount"].sum().reset_index()

top_customers = revenue_by_customer.sort_values(by="Net Amount", ascending=False).head(10)

print("Top 10 Customers By Revenue")
print(top_customers)

plt.figure(figsize=(12,6))
plt.ticklabel_format(style='plain', axis='y')
plt.bar(top_customers["CID"].astype(str), top_customers["Net Amount"])
plt.title("Top 10 Customers By Revenue")
plt.xlabel("Customer ID")
plt.ylabel("Net Amount")
plt.grid(axis="y", linestyle="--", alpha=0.4)
plt.tight_layout()
plt.show()


# =========================================================================================================================================================================================================================
# 12. Average Transaction Value By Location
# =========================================================================================================================================================================================================================

avg_transaction_location = df.groupby("Location")["Net Amount"].mean().reset_index()

sorted_avg_transaction_location = avg_transaction_location.sort_values(by="Net Amount", ascending=False)

print(sorted_avg_transaction_location)

plt.figure(figsize=(17,8))
plt.ticklabel_format(style='plain', axis='y')
plt.bar(sorted_avg_transaction_location["Location"][::-1], sorted_avg_transaction_location["Net Amount"][::-1])
plt.title("Average Transaction Value By Location")
plt.xlabel("Location")
plt.ylabel("Average Transaction Value")
plt.xticks(rotation=45)
plt.grid(axis="y", linestyle="--", alpha=0.4)
plt.tight_layout()
plt.show()


# =========================================================================================================================================================================================================================
# 13. Transactions By Product Category
# =========================================================================================================================================================================================================================

transactions_by_category = df.groupby("Product Category")["TID"].count().reset_index()

sorted_transactions_by_category = transactions_by_category.sort_values(by="TID", ascending=False)

print(sorted_transactions_by_category)

plt.figure(figsize=(17,8))
plt.bar(sorted_transactions_by_category["Product Category"][::-1], sorted_transactions_by_category["TID"][::-1])
plt.title("Number of Transactions By Product Category")
plt.xlabel("Product Category")
plt.ylabel("Number of Transactions")
plt.xticks(rotation=45)
plt.grid(axis="y", linestyle="--", alpha=0.4)
plt.tight_layout()
plt.show()


# =========================================================================================================================================================================================================================
# 14. Monthly Transaction Count
# =========================================================================================================================================================================================================================

transactions_by_month = df.groupby("YearMonth")["TID"].count().reset_index()

print(transactions_by_month)

plt.figure(figsize=(12,6))
plt.plot(transactions_by_month["YearMonth"], transactions_by_month["TID"], marker="o")
plt.title("Transactions By Month")
plt.xlabel("Date")
plt.ylabel("Number of Transactions")
plt.grid(axis="y", linestyle="--", alpha=0.4)
plt.tight_layout()
plt.show()


# =========================================================================================================================================================================================================================
# 15. Discount Impact Analysis
# =========================================================================================================================================================================================================================

df["Discount"] = df["Gross Amount"] - df["Net Amount"]

discount_by_category = df.groupby("Product Category")["Discount"].sum().reset_index()

sorted_discount_by_category = discount_by_category.sort_values(by="Discount", ascending=False)

print(sorted_discount_by_category)

plt.figure(figsize=(17,8))
plt.bar(sorted_discount_by_category["Product Category"][::-1], sorted_discount_by_category["Discount"][::-1]/1000)
plt.title("Total Discounts By Product Category")
plt.xlabel("Product Category")
plt.ylabel("Discount Amount (In Thousands)")
plt.xticks(rotation=45)
plt.grid(axis="y", linestyle="--", alpha=0.4)
plt.tight_layout()
plt.show()


# =========================================================================================================================================================================================================================
# 16. Customer Purchase Frequency
# =========================================================================================================================================================================================================================

customer_frequency = df.groupby("CID")["TID"].count().reset_index()

customer_frequency.columns = ["CID", "Purchase Count"]

print(customer_frequency.describe())

plt.figure(figsize=(12,6))
plt.hist(customer_frequency["Purchase Count"], bins=30, edgecolor="black")
plt.title("Customer Purchase Frequency Distribution")
plt.xlabel("Number of Purchases")
plt.ylabel("Number of Customers")
plt.grid(axis="y", linestyle="--", alpha=0.4)
plt.tight_layout()
plt.show()


# =========================================================================================================================================================================================================================
# 17. Top Performing Months
# =========================================================================================================================================================================================================================

top_months = net_amount_by_month.sort_values(by="Net Amount", ascending=False).head(10)

print("Top 10 Revenue Months")
print(top_months)


# =========================================================================================================================================================================================================================
# 18. Final Summary Metrics
# =========================================================================================================================================================================================================================

average_transaction_value = total_net_revenue / total_transactions

print("==========================================")
print("FINAL BUSINESS SUMMARY")
print("==========================================")

print(f"Total Revenue: {total_net_revenue}")
print(f"Total Transactions: {total_transactions}")
print(f"Total Customers: {total_unique_customers}")
print(f"Average Transaction Value: {average_transaction_value}")
print(f"Average Revenue Per Customer: {average_revenue_per_customer}")
print(f"Total Discounts Given: {total_discount_amount}")