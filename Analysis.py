# -*- coding: utf-8 -*-
"""
Created on Sun Dec  8 08:14:45 2024

@author: Maedot
"""

import pandas as pd
import matplotlib.pyplot as plt

file_path = 'online_retail.xlsx'

# Reading the Excel file and specifying the second sheet
data = pd.read_excel(file_path, sheet_name=1)

# Print missing values in each column
print("Missing values in each column:")
print(data.isnull().sum())  # To figure out missing values

# Drop rows with missing CustomerID and Price
data = data.dropna(subset=['CustomerID', 'Price'])

# Filter out zeros and negative values in Price and Quantity
data = data[(data['Quantity'] > 0) & (data['Price'] > 0)]

# Print the head of the cleaned data
print("\nHead of cleaned data:")
print(data.head())

# Print data info: to check data type and missing values
print("\nData Info:")
print(data.info())

# Print descriptive statistics for the data
print("\nDescriptive statistics for the data:")
print(data.describe())

# Print statistics for the 'Price' column
print("\nStatistics for 'Price' column:")
print(data[['Price']].describe())

# Calculate total customers
unique_customers = data['CustomerID'].nunique()

# Calculate number of products sold
unique_products = data['StockCode'].nunique()

# Calculate total transactions
total_transactions = data['Invoice'].nunique()

# Print the results for customers, products, and transactions
print(f"\nTotal Customers: There were {unique_customers} unique customers.")
print(f"Number of Products Sold: The store sold {unique_products} different products.")
print(f"Total Transactions: The dataset includes {total_transactions} transactions.")

# Calculate total spending for each transaction
data['TotalSpend'] = data['Quantity'] * data['Price']

# Group transactions by CustomerID and sum TotalSpend for each customer
customer_spending = data.groupby('CustomerID')['TotalSpend'].sum()

# Calculate the average amount each customer spends per transaction
avg_spending_per_transaction = data.groupby('CustomerID')['TotalSpend'].mean()

# Count the number of transactions for each customer
transactions_per_customer = data.groupby('CustomerID').size()

# Create a summary table for spending behavior for each customer
spending_analysis = pd.DataFrame({
    'TotalSpend': customer_spending,
    'AvgSpendPerTransaction': avg_spending_per_transaction,
    'TransactionCount': transactions_per_customer
}).reset_index()

# Print customer spending analysis summary
print("\nCustomer Spending Analysis:")
print(spending_analysis.head())

# Convert InvoiceDate to a datetime object
data['InvoiceDate'] = pd.to_datetime(data['InvoiceDate'])

# Calculate how many unique days each customer made purchases
purchase_frequency = data.groupby('CustomerID')['InvoiceDate'].nunique()

# Sort data by CustomerID and InvoiceDate so that earlier transactions come first
data = data.sort_values(['CustomerID', 'InvoiceDate'])

# Find the previous purchase date for each transaction and calculate the number of days between purchases
data['PreviousPurchaseDate'] = data.groupby('CustomerID')['InvoiceDate'].shift()
data['DaysBetweenPurchases'] = (data['InvoiceDate'] - data['PreviousPurchaseDate']).dt.days

# Calculate the average time between purchases for each customer
avg_time_between_purchases = data.groupby('CustomerID')['DaysBetweenPurchases'].mean()

# Summarize purchase frequency and average time between purchases for each customer
frequency_analysis = pd.DataFrame({
    'UniquePurchaseDates': purchase_frequency,
    'AvgDaysBetweenPurchases': avg_time_between_purchases
}).reset_index()

# Print customer frequency analysis summary
print("\nCustomer Frequency Analysis:")
print(frequency_analysis.head())

# Combine spending analysis and frequency analysis into one table, giving a full picture of customer behavior
customer_behavior = spending_analysis.merge(frequency_analysis, on='CustomerID')

# Print combined customer behavior summary
print("\nCustomer Behavior Summary:")
print(customer_behavior.head())


# Plot Average Spend per Transaction
plt.figure(figsize=(10, 6))
top_avg_spenders = spending_analysis.sort_values('AvgSpendPerTransaction', ascending=False).head(20)
plt.bar(top_avg_spenders['CustomerID'].astype(str), top_avg_spenders['AvgSpendPerTransaction'], color='green')
plt.title('Top 20 Customers by Average Spend per Transaction')
plt.xlabel('Customer ID')
plt.ylabel('Average Spend per Transaction')
plt.xticks(rotation=90)
plt.tight_layout()
plt.show()

# Plot Number of Transactions per Customer
plt.figure(figsize=(10, 6))
top_transaction_counts = spending_analysis.sort_values('TransactionCount', ascending=False).head(20)
plt.bar(top_transaction_counts['CustomerID'].astype(str), top_transaction_counts['TransactionCount'], color='orange')
plt.title('Top 20 Customers by Number of Transactions')
plt.xlabel('Customer ID')
plt.ylabel('Number of Transactions')
plt.xticks(rotation=90)
plt.tight_layout()
plt.show()

# Plot Unique Purchase Dates per Customer
plt.figure(figsize=(10, 6))
top_unique_purchase_dates = frequency_analysis.sort_values('UniquePurchaseDates', ascending=False).head(20)
plt.bar(top_unique_purchase_dates['CustomerID'].astype(str), top_unique_purchase_dates['UniquePurchaseDates'], color='purple')
plt.title('Top 20 Customers by Number of Unique Purchase Dates')
plt.xlabel('Customer ID')
plt.ylabel('Number of Unique Purchase Dates')
plt.xticks(rotation=90)
plt.tight_layout()
plt.show()

# Inspect calculated DaysBetweenPurchases
print(data[['CustomerID', 'InvoiceDate', 'PreviousPurchaseDate', 'DaysBetweenPurchases']].head(20))
