import pandas as pd
import openpyxl
import win32com.client

# Reading sample superstore dataset
data_file_path = 'E:\\zPankaj\\Sample Superstore Invoices project\\Source\\Sample - Superstore.xlsx'
df = pd.read_excel(data_file_path, sheet_name='Orders')

print(df)

df.drop_duplicates(inplace=True)


#Keeping only these columns
df = df[['Order ID', 'Order Date', 'Ship Date', 
       'Customer ID', 'Customer Name', 'Segment', 'Country', 'City', 'State',
       'Postal Code','Product ID',
       'Product Name', 'Sales', 'Quantity', 'Discount', 'Profit']]

"""
Group Data 
Order ID wise
Then Customer ID wise

"""
print(df.columns)