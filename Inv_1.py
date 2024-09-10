import pandas as pd
import openpyxl
import win32com.client

# Reading sample superstore dataset
data_file_path = 'E:\\zPankaj\\Sample Superstore Invoices project\\Source\\Sample - Superstore.xlsx'
df = pd.read_excel(data_file_path, sheet_name='Orders')

print(df)

df.drop_duplicates(inplace=True)


#Keeping only these columns
df = df[['Order ID', 'Order Date', 'Ship Date', 'Region',
       'Customer ID', 'Customer Name', 'Segment', 'Country', 'City', 'State',
       'Postal Code','Product ID',
       'Product Name', 'Sales', 'Quantity', 'Discount', 'Profit']]

# Group Data 
# Order ID wise

print(df.columns)


# df = df['Region']
df = df[df['Region'] == 'West']

# order_id_list0 = df['Order ID'].tolist()
# print(len(order_id_list0))

#Creating list of distinct  Order ID
order_id_list = df['Order ID'].unique().tolist()
print(len(order_id_list))

#Note: There are 5009 distinct Order ID. It would create 5009 invoices in Invoices folder and would take a lot of space. 
#For demonstration, I am adding a Region filter



for ord in order_id_list:
    df1 = df[df['Order ID'] == ord]


#     print('DF1', df1)



