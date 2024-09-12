import pandas as pd
import openpyxl
import win32com.client
import warnings

warnings.filterwarnings("ignore")

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


#Accessing invoice template and pasting data from df1 to invoice template.
inv_temp_wb = openpyxl.load_workbook("Source/invoice template.xlsx")
ws = inv_temp_wb["Invoice"]

# inv_temp_wb.save('C:\\Users\\panka\\OneDrive\\Desktop\\Abcd.xlsx')

# Iterating column and row wise in invoice template
# function to find cell next to Or below required fiels in invoice.
def cell_for_entering_value_finder(inv_field):
    for i in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for j in i:
            # print("Sheet",j.value)
            if j.value == inv_field:
                if inv_field in ["Product ID","Product Name","Unit Price","Quantity","Price"]:
                # Cell_num will be the cell in which value of inv field will be enter for a customer.
                    cell_num = ws.cell(row=j.row+1, column=j.column)
                else:
                    cell_num = ws.cell(row=j.row, column=j.column+1)

    return cell_num

inv_field = "Customer Name"
a = cell_for_entering_value_finder(inv_field)
print(a)

# fields for invoice from df1

print(df.columns)

#Creating list of distinct  Order ID
order_id_list = df['Order ID'].unique().tolist()
# print(len(order_id_list))

#Note: There are 5009 distinct Order ID. It would create 5009 invoices in Invoices folder and would take a lot of space. 
#For demonstration, I am adding a Region filter
df = df[df['Region'] == 'West']



def fill_invoice(df1):

    try:
        custID = (df1['Customer ID'].unique().tolist())[0]
        # print("Success")
        custname = (df1['Customer Name'].unique().tolist())[0]
        segm = (df1['Segment'].unique().tolist())[0]
        df1['address'] = df1['City'] + df1['State'] + df1['Country']
        address = (df1['address'].unique().tolist())[0]
        oid = (df1['Order ID'].unique().tolist())[0]
        odt = (df1['Order Date'].unique().tolist())[0]
        sdt = (df1['Ship Date'].unique().tolist())[0]
        ord_details = (df1['Order Details'].unique().tolist())[0]
        prod_ID = (df1['Product ID'].unique().tolist())[0]
        prod_name = (df1['Product Name'].unique().tolist())[0]
        unit_price = (df1['Unit Price'].unique().tolist())[0]
        qty = (df1['Quantity'].unique().tolist())[0]
        print("Hello", custID, custname, segm, address, oid, odt, sdt, ord_details, prod_ID, prod_name, unit_price, qty)

    except Exception as e:
        print(e)

for ord in order_id_list:
    df1 = df[df['Order ID'] == ord]
    # print(df1)
    print('DF1', df1)
    if not df1.empty or df1.notnull:
        fill_invoice(df1)