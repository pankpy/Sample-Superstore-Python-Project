
import openpyxl
from openpyxl import cell

inv_temp_wb = openpyxl.load_workbook("Source/invoice template.xlsx")
ws = inv_temp_wb["Invoice"]

# inv_temp_wb.save('C:\\Users\\panka\\OneDrive\\Desktop\\Abcd.xlsx')

# Iterating column and row wise in invoice template to find cell next or below fields, so that values can be enter in them.

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