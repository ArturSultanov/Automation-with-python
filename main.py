# Artur Sultanov June 2023
import openpyxl

inv_file = openpyxl.load_workbook("inventory.xlsx")
product_list = inv_file["Sheet1"]

# Task 1 - how many products per supplier

products_num_per_supplier = {}

for product_row in range(2, product_list.max_row + 1):
    supplier_name = product_list.cell(product_row, 4).value

    if supplier_name in products_num_per_supplier:
        products_num_per_supplier[supplier_name] += 1
    else:
        products_num_per_supplier[supplier_name] = 1

print(products_num_per_supplier)
