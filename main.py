# main.py
# automation in insurance field
# purpose: 1. the script can sort all the required (original) plancode throughout the workbook in Excel 
# 2. duplicate the required (original) plancode to the bottom of the data 
# 3. rename the duplicated plancode to the new plancode
import xlwings as xw
import pandas as pd
import os 

def duplicate_and_modify_rows(file_path, product_pairs):
    # Open the Excel application
    app = xw.App(visible=False)
    
    try:
        wb = app.books.open(file_path)
        
        changes_made = False
        
        # Iterate through all sheets
        for sheet in wb.sheets:
            print(f"Processing sheet: {sheet.name}")
            
            # Find the column with the product names (assuming it's the second column)
            product_column = 2
            
            # Get the used range
            used_range = sheet.used_range
            
            # Get the values of the used range
            values = used_range.value
            
            # List to store rows to be added at the end
            rows_to_add = []
            
            # Iterate through rows
            for row in values:
                if row:
                    current_product = row[product_column - 1]
                    for target_product, new_product in product_pairs:
                        if current_product == target_product:
                            print(f"Found target product '{target_product}' in sheet '{sheet.name}'")
                            # Create a new row with modified data
                            new_row = list(row)
                            new_row[product_column - 1] = new_product
                            
                            # Add the new row to the list of rows to be added
                            rows_to_add.append(new_row)
                            break  # Move to the next row after finding a match
            
            # Add all new rows at the end of the sheet
            if rows_to_add:
                last_row = used_range.last_cell.row
                sheet.range(f'A{last_row + 1}').value = rows_to_add
                print(f"Added {len(rows_to_add)} new rows to sheet '{sheet.name}'")
                changes_made = True
            else:
                print(f"No matching rows found in sheet '{sheet.name}'")
        
        if changes_made:
            wb.save()
            print(f"Process completed. Changes saved to {file_path}")
        else:
            print(f"No changes were made to the file. No target products found.")
    
    finally:
        # Close the workbook and quit the Excel application
        wb.close()
        app.quit()

# Input file
file_path = r'C:\2024-09 (v2)\7. Table Working\TABLE_conversion\testing\products.xlsm'

# product_list = pd.read_csv(r'C:\2024-09 (v2)\7. Table Working\TABLE_conversion\testing\product_list.csv') 

# pair = []
# for i in range(product_list.shape[0]):
#     tmp_pair = product_list.loc[i,:]
#     print(tmp_pair)
#     pair.append(tuple(tmp_pair))
# print(pair)

product_pairs = [

('AAA1A','AAB1A'),

}

duplicate_and_modify_rows(file_path, product_pairs)
