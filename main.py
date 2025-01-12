# main.py
# automation in insurance field
# purpose: 1. the script can sort all the required (original) plancode throughout the workbook in Excel 
# 2. duplicate the required (original) plancode to the bottom of the data 
# 3. rename the duplicated plancode to the new plancode
# problem: a good function should only contain 30 lines, so better to separate the function into different portions

# main.py
import xlwings as xw

def open_workbook(file_path: str = ""):
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    return app, wb

# print(open_workbook(r"C:\2024-09 (v2)\7. Table Working\TABLE_conversion\testing"))

def process_sheet(sheet, product_pairs):
    print(f"Processing sheet: {sheet.name}")

    used_range = sheet.used_range
    values = used_range.value

    if not values:
        return []
    
    # Convert product pairs to dict for faster lookup
    product_map = dict(product_pairs)
    rows_to_add = []

    for row in values:
        if not row:
            continue
            
        needs_duplication = False
        new_row = None

        # Check each cell in the row for product names
        for col_idx, cell_value in enumerate(row):
            if isinstance(cell_value, str) and cell_value in product_map:
                if not needs_duplication:
                    needs_duplication = True
                    new_row = list(row)
                new_row[col_idx] = product_map[cell_value]

            # Handle cases where product code might be enclosed in double quotes
            elif isinstance(cell_value, str):
                cleaned_value =cell_value.strip('"\\')
                if cleaned_value in product_map:
                    if not needs_duplication:
                        needs_duplication = True
                        new_row = list(row)
                    if cell_value.startswith('"'):
                        new_row[col_idx] = f'"{product_map[cleaned_value]}"'
                    elif cell_value.startswith('\\'):
                        new_row[col_idx] = '\\'
                    else:
                        new_row[col_idx] = product_map[cleaned_value]
        if needs_duplication:
            rows_to_add.append(new_row)

    return rows_to_add

def add_rows_to_sheet(sheet, rows_to_add):
    if rows_to_add:
        last_row = sheet.used_range.last_cell.row
        sheet.range(f'A{last_row + 1}').value = rows_to_add
        print(f"Added {len(rows_to_add)} new rows to sheet '{sheet.name}'")
        return True
    else:
        print(f"No matching rows found in sheet '{sheet.name}'")
        return False

def duplicate_and_modify_rows(file_path, product_pairs):
    app, wb = open_workbook(file_path)

    changes_made = False

    try: 
        for sheet in wb.sheets:
            rows_to_add = process_sheet(sheet, product_pairs)
            if add_rows_to_sheet(sheet, rows_to_add):
                changes_made = True

        if changes_made:
            wb.save()
            print(f"Process completed. Changes saved to {file_path}")
        else:
            print(f"No changes were made to the file. No target products found.")
        
    finally:
        wb.close()
        app.quit()

file_path = r'C:\2024-09 (v2)\7. Table Working\TABLE_conversion\testing\Table conversion_v8.3 1 _products_trial.xlsm'

product_pairs = [
('CGG01A','CGK01A'),    
('CGG01M','CGK01M'),
('CGG05A','CGK05A'),
('CGG05M','CGK05M'),
('CGG10A','CGK10A'),
('CGG10M','CGK10M'),
('CGH01A','CGL01A'),
('CGH01H','CGL01H'),
('CGH01M','CGL01M'),
('CGH05A','CGL05A'),
('CGH05H','CGL05H'),
('CGH05M','CGL05M'),
('CGH10A','CGL10A'),
('CGH10H','CGL10H'),
('CGH10M','CGL10M'),
('CGI01A','CGM01A'),
('CGI01H','CGM01H'),
('CGI01M','CGM01M'),
('CGI05A','CGM05A'),
('CGI05H','CGM05H'),
('CGI05M','CGM05M'),
('CGI10A','CGM10A'),
('CGI10H','CGM10H'),
('CGI10M','CGM10M'),
('CGJ01A','CGN01A'),
('CGJ01H','CGN01H'),
('CGJ01M','CGN01M'),
('CGJ05A','CGN05A'),
('CGJ05H','CGN05H'),
('CGJ05M','CGN05M'),
('CGJ10A','CGN10A'),
('CGJ10H','CGN10H'),
('CGJ10M','CGN10M'),
]

duplicate_and_modify_rows(file_path, product_pairs)

'''
4
!2 PROD_NAME 1 2 3 12
* CGG01A     10 0 0 CGG01A
...
'''
