# main.py
# automation in insurance field
# purpose: 1. the script can sort all the required (original) plancode throughout the workbook in Excel 
# 2. duplicate the required (original) plancode to the bottom of the data 
# 3. rename the duplicated plancode to the new plancode
# problem: a good function should only contain 30 lines, so better to separate the function into different portions
# readable, reusable, refactorable
# author: seishun KW

from dataclasses import dataclass
from typing import List, Tuple, Dict, Optional
import xlwings as xw

@dataclass
class WorkbookSession:
    app: xw.App
    workbook: xw.Book

    @classmethod
    def create(open, file_path: str) -> 'WorkbookSession':
        app = xw.App(visible=False)
        workbook = app.books.open(file_path)
        return open(app, workbook)
    
    def close(self) -> None:
        self.workbook.close()
        self.app.quit()
# original code   
# def open_workbook(file_path: str = ""):
#     app = xw.App(visible=False)
#     wb = app.books.open(file_path)
#     return app, wb

@dataclass
class ProductMapping:
    mapping: Dict[str, str]

    @classmethod
    def from_pairs(cls, pairs: List[Tuple[str, str]]) -> 'ProductMapping':
        return cls(dict(pairs))
    
    def get_new_plancode(self, original_plancode: str) -> Optional[str]:
        cleaned_plancode = original_plancode.strip('"\\')
        return self.mapping.get(cleaned_plancode)
    
    def format_new_code(self, original_plancode: str, new_plancode: str) -> str:
        if original_plancode.startswith('"'):
            return f'"{new_plancode}"'
        if original_plancode.startswith('\\'):
            return '\\'
        return new_plancode
    
class SheetProcessor:
    def __init__(self, sheet: xw.Sheet, product_mapping: ProductMapping):
        self.sheet = sheet
        self.product_mapping = product_mapping

    def process(self) -> List[list]:
        print(f"Processing sheet: {self.sheet.name}")
        values = self.sheet.used_range.value

        if not values:
            return []
        
        return self.find_rows_to_duplicate(values)
    
    def find_rows_to_duplicate(self, values: List[List]) -> List[List]:
        rows_to_add = []

        for row in values:
            if not row:
                continue

            modified_row = self.process_row(row)
            if modified_row:
                rows_to_add.append(modified_row)

        return rows_to_add
    
    def process_row(self, row: List) -> Optional[List]:
        needs_duplication = False
        new_row = None

        for col_idx, cell_value in enumerate(row):
            if not isinstance(cell_value, str):
                continue
            
            new_plancode = self.product_mapping.get_new_plancode(cell_value)
            if new_plancode:
                if not needs_duplication:
                    needs_duplication = True
                    new_row = list(row)
                new_row[col_idx] = self.product_mapping.format_new_code(cell_value, new_plancode)

        return new_row if needs_duplication else None

class ExcelConverter:
    def __init__(self, file_path: str, product_pairs: List[Tuple[str, str]]):
        self.file_path = file_path
        self.product_mapping = ProductMapping.from_pairs(product_pairs)

    def convert(self) -> None:
        session = WorkbookSession.create(self.file_path)
        changes_made = False

        try:
            for sheet in session.workbook.sheets:
                processor = SheetProcessor(sheet, self.product_mapping)
                rows_to_add = processor.process()

                if self.add_rows_to_sheet(sheet, rows_to_add):
                    changes_made = True

            self.save_if_changed(session.workbook, changes_made)
        
        finally:
            session.close()
# def process_sheet(sheet, product_pairs):
#     print(f"Processing sheet: {sheet.name}")

#     used_range = sheet.used_range
#     values = used_range.value

#     if not values:
#         return []
    
#     # Convert product pairs to dict for faster lookup
#     product_map = dict(product_pairs)
#     rows_to_add = []

#     for row in values:
#         if not row:
#             continue
            
#         needs_duplication = False
#         new_row = None

#         # Check each cell in the row for product names
#         for col_idx, cell_value in enumerate(row):
#             if isinstance(cell_value, str) and cell_value in product_map:
#                 if not needs_duplication:
#                     needs_duplication = True
#                     new_row = list(row)
#                 new_row[col_idx] = product_map[cell_value]

#             # Handle cases where product code might be enclosed in double quotes
#             elif isinstance(cell_value, str):
#                 cleaned_value =cell_value.strip('"\\')
#                 if cleaned_value in product_map:
#                     if not needs_duplication:
#                         needs_duplication = True
#                         new_row = list(row)
#                     if cell_value.startswith('"'):
#                         new_row[col_idx] = f'"{product_map[cleaned_value]}"'
#                     elif cell_value.startswith('\\'):
#                         new_row[col_idx] = '\\'
#                     else:
#                         new_row[col_idx] = product_map[cleaned_value]
#         if needs_duplication:
#             rows_to_add.append(new_row)

#     return rows_to_add
    def add_rows_to_sheet(self, sheet: xw.Sheet, rows_to_add: List[List]) -> bool:
        if not rows_to_add:
            print(f"No matching rows found in sheet '{sheet.name}'")
            return False
            
        last_row = sheet.used_range.last_cell.row
        sheet.range(f'A{last_row + 1}').value = rows_to_add
        print(f"Added {len(rows_to_add)} new rows to sheet '{sheet.name}'")
        return True
    
    def save_if_changed(self, workbook: xw.Book, changes_made: bool) -> None:
        if changes_made:
            workbook.save()
            print(f"Process completed. Changes saved to {self.file_path}")
        else:
            print("No changes were made to the file. No target products found.")

# def add_rows_to_sheet(sheet, rows_to_add):
#     if rows_to_add:
#         last_row = sheet.used_range.last_cell.row
#         sheet.range(f'A{last_row + 1}').value = rows_to_add
#         print(f"Added {len(rows_to_add)} new rows to sheet '{sheet.name}'")
#         return True
#     else:
#         print(f"No matching rows found in sheet '{sheet.name}'")
#         return False

# def duplicate_and_modify_rows(file_path, product_pairs):
#     app, wb = open_workbook(file_path)

#     changes_made = False

#     try: 
#         for sheet in wb.sheets:
#             rows_to_add = process_sheet(sheet, product_pairs)
#             if add_rows_to_sheet(sheet, rows_to_add):
#                 changes_made = True

#         if changes_made:
#             wb.save()
#             print(f"Process completed. Changes saved to {file_path}")
#         else:
#             print(f"No changes were made to the file. No target products found.")
        
#     finally:
#         wb.close()
#         app.quit()

def main():
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

    converter = ExcelConverter(file_path, product_pairs)
    converter.convert()

if __name__== "__main__":
    main()

''' ^^^Example^^^
4
!2 PROD_NAME 1 2 3 12
* CGG01A     10 0 0 CGG01A
...
'''
