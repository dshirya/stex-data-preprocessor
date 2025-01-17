import pandas as pd
import re
import os

element_file = 'periodic_table.xlsx'  # File containing valid elements
properties_file = 'element_Mendeleev_numbers.csv'  # File containing element properties
input_file = '1929_new.xlsx'  # Input Excel file with formulas
max_elements = 3  # Maximum number of elements allowed in a formula
element_sheet = None  

class FormulaHandler:
    def __init__(self, element_file, properties_file, max_elements=3, element_sheet=None):
        self.valid_elements = self.load_valid_elements(element_file, element_sheet)
        self.max_elements = max_elements

    def load_valid_elements(self, element_file, sheet_name):
        if element_file.endswith('.xlsx') or element_file.endswith('.xls'):
            df_elements = pd.read_excel(element_file, sheet_name=sheet_name, header=None)
        else:
            df_elements = pd.read_csv(element_file, header=None)
        return set(df_elements[0])

    def parse_formula(self, formula):
        pattern = r'([A-Z][a-z]*)(\d*\.?\d*)'
        matches = re.findall(pattern, formula)
        elements = {}
        for element, count in matches:
            elements[element] = float(count) if count else 1
        return elements

    def filter_formulas(self, df):
        def filter_row(formula):
            elements = self.parse_formula(formula)
            if len(elements) > self.max_elements:
                return False
            for element in elements:
                if element not in self.valid_elements:
                    return False
            return True

        return df[df['Formula'].apply(filter_row)]

    def process_all_sheets(self, file_path, output_file_path=None):
        sheets = pd.read_excel(file_path, sheet_name=None)
        if output_file_path is None:
            base_name, ext = os.path.splitext(file_path)
            output_file_path = f"{base_name}_filtered{ext}"

        processed_sheets = {}

        for sheet_name, df in sheets.items():
            df_filtered = self.filter_formulas(df)
            processed_sheets[sheet_name] = df_filtered

        with pd.ExcelWriter(output_file_path) as writer:
            for sheet_name, df_processed in processed_sheets.items():
                df_processed.to_excel(writer, sheet_name=sheet_name, index=False)

if __name__ == "__main__":
    if element_file.endswith('.xlsx') or element_file.endswith('.xls'):
        sheet_names = pd.ExcelFile(element_file).sheet_names
        if len(sheet_names) > 1:
            print("Available sheets in the periodic table file:")
            for i, sheet_name in enumerate(sheet_names, start=1):
                print(f"{i}: {sheet_name}")
            choice = int(input(f"Select a sheet number (1-{len(sheet_names)}): "))
            element_sheet = sheet_names[choice - 1]

    # Initialize and process
    fh = FormulaHandler(element_file, properties_file, max_elements, element_sheet)
    fh.process_all_sheets(input_file)
    print(f"Filtering complete. Output saved with '_filtered' suffix.")