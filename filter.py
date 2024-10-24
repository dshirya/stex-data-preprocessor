import pandas as pd
import re
import os

file_path = 'filtered_sorted_PuNi3_Cu3Au.xlsx'  # Path to the input Excel file


class FormulaFilter:
    def __init__(self, element_file, max_elements=3, element_sheet=None):
        # Load the valid elements from the specified sheet of the element file
        self.valid_elements = self.load_valid_elements(element_file, element_sheet)
        self.max_elements = max_elements

    def load_valid_elements(self, element_file, sheet_name):
        if element_file.endswith('.xlsx') or element_file.endswith('.xls'):
            df_elements = pd.read_excel(element_file, sheet_name=sheet_name, header=None)
        else:
            df_elements = pd.read_csv(element_file, header=None)

        return set(df_elements[0])

    def parse_formula(self, formula):
        # Regular expression to parse elements and their counts
        pattern = r'([A-Z][a-z]*)(\d*\.?\d*)'
        matches = re.findall(pattern, formula)
        elements = {}
        for element, count in matches:
            if count:
                elements[element] = float(count)
            else:
                elements[element] = 1
        return elements

    def filter_formulas(self, df):
        # Filter to keep rows with valid formulas (<= max_elements elements and all elements in the valid list)
        def filter_row(formula):
            elements = self.parse_formula(formula)
            
            '''Change here if you want to filter exact amount of element in the formula'''

            if len(elements) > self.max_elements:
                return False
            # Check if all elements are in the valid elements set
            for element in elements:
                if element not in self.valid_elements:
                    return False
            return True

        df_filtered = df[df['Formula'].apply(filter_row)]
        return df_filtered

    def process_all_sheets(self, file_path, output_file_path=None):
        sheets = pd.read_excel(file_path, sheet_name=None)

        # Generate the output filename if not provided, with "_filtered" suffix
        if output_file_path is None:
            base_name, ext = os.path.splitext(file_path)
            output_file_path = f"{base_name}_filtered{ext}"

        filtered_sheets = {}

        # Process each sheet
        for sheet_name, df in sheets.items():
            filtered_sheets[sheet_name] = self.filter_formulas(df)

        # Write the filtered sheets to a new Excel file
        with pd.ExcelWriter(output_file_path) as writer:
            for sheet_name, df_filtered in filtered_sheets.items():
                df_filtered.to_excel(writer, sheet_name=sheet_name, index=False)

# Function to display available sheets and let the user select one
def choose_sheet(excel_file):
    # Load all sheet names from the Excel file
    sheet_names = pd.ExcelFile(excel_file).sheet_names
    print("Available sheets:")
    for i, sheet_name in enumerate(sheet_names, start=1):
        print(f"{i}: {sheet_name}")

    # Ask the user to select a sheet by number
    sheet_choice = int(input(f"Select a sheet number (1-{len(sheet_names)}): "))
    return sheet_names[sheet_choice - 1]

# Usage example:
element_file = 'periodic_table.xlsx'  # Path to the Excel file containing valid elements

# Prompt user to select the sheet from the periodic table Excel file if applicable
if element_file.endswith('.xlsx') or element_file.endswith('.xls'):
    element_sheet = choose_sheet(element_file)
else:
    element_sheet = None  # Not applicable for CSV



# User input for maximum elements allowed in the formula
max_elements = int(input("Enter the maximum number of elements allowed in the formula: "))

# Create an instance of the class with max_elements parameter and chosen periodic table sheet
ff = FormulaFilter(element_file, max_elements, element_sheet)

# Process the Excel file and output the result with "_filtered" suffix
ff.process_all_sheets(file_path)