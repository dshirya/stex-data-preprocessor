import pandas as pd
import re
import os

# User-configurable settings
element_file = 'periodic_table.xlsx'  # File containing valid elements
input_file = 'CsCl-cP2.csv'            # Input CSV file with formulas
max_elements = 5                       # Maximum number of elements allowed in the Formula column
max_elements_site = 4                  # Maximum number of elements allowed in each site occupancy column
element_sheet = None                   # Set to None for CSV or specify sheet for Excel (for element_file)

class FormulaHandler:
    def __init__(self, element_file, max_elements=2, max_elements_site=1, element_sheet=None):
        self.valid_elements = self.load_valid_elements(element_file, element_sheet)
        self.max_elements = max_elements
        self.max_elements_site = max_elements_site

    def load_valid_elements(self, element_file, sheet_name):
        if element_file.endswith('.xlsx') or element_file.endswith('.xls'):
            df_elements = pd.read_excel(element_file, sheet_name=sheet_name, header=None)
        else:
            df_elements = pd.read_csv(element_file, header=None)
        return set(df_elements[0])

    def parse_formula(self, formula):
        # Skip processing if formula is missing or not a valid string.
        if pd.isna(formula) or not isinstance(formula, str) or not formula.strip():
            return {}
        pattern = r'([A-Z][a-z]*)(\d*\.?\d*)'
        matches = re.findall(pattern, formula)
        elements = {}
        for element, count in matches:
            elements[element] = float(count) if count else 1
        return elements

    def filter_formulas(self, df):
        # Determine the site occupancy columns: all columns after "Num Elements"
        try:
            num_elements_idx = df.columns.get_loc("Num Elements")
        except KeyError:
            num_elements_idx = len(df.columns)
        site_columns = df.columns[num_elements_idx+1:]

        def filter_row(row):
            # Check the main 'Formula' column using max_elements.
            formula_str = row['Formula']
            parsed_formula = self.parse_formula(formula_str)
            if not parsed_formula or len(parsed_formula) > self.max_elements:
                return False
            if any(element not in self.valid_elements for element in parsed_formula):
                return False

            # Check each site occupancy column individually using max_elements_site.
            for col in site_columns:
                cell = row[col]
                if pd.isna(cell) or not isinstance(cell, str) or not cell.strip():
                    continue
                parsed_site = self.parse_formula(cell)
                if not parsed_site:
                    continue
                if any(element not in self.valid_elements for element in parsed_site):
                    return False
                if len(parsed_site) > self.max_elements_site:
                    return False

            return True

        return df[df.apply(filter_row, axis=1)]
    
    def process_csv_file(self, file_path, output_file_path=None):
        # Read the CSV input file.
        df = pd.read_csv(file_path)
        df_filtered = self.filter_formulas(df)
        if output_file_path is None:
            base_name, ext = os.path.splitext(file_path)
            output_file_path = f"{base_name}_processed{ext}"
        # Write the processed data to a new CSV file.
        df_filtered.to_csv(output_file_path, index=False)
        print(f"Processed data written to {output_file_path}")
    
    def separate_by_notes(self, processed_file_path):
        """
        Reads the processed CSV file, groups rows by the 'Notes' column (with empty notes treated as 'rt'),
        and writes each group to a separate CSV file. The output file name contains the note value.
        """
        df = pd.read_csv(processed_file_path)
        # Create a grouping column: empty notes become 'rt'
        df['NoteGroup'] = df['Notes'].fillna('').apply(lambda x: re.sub(r'[\\/]', '_', x.strip()) if x.strip() != '' else 'rt')
        base_name, ext = os.path.splitext(processed_file_path)
        for note, group in df.groupby('NoteGroup'):
            output_file = f"{base_name}_{note}{ext}"
            group.to_csv(output_file, index=False)
            print(f"Written {len(group)} rows to {output_file}")

if __name__ == "__main__":
    # If the element file is an Excel file with multiple sheets, let the user choose the appropriate sheet.
    if element_file.endswith('.xlsx') or element_file.endswith('.xls'):
        sheet_names = pd.ExcelFile(element_file).sheet_names
        if len(sheet_names) > 1:
            print("Available sheets in the periodic table file:")
            for i, sheet_name in enumerate(sheet_names, start=1):
                print(f"{i}: {sheet_name}")
            choice = int(input(f"Select a sheet number (1-{len(sheet_names)}): "))
            element_sheet = sheet_names[choice - 1]

    fh = FormulaHandler(element_file, max_elements, max_elements_site, element_sheet)
    # Process the CSV file.
    fh.process_csv_file(input_file)
    
    # Separate processed data by 'Notes'.
    processed_file = os.path.splitext(input_file)[0] + "_processed.csv"
    fh.separate_by_notes(processed_file)