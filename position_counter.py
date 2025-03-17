import pandas as pd
import re
from collections import defaultdict

def parse_formula(formula):
    """Parses a ternary chemical formula into its three positions."""
    elements = re.findall(r'[A-Z][a-z]?', formula)
    return elements if len(elements) == 3 else None

def count_element_positions(df, column_name="Formula"):
    """Counts occurrences of elements in each position."""
    position_counts = {1: defaultdict(int), 2: defaultdict(int), 3: defaultdict(int)}
    
    for formula in df[column_name].dropna():
        parsed = parse_formula(formula)
        if parsed:
            for pos, element in enumerate(parsed, 1):
                position_counts[pos][element] += 1

    return position_counts

def process_excel(file_path):
    """Processes an Excel file with multiple sheets and counts element positions."""
    xls = pd.ExcelFile(file_path)
    results = {}

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        if "Formula" in df.columns:
            results[sheet] = count_element_positions(df)
    
    return results

# Example usage:
file_path = "Project_2_processed.xlsx"  # Replace with your actual file path
output = process_excel(file_path)

# Print the results
for sheet, counts in output.items():
    print(f"Sheet: {sheet}")
    for pos, elements in counts.items():
        print(f"  Position {pos}: {dict(elements)}")