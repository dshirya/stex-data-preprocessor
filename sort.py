import pandas as pd
import re

# Read the Excel file containing the element properties table
element_properties_file = 'element_properties_for_ML.xlsx'
element_properties_df = pd.read_excel(element_properties_file)

# Define the column number for sorting
sorting_column_number = 5 

element_data = dict(zip(element_properties_df.iloc[:, 0], element_properties_df.iloc[:, sorting_column_number]))

# Read the Excel file
input_file = 'Data_PuNi3_Cu3Au.xlsx'
xls = pd.ExcelFile(input_file)

# To work with each sheet
output_data = {}
for sheet_name in xls.sheet_names:
    df = pd.read_excel(input_file, sheet_name=sheet_name)
    
    def rearrange_formula(formula):
        pattern = r'([A-Z][a-z]*)(\d*)'
        elements = re.findall(pattern, formula)
        
        #elements.sort(key=lambda x: (-element_data.get(x[0], float('-inf')), -int(x[1]) if x[1] else -1)) if you want to sort from the most to the least 
        elements.sort(key=lambda x: (element_data.get(x[0], float('inf')), int(x[1]) if x[1] else 1))
        
        rearranged_formula = ''.join([element + index for element, index in elements])
        
        return rearranged_formula
    
    df['Formula'] = df['Formula'].apply(rearrange_formula)
    
    output_data[sheet_name] = df

# Save the sorted DataFrames to a Excel file
output_file = 'Data_PuNi3_Cu3Au_sorted.xlsx'
with pd.ExcelWriter(output_file) as writer:
    for sheet_name, df in output_data.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

