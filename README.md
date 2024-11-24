# STEx data preprocessor  

This tool processes an Excel file containing chemical formulas, filters them based on a set of valid elements and the maximum number of elements in a formula, and generates a filtered output Excel file.

## Features  
1. Filter by Valid Elements: Ensures formulas only contain elements listed in a specified periodic table file
2. Sort the formulas by Mendeleev number (ascending order)
3. Limit Number of Elements: Filters formulas with up to a user-defined maximum number of elements
4. Support for Multiple Sheets: Processes all sheets in the input Excel file
5. Interactive Sheet Selection: Allows users to select a sheet from the periodic table file for valid elements

## Example Workflow  

1. Adjust the file periodic_table.xlsx, so there will be only the elements in your table
2. Change the file_path to the file with your data (double check that the Formula column titled the right way)
3. Run the script:
```
python preprocessor.py
```
4. Choose the table (sheet in the periodic_table.xlsx file that you want to preprocess to
5. Check the filtered output file (name_preprocessed.xlsx) in the same directory. 

## Requirements

```
pip install pandas openpyxl
```
