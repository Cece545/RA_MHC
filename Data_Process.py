import pandas as pd

def extract_third_column_to_xsl(input_file, output_file):
    # Load the spreadsheet
    data = pd.read_excel(input_file)
    
    # Select the third column, note that Python uses 0-based indexing
    # Change '2' to another number if the third column is differently indexed
    third_column = data.iloc[2:, 2]  # Start from the third row (index 2)

    # Open a text file for writing
    with open(output_file, 'w') as f:
        # Write each item from the third column into the file
        for idx, value in enumerate(third_column, start=1):
            if isinstance(value, str) and len(value) >= 8 and '+' not in value:
                f.write(f'>{idx}\n{value}\n')


input_xls_path = '/Users/ceceliazhang/Downloads/epitope_table_export_1714570631.xlsx'
output_txt_path = '/Users/ceceliazhang/Downloads/output_bindingMHC1.txt'
extract_third_column_to_xsl(input_xls_path, output_txt_path)
