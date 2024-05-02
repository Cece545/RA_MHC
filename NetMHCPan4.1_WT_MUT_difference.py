import pandas as pd

# Load the Excel files
wildtype_ra = pd.read_csv('/Users/ceceliazhang/Downloads/RArelated_wt.xls', delimiter='\t')
wildtype_nonbind = pd.read_csv('/Users/ceceliazhang/Downloads/nonbind_wt.xls', delimiter='\t')
mutant_ra = pd.read_csv('/Users/ceceliazhang/Downloads/RArelated_mutant.xls', delimiter='\t')
mutant_nonbind = pd.read_csv('/Users/ceceliazhang/Downloads/nonbind_mutant.xls', delimiter='\t')

# Concatenate the wildtype DataFrames vertically
# wildtype_combined = pd.concat([wildtype_ra, wildtype_nonbind], ignore_index=True)

# Concatenate the mutant DataFrames vertically
# mutant_combined = pd.concat([mutant_ra, mutant_nonbind], ignore_index=True)

# Concatenate the combined wildtype and mutant DataFrames horizontally
combined_ra = pd.concat([wildtype_ra, mutant_ra], axis=1)
combined_nonbind = pd.concat([wildtype_nonbind, mutant_nonbind], axis=1)

# Ensure the peptide sequences match
wildtype_identity_col = 1  
mutant_identity_col = 1    

if not wildtype_ra.iloc[:, wildtype_identity_col].equals(mutant_ra.iloc[:, mutant_identity_col]):
    raise ValueError("Identity mismatch between Wildtype and Mutant samples.")
if not wildtype_nonbind.iloc[:, wildtype_identity_col].equals(mutant_nonbind.iloc[:, mutant_identity_col]):
    raise ValueError("Identity mismatch between Wildtype and Mutant samples.")

# Add columns for absolute differences and percentage differences based on the 8th column
# Convert the relevant columns to numeric, handling non-numeric values by converting them to NaN
data_column_wt = 7  # Index for data column in WT
data_column_mut = 7 + len(wildtype_ra.columns)  # Index for data column in Mutant

combined_ra.iloc[:, data_column_wt] = pd.to_numeric(combined_ra.iloc[:, data_column_wt], errors='coerce')
combined_ra.iloc[:, data_column_mut] = pd.to_numeric(combined_ra.iloc[:, data_column_mut], errors='coerce')
combined_nonbind.iloc[:, data_column_wt] = pd.to_numeric(combined_nonbind.iloc[:, data_column_wt], errors='coerce')
combined_nonbind.iloc[:, data_column_mut] = pd.to_numeric(combined_nonbind.iloc[:, data_column_mut], errors='coerce')

# Add columns for absolute differences and percentage differences
combined_ra['Absolute Difference'] = (combined_ra.iloc[:, data_column_mut] - combined_ra.iloc[:, data_column_wt])
combined_ra['Percent Difference'] = combined_ra['Absolute Difference'] / combined_ra.iloc[:, data_column_wt]
combined_nonbind['Absolute Difference'] = (combined_nonbind.iloc[:, data_column_mut] - combined_nonbind.iloc[:, data_column_wt])
combined_nonbind['Percent Difference'] = combined_nonbind['Absolute Difference'] / combined_nonbind.iloc[:, data_column_wt]


# Write the final DataFrame to a new Excel file
combined_ra.to_excel('/Users/ceceliazhang/Downloads/combined_rarelated_data.xlsx', index=False)
combined_nonbind.to_excel('/Users/ceceliazhang/Downloads/combined_nonbinder_data.xlsx', index=False)

sorted_df_ra = combined_ra.sort_values(by='Absolute Difference', ascending=True)
sorted_pctdf_ra = combined_ra.sort_values(by='Percent Difference', ascending=True)

sorted_df_nb = combined_nonbind.sort_values(by='Absolute Difference', ascending=True)
sorted_pctdf_nb = combined_nonbind.sort_values(by='Percent Difference', ascending=True)

sorted_df_ra.to_excel('/Users/ceceliazhang/Downloads/sorted_numerical_rarelated.xlsx', index=False)
sorted_pctdf_ra.to_excel('/Users/ceceliazhang/Downloads/sorted_percent_rarelated.xlsx', index=False)
sorted_df_nb.to_excel('/Users/ceceliazhang/Downloads/sorted_numerical_nonbind.xlsx', index=False)
sorted_pctdf_nb.to_excel('/Users/ceceliazhang/Downloads/sorted_percent_nonbind.xlsx', index=False)