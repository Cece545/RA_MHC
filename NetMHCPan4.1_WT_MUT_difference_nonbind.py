import pandas as pd

# Load the Excel files
nonbind1_w_waff = pd.read_excel('/Users/ceceliazhang/Downloads/nonbind1_wAff.xlsx')
nonbind1_m_waff = pd.read_excel('/Users/ceceliazhang/Downloads/nonbind1_mt_wAff.xlsx')
nonbind2_w_waff = pd.read_excel('/Users/ceceliazhang/Downloads/nonbind2_wAff.xlsx')
nonbind2_m_waff = pd.read_excel('/Users/ceceliazhang/Downloads/nonbind2_mt_wAff.xlsx')
nonbind3_w_waff = pd.read_excel('/Users/ceceliazhang/Downloads/nonbind3_wAff.xlsx')
nonbind3_m_waff = pd.read_excel('/Users/ceceliazhang/Downloads/nonbind3_mt_wAff.xlsx')
nonbind4_w_waff = pd.read_excel('/Users/ceceliazhang/Downloads/nonbind4_wAff.xlsx')
nonbind4_m_waff = pd.read_excel('/Users/ceceliazhang/Downloads/nonbind4_mt_wAff.xlsx')

# Concatenate the wildtype DataFrames vertically
wildtype_combined = pd.concat([nonbind1_w_waff, nonbind2_w_waff, nonbind3_w_waff, nonbind4_w_waff], ignore_index=True, axis=0)

# Concatenate the mutant DataFrames vertically
mutant_combined = pd.concat([nonbind1_m_waff, nonbind2_m_waff, nonbind3_m_waff, nonbind4_m_waff], ignore_index=True, axis=0)

# Concatenate the combined wildtype and mutant DataFrames horizontally
combined_nonbind = pd.concat([wildtype_combined, mutant_combined], axis=1)

# Ensure the peptide sequences match
wildtype_identity_col = 1  
mutant_identity_col = 1    

if not wildtype_combined.iloc[:, wildtype_identity_col].equals(mutant_combined.iloc[:, mutant_identity_col]):
    raise ValueError("Identity mismatch between Wildtype and Mutant samples.")

# Add columns for absolute differences and percentage differences based on the 8th column
# Convert the relevant columns to numeric, handling non-numeric values by converting them to NaN
data_column_wt = 11  # Index for data column in WT
data_column_mut = 11 + len(wildtype_combined.columns)  # Index for data column in Mutant

combined_nonbind.iloc[:, data_column_wt] = pd.to_numeric(combined_nonbind.iloc[:, data_column_wt], errors='coerce')
combined_nonbind.iloc[:, data_column_mut] = pd.to_numeric(combined_nonbind.iloc[:, data_column_mut], errors='coerce')

# Add columns for absolute differences and percentage differences
combined_nonbind['Absolute Difference'] = (combined_nonbind.iloc[:, data_column_mut] - combined_nonbind.iloc[:, data_column_wt])
combined_nonbind['Percent Difference'] = combined_nonbind['Absolute Difference'] / combined_nonbind.iloc[:, data_column_wt]


# Write the final DataFrame to a new Excel file
combined_nonbind.to_excel('/Users/ceceliazhang/Downloads/combined_nonbinder_data_all.xlsx', index=False)


sorted_df_all = combined_nonbind.sort_values(by='Absolute Difference', ascending=True)
sorted_pctdf_all = combined_nonbind.sort_values(by='Percent Difference', ascending=True)

sorted_df_all.to_excel('/Users/ceceliazhang/Downloads/sorted_numerical_nonbind_all.xlsx', index=False)
sorted_pctdf_all.to_excel('/Users/ceceliazhang/Downloads/sorted_percent_nonbind_all.xlsx', index=False)