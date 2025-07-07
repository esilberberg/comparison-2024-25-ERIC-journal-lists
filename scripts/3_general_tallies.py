import pandas as pd


excel_file_name = '3-journals-with-openalex-data.xlsx'

df = pd.read_excel(excel_file_name)

# Journal counts per category
df.count()
all_in_2024 = df[df['Not_on_2024'] != 'X']
all_in_2025 = df[df['Not_on_2025'] != 'X']
unique_in_2024 = df[df['Not_on_2025'] == 'X']
unique_in_2025 = df[df['Not_on_2024'] == 'X']
inactive_indexing = df[df['Indexing Status'] == 'Inactive']
active_indexing_2025 = all_in_2025[all_in_2025['Indexing Status'] == 'Active']


categories = {
    'all': df,
    'All in 2024': all_in_2024,
    'All in 2025': all_in_2025,
    'Active indexing 2025': active_indexing_2025,
    'Inactive indexing': inactive_indexing,
    'Unique 2024': unique_in_2024,
    'Unique 2025': unique_in_2025  
}

# Percentage OA
# for name, value in categories.items():
#     print(f"Variable Name: {name}---------------")
#     print(value['is_oa'].count())
#     print(value['is_oa'].value_counts())
#     print(value['is_oa'].value_counts(normalize=True))

# Percentage in DOAJ
# for name, value in categories.items():
#     print(f"Variable Name: {name}---------------")
#     print(value['is_in_doaj'].count())
#     print(value['is_in_doaj'].value_counts())
#     print(value['is_in_doaj'].value_counts(normalize=True))

# Median H-Index
# for name, value in categories.items():
#     print(f"Variable Name: {name}---------------")
#     print(value['h_index'].count())
#     print(value['h_index'].median())

# Countries
# for name, value in categories.items():
#     print(f"Variable Name: {name}---------------")
#     print(value['country'].count())
#     print(value['country'].value_counts())
#     print(value['country'].value_counts(normalize=True))


# output_excel_file = 'country_analysis_summary.xlsx'
# with pd.ExcelWriter(output_excel_file, engine='xlsxwriter') as writer:
#     for name, df_category in categories.items():
#         # Get the country counts (absolute)
#         country_counts = df_category['country'].value_counts().reset_index()
#         country_counts.columns = ['Country', 'Count'] # Rename columns for clarity

#         # Get the country percentages
#         country_percentages = df_category['country'].value_counts(normalize=True).reset_index()
#         country_percentages.columns = ['Country', 'Percentage'] # Rename columns

#         summary_df = pd.merge(country_counts, country_percentages, on='Country', how='left')

#         # Add a total count row
#         total_count = len(df_category['country'].dropna()) # Count non-null countries
#         total_row = pd.DataFrame([{'Country': 'TOTAL (non-null)', 'Count': total_count, 'Percentage': 1.0}]) # 1.0 for 100%
#         summary_df = pd.concat([summary_df, total_row], ignore_index=True)


#         # Write the combined summary DataFrame to a new sheet
#         sheet_name = name.replace(' ', '_').replace('/', '_').replace('\\', '_')[:31] # Max 31 chars
#         summary_df.to_excel(writer, sheet_name=sheet_name, index=False)

#         workbook = writer.book
#         worksheet = writer.sheets[sheet_name]
#         worksheet.write(0, 0, f"Country Analysis for: {name}") # Add title to cell A1



# Publishers
# for name, value in categories.items():
#     print(f"Variable Name: {name}---------------")
#     print(value['publisher'].count())
#     print(value['publisher'].value_counts())
#     print(value['publisher'].value_counts(normalize=True))


# output_excel_file = 'publisher_analysis_summary.xlsx'
# with pd.ExcelWriter(output_excel_file, engine='xlsxwriter') as writer:
#     for name, df_category in categories.items():
#         # Get the country counts (absolute)
#         country_counts = df_category['publisher'].value_counts().reset_index()
#         country_counts.columns = ['Publisher', 'Count'] # Rename columns for clarity

#         # Get the country percentages
#         country_percentages = df_category['publisher'].value_counts(normalize=True).reset_index()
#         country_percentages.columns = ['Publisher', 'Percentage'] # Rename columns

#         # Optionally, you might want to merge these into one DataFrame for each sheet
#         # Or write them to separate sections of the same sheet
#         summary_df = pd.merge(country_counts, country_percentages, on='Publisher', how='left')

#         # Add a total count row
#         total_count = len(df_category['publisher'].dropna()) # Count non-null countries
#         total_row = pd.DataFrame([{'Publishers': 'TOTAL (non-null)', 'Count': total_count, 'Percentage': 1.0}]) # 1.0 for 100%
#         summary_df = pd.concat([summary_df, total_row], ignore_index=True)


#         # Write the combined summary DataFrame to a new sheet
#         # Use a sanitized sheet name (Excel sheet names have limits)
#         sheet_name = name.replace(' ', '_').replace('/', '_').replace('\\', '_')[:31] # Max 31 chars
#         summary_df.to_excel(writer, sheet_name=sheet_name, index=False)

#         # Optionally, add some introductory text if you want
#         workbook = writer.book
#         worksheet = writer.sheets[sheet_name]
#         worksheet.write(0, 0, f"Publisher Analysis for: {name}") # Add title to cell A1
#         # Push the DataFrame down to start from row 2 (index 1) to make space for title
#         # The to_excel call above would need startrow=1 if you use this
#         # For simplicity, I've just added a cell A1 title and let the DF start from A1.
