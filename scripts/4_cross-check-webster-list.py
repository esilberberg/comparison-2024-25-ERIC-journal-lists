import pandas as pd

excel_file = '3-journals-with-openalex-data.xlsx'
webster_file = 'Webster-journals-no-longer-indexed.xlsx'

df = pd.read_excel(excel_file)
webster_df = pd.read_excel(webster_file)

print(webster_df.count())

# target_journals = df[df['issn_l'].isin(webster_df['issn_l'])]
# print(target_journals)