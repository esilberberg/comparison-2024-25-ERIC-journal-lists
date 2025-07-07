import pandas as pd


excel_file_name = '3-journals-with-openalex-data.xlsx'

df = pd.read_excel(excel_file_name)

inactive_indexing = df[df['Indexing Status'] == 'Inactive']

journal_titles = []

def get_journal_title(row):
    title = row['Journal Name']
    journal_titles.append(str(title))
    return title

inactive_indexing.apply(get_journal_title, axis=1)

print(journal_titles)