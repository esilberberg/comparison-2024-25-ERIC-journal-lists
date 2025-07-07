import pandas as pd
import requests

excel_file_name = '2-journals-and-issn.xlsx'


df = pd.read_excel(excel_file_name)

def get_openalex_data(issn):
    """Fetches select journal metadata from OpenAlex API via the ISSN-L."""
    if issn is None or issn == "":
        print('No ISSN-L')
        # Returns None values to avoid error when assigning later to Dataframe
        return (None,) * 10 
    
    base_url = 'https://api.openalex.org/sources/issn:'
    api_endpoint = base_url + str(issn)
    response = requests.get(api_endpoint)

    if response.status_code != 200:
        print(f"Error retrieving data! (ISSN: {issn}) (Status Code: {response.status_code})")
        return (None,) * 10  # Returns None values to avoid error when assigning later to Dataframe

    data = response.json()

    publisher = data.get('host_organization_name')
    homepage = data.get('homepage_url')
    is_oa = data.get('is_oa')
    is_in_doaj = data.get('is_in_doaj')
    apc_usd = data.get('apc_usd')
    country = data.get('country_code')

    if 'societies' not in data or not data['societies']:
        organization = None
    else:
        organization = data['societies'][0]['organization']

    try:
        h_index = data['summary_stats']['h_index']
    except (KeyError, IndexError):
        h_index = None
    
    try:
        impact_factor = data['summary_stats']['2yr_mean_citedness']
    except (KeyError, IndexError):
        impact_factor = None

    try:
        wikidata_id = data['ids']['wikidata']
        if wikidata_id:
            wikidata_id = wikidata_id.replace('https://www.wikidata.org/entity/', '')
            wikidata_id = wikidata_id.replace('http://www.wikidata.org/entity/', '')
    except (KeyError, IndexError):
        wikidata_id = None

    print(issn)
    print(publisher, homepage, organization, country, h_index, impact_factor, is_oa, is_in_doaj, apc_usd, wikidata_id)
    return publisher, homepage, organization, country, h_index, impact_factor, is_oa, is_in_doaj, apc_usd, wikidata_id



new_columns = ['publisher', 'homepage', 'organization', 'country', 'h_index', 'impact_factor', 'is_oa', 'is_in_doaj', 'apc_usd', 'wikidata_id']


df[new_columns] = df['issn_l'].apply(
    lambda x: pd.Series(get_openalex_data(x), index=new_columns)
)

output_excel_file = "journals-with-openalex-data.xlsx"

df.to_excel(output_excel_file, index=False)