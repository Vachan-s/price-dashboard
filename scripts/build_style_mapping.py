import pandas as pd
import json
import re

# Load Myntra dump
myntra_df = pd.read_excel('data/Myntra_Dump_June.xlsx')

def extract_prefix(code):
    if pd.isna(code): return None
    m = re.match(r'10(X[A-Z]+\d+)', str(code))
    return m.group(1) if m else None

myntra_df['sku_prefix'] = myntra_df['10xu SKU code'].apply(extract_prefix)
myntra_unique = myntra_df.dropna(subset=['sku_prefix']).groupby('sku_prefix').first().reset_index()

# Load Catalog Master - D2C sheet (real headers are on the 3rd row of the sheet)
catalog_df = pd.read_excel('data/Catalog_Master.xlsx', sheet_name='Catalog - SKU level - D2C', header=2)
catalog_df = catalog_df[catalog_df['Style ID'].notna() & catalog_df['Style ID'].str.startswith('X', na=False)]
catalog_unique = catalog_df.groupby('Style ID').first().reset_index()

# Build mapping
mapping = {}
for _, row in myntra_unique.iterrows():
    sku = row['sku_prefix']
    mapping[sku] = {
        'myntra_style_name': str(row.get('style_name', '') or ''),
        'myntra_display_name': str(row.get('list_display_name', '') or ''),
        'gender': str(row.get('gender', '') or '')
    }

for _, row in catalog_unique.iterrows():
    sku = str(row['Style ID'])
    if sku not in mapping:
        mapping[sku] = {'myntra_style_name': '', 'myntra_display_name': '', 'gender': ''}
    mapping[sku]['tenxyou_website_name'] = str(row.get('Website Item Name', '') or '')

with open('data/style_mapping.json', 'w') as f:
    json.dump(mapping, f, indent=2)

print(f'Built mapping for {len(mapping)} SKUs')
for k, v in list(mapping.items())[:5]:
    print(f'{k}: {v}')
