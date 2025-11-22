#!/usr/bin/env python3


from datetime import datetime
import pandas as pd
import sys
from tabulate import tabulate


for_Volkan_df = pd.read_excel('for_Volkan.xlsx')
print(for_Volkan_df.info())
print(for_Volkan_df.head(10))

# for_Volkan_df_cleaned = for_Volkan_df.dropna(subset=['Drywall MSF CM MTD (All)'])

# print(for_Volkan_df_cleaned.info())
# print(for_Volkan_df_cleaned.head(10))

for_Volkan_df = for_Volkan_df.rename(columns={'Branch Code': 'Branch', 'Product Code': 'PRODUCT_CODE'})

df_june = pd.read_excel('June_Net_Net_Price.xlsx')
df_june_additional = pd.read_excel('June_Net_Net_Price_Additional.xlsx')

df_july = pd.read_excel('Net_Net_Price_0728_JULY.xlsx')
df_july = df_july.rename(columns={'VENDOR_NAME': 'Vendor Name', 'LOCATION_CODE': 'Branch'})
df_july = df_july[['Vendor Name', 'Branch', 'PRODUCT_CODE', 'NET_NET_PRICE']]

df_august = pd.read_excel('082925_Input_output_AUGUST.xlsx')
df_august = df_august.rename(columns={'VENDOR_NAME': 'Vendor Name', 'LOCATION_CODE': 'Branch'})
df_august = df_august[['Vendor Name', 'Branch', 'PRODUCT_CODE', 'NET_NET_PRICE']]

df_september = pd.read_excel('092925_Input_output_SEPTEMBER.xlsx')
df_september = df_september.rename(columns={'VENDOR_NAME': 'Vendor Name', 'LOCATION_CODE': 'Branch'})
df_september = df_september[['Vendor Name', 'Branch', 'PRODUCT_CODE', 'NET_NET_PRICE']]

df_october = pd.read_excel('11102025_Input_output_OCTOBER.xlsx')
df_october = df_october.rename(columns={'VENDOR_NAME': 'Vendor Name', 'LOCATION_CODE': 'Branch'})
df_october = df_october[['Vendor Name', 'Branch', 'PRODUCT_CODE', 'NET_NET_PRICE']]

# Vendor mapping
vendor_mapping = {
  'AG': 'AMERICAN GYPSUM EDI',
  'CT': 'CERTAINTEED GYPSUM EDI',
  'GP': 'GEORGIA-PACIFIC GYPSUM (EDI)',
  'NG': 'NATIONAL GYPSUM (EDI)',
  'PB': 'PABCO BUILDING PRODUCTS LLC',
  'PR': 'PANEL REY/ABAMAX  (EDI)',
  'USG': 'UNITED STATES GYPSUM (EDI)'
}

def vendor_mapping_function(row):
    return vendor_mapping[row['Vendor']]

df_june['Vendor Name'] = df_june.apply(vendor_mapping_function, axis=1)

df_june_additional = df_june_additional.rename(columns={'Vendor name': 'Vendor Name', 'Branch Code': 'Branch', 'Pricing Group Code': 'PRODUCT_CODE', 'June missing net prices': 'NET_NET_PRICE'})

df_june_minimal = df_june[['Vendor Name', 'Branch', 'PRODUCT_CODE', 'NET_NET_PRICE']]
df_june_additional_minimal = df_june_additional[['Vendor Name', 'Branch', 'PRODUCT_CODE', 'NET_NET_PRICE']]

# Combine June data
df_june_total = pd.concat([df_june_minimal, df_june_additional_minimal], ignore_index=True)

# Check if rows in df1 match on 'colA' and 'colB' in df2
#  mask = df1['colA'].isin(df2['colA']) & df1['colB'].isin(df2['colB'])
#  matching_rows = df1[mask]
#  print(f"Matching rows based on 'colA' and 'colB':\n{matching_rows}")

# Merge on matching columns (e.g., colA with subA, colB with subB)
# Rename columns in df2 to match df1 for easier merging if necessary
# merged_df = pd.merge(df1.reset_index(), df2.rename(columns={'subA': 'colA', 'subB': 'colB'}), on=['colA', 'colB'], how='inner')   for_Volkan_df_cleaned
#  merged_df = pd.merge(df_june_total, for_Volkan_df.rename(columns={'Branch Code': 'Branch', 'Product Code': 'PRODUCT_CODE'}), on=['Vendor Name', 'Branch', 'PRODUCT_CODE'], how='inner')
merged_df = pd.merge(df_june_total.reset_index(), df_june_total, on=['Vendor Name', 'Branch', 'PRODUCT_CODE'], how='inner')

# Get original row indices from df1
df_june_total_matching_indices = merged_df['index'].tolist()

# Get original row indices from df2 (assuming df2 also has a default integer index)
# You might need to add a temporary index column to df2 before merging if you want its original indices
for_Volkan_df_cleaned_matching_indices = merged_df.index.tolist() # This will be the index of the merged_df, which corresponds to matches in df2

print("df_june_total_matching_indices:", len(df_june_total_matching_indices))
print(df_june_total_matching_indices[0:5])
print()
print("for_Volkan_df_cleaned_matching_indices:", len(for_Volkan_df_cleaned_matching_indices))
print(for_Volkan_df_cleaned_matching_indices[0:5])

for i in df_june_total_matching_indices[0:5]:
    print(i, df_june_total[['Vendor Name', 'Branch', 'PRODUCT_CODE', 'NET_NET_PRICE']].loc[i])
    print()

print()

for i in for_Volkan_df_cleaned_matching_indices[0:5]:
    print(for_Volkan_df[['Vendor Name', 'Branch', 'PRODUCT_CODE']].loc[i])
    print()

sys.exit()

# Define the subrows (columns) to compare
subrows_df1 = ['Branch', 'Product Cod', 'Vendor Name']
subrows_df_june_total = ['Branch', 'PRODUCT_CODE', 'Vendor Name']

# Create a MultiIndex from the subrows of df2 for efficient lookup
# Ensure the order of columns in from_frame matches the order you're comparing
multiindex_df1_renamed = pd.MultiIndex.from_frame(duplicate_rows_last_cleaned[subrows_df1])
multiindex_df_june_total = pd.MultiIndex.from_frame(df_june_total[subrows_df_june_total])

# Find rows in df1 where the subrow combination exists in df2's subrows
# We construct a similar MultiIndex for df1's subrows for comparison
matching_rows_in_df_june_total = df_june_total[pd.MultiIndex.from_frame(df_june_total[subrows_df_june_total]).isin(multiindex_df1_renamed)]
matching_rows_in_df1_renamed = duplicate_rows_last_cleaned[pd.MultiIndex.from_frame(duplicate_rows_last_cleaned[subrows_df1]).isin(multiindex_df_june_total)]

# Get the row numbers (indices) of these matching rows in df1
row_numbers_df_june_total = matching_rows_in_df_june_total.index.tolist()
row_numbers_df1_renamed = matching_rows_in_df1_renamed.index.tolist()

for i in row_numbers_df1_renamed:
    duplicate_rows_last_cleaned.loc[i, 'NET_NET_JUNE_EXISTS'] = 'Yes'
    duplicate_rows_last_cleaned.loc[i, 'NET_NET_JUNE_AMOUNT'] = 100.0

# df1 = df1_renamed.rename(columns={'Branch': 'Branch Code', 'PRODUCT_CODE': 'Product Code'})
# current_datetime = datetime.now()
# formatted_date_time = current_datetime.strftime("%Y-%m-%d_%H_%M_%S")
# output_file = f'output_{formatted_date_time}.xlsx'
# df1.to_excel(output_file, index=False)

value_counts_june = duplicate_rows_last_cleaned['NET_NET_JUNE_EXISTS'].value_counts()
count_YES_june = value_counts_june['Yes']
count_NO_june = value_counts_june['No']
total_june = len(duplicate_rows_last_cleaned)
yes_pct_june = round(100.0 * (count_YES_june / total_june), 4)
no_pct_june = round(100.0 * (count_NO_june / total_june), 4)

print('\nJUNE')
print('  Yes:', count_YES_june, f'[{yes_pct_june} %]')
print('   No:', count_NO_june, f'[{no_pct_june} %]')
print('Total:', total_june)


df_june_forward = df_june_total[['Vendor Name', 'Branch', 'PRODUCT_CODE', 'NET_NET_PRICE']]
# Combine June and July data
df_july_total = pd.concat([df_june_forward, df_july], ignore_index=True).drop_duplicates()
df_july_total = df_july_total.sort_values(by='Branch')
# Add month year label
df_july_total['Month_Year'] = 'July 2025'

# Define the subrows (columns) to compare
subrows_df1 = ['Branch', 'PRODUCT_CODE', 'Vendor Name']
subrows_df_july_total = ['Branch', 'PRODUCT_CODE', 'Vendor Name']

# Create a MultiIndex from the subrows of df2 for efficient lookup
# Ensure the order of columns in from_frame matches the order you're comparing
multiindex_df1_renamed = pd.MultiIndex.from_frame(duplicate_rows_last_cleaned[subrows_df1])
multiindex_df_july_total = pd.MultiIndex.from_frame(df_july_total[subrows_df_july_total])

# Find rows in df1 where the subrow combination exists in df2's subrows
# We construct a similar MultiIndex for df1's subrows for comparison
matching_rows_in_df_july_total = df_july_total[pd.MultiIndex.from_frame(df_july_total[subrows_df_july_total]).isin(multiindex_df1_renamed)]
matching_rows_in_df1_renamed = duplicate_rows_last_cleaned[pd.MultiIndex.from_frame(duplicate_rows_last_cleaned[subrows_df1]).isin(multiindex_df_july_total)]

# Get the row numbers (indices) of these matching rows in df1
row_numbers_df_july_total = matching_rows_in_df_july_total.index.tolist()
row_numbers_df1_renamed = matching_rows_in_df1_renamed.index.tolist()

for i in row_numbers_df1_renamed:
    duplicate_rows_last_cleaned.loc[i, 'NET_NET_JULY_EXISTS'] = 'Yes'
    duplicate_rows_last_cleaned.loc[i, 'NET_NET_JULY_AMOUNT'] = 100.0

# df1 = df1_renamed.rename(columns={'Branch': 'Branch Code', 'PRODUCT_CODE': 'Product Code'})
# current_datetime = datetime.now()
# formatted_date_time = current_datetime.strftime("%Y-%m-%d_%H_%M_%S")
# output_file = f'output_{formatted_date_time}.xlsx'
# df1.to_excel(output_file, index=False)

value_counts_july = duplicate_rows_last_cleaned['NET_NET_JULY_EXISTS'].value_counts()
count_YES_july = value_counts_july['Yes']
count_NO_july = value_counts_july['No']
total_july = len(duplicate_rows_last_cleaned)
yes_pct_july = round(100.0 * (count_YES_july / total_july), 4)
no_pct_july = round(100.0 * (count_NO_july / total_july), 4)

print('\nJULY')
print('  Yes:', count_YES_july, f'[{yes_pct_july} %]')
print('   No:', count_NO_july, f'[{no_pct_july} %]')
print('Total:', total_july) 


df_july_forward = df_july_total[['Vendor Name', 'Branch', 'PRODUCT_CODE', 'NET_NET_PRICE']]
# Combine June and July data
df_august_total = pd.concat([df_july_forward, df_august], ignore_index=True).drop_duplicates()
df_august_total = df_august_total.sort_values(by='Branch')
# Add month year label
df_august_total['Month_Year'] = 'August 2025'

# Define the subrows (columns) to compare
subrows_df1 = ['Branch', 'PRODUCT_CODE', 'Vendor Name']
subrows_df_august_total = ['Branch', 'PRODUCT_CODE', 'Vendor Name']

# Create a MultiIndex from the subrows of df2 for efficient lookup
# Ensure the order of columns in from_frame matches the order you're comparing
multiindex_df1_renamed = pd.MultiIndex.from_frame(duplicate_rows_last_cleaned[subrows_df1])
multiindex_df_august_total = pd.MultiIndex.from_frame(df_august_total[subrows_df_august_total])

# Find rows in df1 where the subrow combination exists in df2's subrows
# We construct a similar MultiIndex for df1's subrows for comparison
matching_rows_in_df_august_total = df_august_total[pd.MultiIndex.from_frame(df_august_total[subrows_df_august_total]).isin(multiindex_df1_renamed)]
matching_rows_in_df1_renamed = duplicate_rows_last_cleaned[pd.MultiIndex.from_frame(duplicate_rows_last_cleaned[subrows_df1]).isin(multiindex_df_august_total)]

# Get the row numbers (indices) of these matching rows in df1
row_numbers_df_august_total = matching_rows_in_df_august_total.index.tolist()
row_numbers_df1_renamed = matching_rows_in_df1_renamed.index.tolist()

for i in row_numbers_df1_renamed:
    duplicate_rows_last_cleaned.loc[i, 'NET_NET_AUGUST_EXISTS'] = 'Yes'
    duplicate_rows_last_cleaned.loc[i, 'NET_NET_AUGUST_AMOUNT'] = 100.0
    

# df1 = df1_renamed.rename(columns={'Branch': 'Branch Code', 'PRODUCT_CODE': 'Product Code'})
# current_datetime = datetime.now()
# formatted_date_time = current_datetime.strftime("%Y-%m-%d_%H_%M_%S")
# output_file = f'output_{formatted_date_time}.xlsx'
# df1.to_excel(output_file, index=False)

value_counts_august = duplicate_rows_last_cleaned['NET_NET_AUGUST_EXISTS'].value_counts()
count_YES_august = value_counts_august['Yes']
count_NO_august = value_counts_august['No']
total_august = len(duplicate_rows_last_cleaned)
yes_pct_august = round(100.0 * (count_YES_august / total_august), 4)
no_pct_august = round(100.0 * (count_NO_august / total_august), 4)

print('\nAUGUST')
print('  Yes:', count_YES_august, f'[{yes_pct_august} %]')
print('   No:', count_NO_august, f'[{no_pct_august} %]')
print('Total:', total_august) 


df_august_forward = df_august_total[['Vendor Name', 'Branch', 'PRODUCT_CODE', 'NET_NET_PRICE']]
# Combine June and July data
df_september_total = pd.concat([df_august_forward, df_september], ignore_index=True).drop_duplicates()
df_september_total = df_september_total.sort_values(by='Branch')
# Add month year label
df_september_total['Month_Year'] = 'September 2025'

# Define the subrows (columns) to compare
subrows_df1 = ['Branch', 'PRODUCT_CODE', 'Vendor Name']
subrows_df_september_total = ['Branch', 'PRODUCT_CODE', 'Vendor Name']

# Create a MultiIndex from the subrows of df2 for efficient lookup
# Ensure the order of columns in from_frame matches the order you're comparing
multiindex_df1_renamed = pd.MultiIndex.from_frame(duplicate_rows_last_cleaned[subrows_df1])
multiindex_df_september_total = pd.MultiIndex.from_frame(df_september_total[subrows_df_september_total])

# Find rows in df1 where the subrow combination exists in df2's subrows
# We construct a similar MultiIndex for df1's subrows for comparison
matching_rows_in_df_september_total = df_september_total[pd.MultiIndex.from_frame(df_september_total[subrows_df_september_total]).isin(multiindex_df1_renamed)]
matching_rows_in_df1_renamed = duplicate_rows_last_cleaned[pd.MultiIndex.from_frame(duplicate_rows_last_cleaned[subrows_df1]).isin(multiindex_df_september_total)]

# Get the row numbers (indices) of these matching rows in df1
row_numbers_df_september_total = matching_rows_in_df_september_total.index.tolist()
row_numbers_df1_renamed = matching_rows_in_df1_renamed.index.tolist()

for i in row_numbers_df1_renamed:
    duplicate_rows_last_cleaned.loc[i, 'NET_NET_SEPTEMBER_EXISTS'] = 'Yes'
    duplicate_rows_last_cleaned.loc[i, 'NET_NET_SEPTEMBER_AMOUNT'] = 100.0

# df1 = df1_renamed.rename(columns={'Branch': 'Branch Code', 'PRODUCT_CODE': 'Product Code'})
# current_datetime = datetime.now()
# formatted_date_time = current_datetime.strftime("%Y-%m-%d_%H_%M_%S")
# output_file = f'output_{formatted_date_time}.xlsx'
# df1.to_excel(output_file, index=False)

value_counts_september = duplicate_rows_last_cleaned['NET_NET_SEPTEMBER_EXISTS'].value_counts()
count_YES_september = value_counts_september['Yes']
count_NO_september = value_counts_september['No']
total_september = len(duplicate_rows_last_cleaned)
yes_pct_september = round(100.0 * (count_YES_september / total_september), 4)
no_pct_september = round(100.0 * (count_NO_september / total_september), 4)

print('\nSEPTEMBER')
print('  Yes:', count_YES_september, f'[{yes_pct_september} %]')
print('   No:', count_NO_september, f'[{no_pct_september} %]')
print('Total:', total_september) 


df_september_forward = df_september_total[['Vendor Name', 'Branch', 'PRODUCT_CODE', 'NET_NET_PRICE']]
# Combine September and October data
df_october_total = pd.concat([df_september_forward, df_october], ignore_index=True).drop_duplicates()
df_october_total = df_october_total.sort_values(by='Branch')
# Add month year label
df_october_total['Month_Year'] = 'OCTOBER 2025'

# Define the subrows (columns) to compare
subrows_df1 = ['Branch', 'PRODUCT_CODE', 'Vendor Name']
subrows_df_october_total = ['Branch', 'PRODUCT_CODE', 'Vendor Name']

# Create a MultiIndex from the subrows of df2 for efficient lookup
# Ensure the order of columns in from_frame matches the order you're comparing
multiindex_df1_renamed = pd.MultiIndex.from_frame(df1_renamed[subrows_df1])
multiindex_df_october_total = pd.MultiIndex.from_frame(df_october_total[subrows_df_october_total])

# Find rows in df1 where the subrow combination exists in df2's subrows
# We construct a similar MultiIndex for df1's subrows for comparison
matching_rows_in_df_october_total = df_october_total[pd.MultiIndex.from_frame(df_october_total[subrows_df_october_total]).isin(multiindex_df1_renamed)]
matching_rows_in_df1_renamed = df1_renamed[pd.MultiIndex.from_frame(df1_renamed[subrows_df1]).isin(multiindex_df_october_total)]

# Get the row numbers (indices) of these matching rows in df1
row_numbers_df_october_total = matching_rows_in_df_october_total.index.tolist()
row_numbers_df1_renamed = matching_rows_in_df1_renamed.index.tolist()

for i in row_numbers_df1_renamed:
    df1_renamed.loc[i, 'NET_NET_OCTOBER_EXISTS'] = 'Yes'

# df1 = df1_renamed.rename(columns={'Branch': 'Branch Code', 'PRODUCT_CODE': 'Product Code'})
# current_datetime = datetime.now()
# formatted_date_time = current_datetime.strftime("%Y-%m-%d_%H_%M_%S")
# output_file = f'output_{formatted_date_time}.xlsx'
# df1.to_excel(output_file, index=False)

value_counts_october = df1_renamed['NET_NET_OCTOBER_EXISTS'].value_counts()
count_YES_october = value_counts_october['Yes']
count_NO_october = value_counts_october['No']
total_october = len(df1_renamed)
yes_pct_october = round(100.0 * (count_YES_october / total_october), 4)
no_pct_october = round(100.0 * (count_NO_october / total_october), 4)

print('\nOCTOBER')
print('  Yes:', count_YES_october)
print('   No:', count_NO_october)
print('Total:', total_october)


# Sample sales data
data = [
    ["NET NET DATA PRSNT", f'{count_YES_june} [{yes_pct_june} %]', f'{count_YES_july} [{yes_pct_july} %]', f'{count_YES_august} [{yes_pct_august} %]', f'{count_YES_september} [{yes_pct_september} %]', f'{count_YES_october} [{yes_pct_october} %]'],
    ["NET NET DATA ABSNT", f'{count_NO_june} [{no_pct_june} %]', f'{count_NO_july} [{no_pct_july} %]', f'{count_NO_august} [{no_pct_august} %]', f'{count_NO_september} [{no_pct_september} %]', f'{count_NO_october} [{no_pct_october} %]'],
    ["TOTAL DATA", total_june, total_july, total_august, total_september, total_october]
]

headers = ["METRIC", "JUNE 2025", "JULY 2025", "AUGUST 2025", "SEPTEMBER 2025", "OCTOBER 2025"]

table_fancy = tabulate(data, headers=headers, tablefmt="fancy_grid")
print()
print(table_fancy)

final_df = duplicate_rows_last_cleaned.rename(columns={'Branch': 'Branch Code', 'PRODUCT_CODE': 'Product Code'})
current_datetime = datetime.now()
formatted_date_time = current_datetime.strftime("%Y-%m-%d_%H_%M_%S")
output_file = f'output_{formatted_date_time}.xlsx'
final_df.to_excel(output_file, index=False)







