import pandas as pd
import sys
from datetime import datetime


print('Begin Test...\n')

# Sample DataFrames
df1 = pd.DataFrame({
    'colA': [1, 2, 3, 4],
    'colB': ['A', 'B', 'C', 'D'],
    'colC': [10, 20, 30, 40]
})

df2 = pd.DataFrame({
    'colA': [3, 4, 5, 6],
    'colB': ['C', 'D', 'E', 'F'],
    'colC': [30, 40, 50, 60]
})

# Merge the DataFrames with an indicator
merged_df = pd.merge(df1.reset_index(), df2.reset_index(), on=list(df1.columns), 
                     how='inner', suffixes=('_df1', '_df2'))

# Extract the original row indices
df1_matching_indices = merged_df['index_df1'].tolist()
df2_matching_indices = merged_df['index_df2'].tolist()

print("Original df1:")
print(df1)
print("\nOriginal df2:")
print(df2)
print("\nRow indices from df1 that exist in df2:", df1_matching_indices)
print("Row indices from df2 that exist in df1:", df2_matching_indices)

print('End Test...\n\n\n')

print('Processing JUNE 2025...')
V = pd.read_excel('from_Volkan.xlsx')
V = V.rename(columns={'Branch Code': 'Branch', 'Pricing Group Code': 'Type'})
# V = V.dropna(subset=['Drywall MSF CM MTD (All)'])
# V = V[['Vendor Name', 'Branch', 'PRODUCT_CODE']]

J = pd.read_excel('June_Net_Net_Price.xlsx')
# J_2 = pd.read_excel('June_Net_Net_Price_Additional.xlsx')

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

J['Vendor Name'] = J.apply(vendor_mapping_function, axis=1)

# J_2 = J_2.rename(columns={'Vendor name': 'Vendor Name', 'Branch Code': 'Branch', 'Pricing Group Code': 'PRODUCT_CODE', 'June missing net prices': 'NET_NET_PRICE'})

NN_JUN = J[['Vendor Name', 'Branch', 'Type', 'NET_NET_PRICE']]
# J_2 = J_2[['Vendor Name', 'Branch', 'Type', 'NET_NET_PRICE']]

# Combine June data
# NN_JUNE = pd.concat([J_1, J_2], ignore_index=True)

NN_JUN['Vendor Name'] = NN_JUN['Vendor Name'].str.strip()
NN_JUN['Type'] = NN_JUN['Type'].str.strip()
V['Vendor Name'] = V['Vendor Name'].str.strip()
V['Type'] = V['Type'].str.strip()

# NN_JUN = NN_JUN.sort_values(by='Branch')
# V = V.sort_values(by='Branch')

# Merge the DataFrames with an indicator
merged_df = pd.merge(V.reset_index(), NN_JUN.reset_index(), on=list(['Vendor Name', 'Branch', 'Type']), 
                     how='inner', suffixes=('_V', '_J'))

# Extract the original row indices
V_matching_indices_JUN = merged_df['index_V'].tolist()
J_matching_indices_JUN = merged_df['index_J'].tolist()

for indx_V, indx_NN in zip(V_matching_indices_JUN, J_matching_indices_JUN):
    are_equal = (
        (V.loc[indx_V, 'Vendor Name'] == NN_JUN.loc[indx_NN, 'Vendor Name']) and
        (V.loc[indx_V, 'Branch'] == NN_JUN.loc[indx_NN, 'Branch']) and
        (V.loc[indx_V, 'Type'] == NN_JUN.loc[indx_NN, 'Type'])
    )
    print(indx_V, indx_NN, are_equal, '--->', float(f'{NN_JUN.loc[indx_NN, 'NET_NET_PRICE']:.2f}'))
    assert are_equal == True
    V.at[indx_V, 'NET_NET_JUN'] = float(f'{NN_JUN.loc[indx_NN, 'NET_NET_PRICE']:.2f}')
    print('Assigned ??? --------->', V.loc[indx_V, 'NET_NET_JUN'])

print('Out of the for loop...')
print(V.info())
print(V.head(10))

final_df = V.rename(columns={'Branch': 'Branch Code', 'Type': 'Pricing Group Code'})
current_datetime = datetime.now()
formatted_date_time = current_datetime.strftime("%Y-%m-%d_%H_%M_%S")
output_file = f'OUTPUT_NET_NET_PRICE_{formatted_date_time}.xlsx'
final_df.to_excel(output_file, index=False)
sys.exit()

#  Now JULY...

print('Processing July 2025...')

JUL = pd.read_excel('Net_Net_Price_0728_JULY.xlsx')
JUL = JUL.rename(columns={'VENDOR_NAME': 'Vendor Name', 'LOCATION_CODE': 'Branch', 'PRICE_TYPE': 'Type'})

JUL = JUL[['Vendor Name', 'Branch', 'Type', 'NET_NET_PRICE']]
NN_JUL = pd.concat([NN_JUN, JUL], ignore_index=True)
NN_JUL = NN_JUL.drop_duplicates(keep='last')

NN_JUL['Vendor Name'] = NN_JUL['Vendor Name'].str.strip()
NN_JUL['Type'] = NN_JUL['Type'].str.strip()
V['Vendor Name'] = V['Vendor Name'].str.strip()
V['Type'] = V['Type'].str.strip()

NN_JULY = NN_JUL.sort_values(by='Branch')
V = V.sort_values(by='Branch')

# Merge the DataFrames with an indicator
merged_df = pd.merge(V.reset_index(), NN_JULY.reset_index(), on=list(['Vendor Name', 'Branch', 'Type']), 
                     how='inner', suffixes=('_V', '_J'))

# Extract the original row indices
V_matching_indices_JUL = merged_df['index_V'].tolist()
J_matching_indices_JUL = merged_df['index_J'].tolist()

for indx_V, indx_NN in zip(V_matching_indices_JUL, J_matching_indices_JUL):
    entry_V = V.loc[indx_V, 'Vendor Name']
    entry_NN = NN_JULY.loc[indx_NN, 'Vendor Name']
    assert entry_V == entry_NN
    V.at[indx_V, 'NET_NET_JUL'] = NN_JULY.loc[indx_NN, 'NET_NET_PRICE']

# V = V.rename(columns={'Branch Code': 'Branch', 'Pricing Group Code': 'Type'})
final_df = V.rename(columns={'Branch': 'Branch Code', 'Type': 'Pricing Group Code'})
current_datetime = datetime.now()
formatted_date_time = current_datetime.strftime("%Y-%m-%d_%H_%M_%S")
output_file = 'missing_NET_NET_COSTS_JUNE_JULY_AUG_SEPT_OCT.xlsx'
final_df.to_excel(output_file, index=False)


#  Now AUGUST...

print('Processing AUGUST 2025...')

# V = V.dropna(subset=['Drywall MSF CM MTD (All)'])
# V = V[['Vendor Name', 'Branch', 'PRODUCT_CODE']]

AUG = pd.read_excel('082925_Input_output_AUGUST.xlsx')
AUG = AUG.rename(columns={'VENDOR_NAME': 'Vendor Name', 'LOCATION_CODE': 'Branch', 'PRICE_TYPE': 'Type'})

AUG = AUG[['Vendor Name', 'Branch', 'Type', 'NET_NET_PRICE']]
NN_AUG = pd.concat([NN_JULY, AUG], ignore_index=True)
NN_AUG = NN_AUG.drop_duplicates(keep='last')

NN_AUG['Vendor Name'] = NN_AUG['Vendor Name'].str.strip()
NN_AUG['Type'] = NN_AUG['Type'].str.strip()
V['Vendor Name'] = V['Vendor Name'].str.strip()
V['Type'] = V['Type'].str.strip()

NN_AUG = NN_AUG.sort_values(by='Branch')
V = V.sort_values(by='Branch')

# Merge the DataFrames with an indicator
merged_df = pd.merge(V.reset_index(), NN_AUG.reset_index(), on=list(['Vendor Name', 'Branch', 'Type']), 
                     how='inner', suffixes=('_V', '_J'))

# Extract the original row indices
V_matching_indices_AUG = merged_df['index_V'].tolist()
J_matching_indices_AUG = merged_df['index_J'].tolist()

for indx_V, indx_NN in zip(V_matching_indices_AUG, J_matching_indices_AUG):
    entry_V = V.loc[indx_V, 'Vendor Name']
    entry_NN = NN_AUG.loc[indx_NN, 'Vendor Name']
    assert entry_V == entry_NN
    V.at[indx_V, 'NET_NET_AUG'] = NN_AUG.loc[indx_NN, 'NET_NET_PRICE']


#  Now SEPTEMBER...

print('Processing SEPTEMBER 2025...')

# V = V.dropna(subset=['Drywall MSF CM MTD (All)'])
# V = V[['Vendor Name', 'Branch', 'PRODUCT_CODE']]

SEP = pd.read_excel('092925_Input_output_SEPTEMBER.xlsx')
SEP = SEP.rename(columns={'VENDOR_NAME': 'Vendor Name', 'LOCATION_CODE': 'Branch', 'PRICE_TYPE': 'Type'})

SEP = SEP[['Vendor Name', 'Branch', 'Type', 'NET_NET_PRICE']]
NN_SEP = pd.concat([NN_AUG, SEP], ignore_index=True)
NN_SEP = NN_SEP.drop_duplicates(keep='last')

NN_SEP['Vendor Name'] = NN_SEP['Vendor Name'].str.strip()
NN_SEP['Type'] = NN_SEP['Type'].str.strip()
V['Vendor Name'] = V['Vendor Name'].str.strip()
V['Type'] = V['Type'].str.strip()

NN_SEP = NN_SEP.sort_values(by='Branch')
V = V.sort_values(by='Branch')

# Merge the DataFrames with an indicator
merged_df = pd.merge(V.reset_index(), NN_SEP.reset_index(), on=list(['Vendor Name', 'Branch', 'Type']), 
                     how='inner', suffixes=('_V', '_J'))

# Extract the original row indices
V_matching_indices_SEP = merged_df['index_V'].tolist()
J_matching_indices_SEP = merged_df['index_J'].tolist()

for indx_V, indx_NN in zip(V_matching_indices_SEP, J_matching_indices_SEP):
    entry_V = V.loc[indx_V, 'Vendor Name']
    entry_NN = NN_SEP.loc[indx_NN, 'Vendor Name']
    assert entry_V == entry_NN
    V.at[indx_V, 'NET_NET_SEP'] = NN_SEP.loc[indx_NN, 'NET_NET_PRICE']


#  Now OCTOBER...

print('Processing OCTOBER 2025...')

# V = V.dropna(subset=['Drywall MSF CM MTD (All)'])
# V = V[['Vendor Name', 'Branch', 'PRODUCT_CODE']]

OCT = pd.read_excel('11102025_Input_output_OCTOBER.xlsx')
OCT = OCT.rename(columns={'VENDOR_NAME': 'Vendor Name', 'LOCATION_CODE': 'Branch', 'PRICE_TYPE': 'Type'})

OCT = OCT[['Vendor Name', 'Branch', 'Type', 'NET_NET_PRICE']]
NN_OCT = pd.concat([NN_SEP, OCT], ignore_index=True)
NN_OCT = NN_OCT.drop_duplicates(keep='last')

NN_OCT['Vendor Name'] = NN_OCT['Vendor Name'].str.strip()
NN_OCT['Type'] = NN_OCT['Type'].str.strip()
V['Vendor Name'] = V['Vendor Name'].str.strip()
V['Type'] = V['Type'].str.strip()

NN_OCT = NN_OCT.sort_values(by='Branch')
V = V.sort_values(by='Branch')

# Merge the DataFrames with an indicator
merged_df = pd.merge(V.reset_index(), NN_OCT.reset_index(), on=list(['Vendor Name', 'Branch', 'Type']), 
                     how='inner', suffixes=('_V', '_J'))

# Extract the original row indices
V_matching_indices_OCT = merged_df['index_V'].tolist()
J_matching_indices_OCT = merged_df['index_J'].tolist()

for indx_V, indx_NN in zip(V_matching_indices_OCT, J_matching_indices_OCT):
    entry_V = V.loc[indx_V, 'Vendor Name']
    entry_NN = NN_OCT.loc[indx_NN, 'Vendor Name']
    assert entry_V == entry_NN
    V.at[indx_V, 'NET_NET_OCT'] = NN_OCT.loc[indx_NN, 'NET_NET_PRICE']

# V = V.rename(columns={'Branch Code': 'Branch', 'Pricing Group Code': 'Type'})
final_df = V.rename(columns={'Branch': 'Branch Code', 'Type': 'Pricing Group Code'})
current_datetime = datetime.now()
formatted_date_time = current_datetime.strftime("%Y-%m-%d_%H_%M_%S")
output_file = f'OUTPUT_NET_NET_COSTS_JUNE_JULY_AUG_SEPT_OCT_{formatted_date_time}.xlsx'
final_df.to_excel(output_file, index=False)


# Data Summary...
print('JUN 2025:')
print('  ', len(V_matching_indices_JUN))
print('  ', len(J_matching_indices_JUN))
print('  ', len(NN_JUN))
print('  ', len(V))
print()

print('JUL 2025:')
print('  ', len(V_matching_indices_JUL))
print('  ', len(J_matching_indices_JUL))
print('  ', len(NN_JUL))
print('  ', len(V))
print()

print('JUNE 2025:')
print('  ', len(V_matching_indices_AUG))
print('  ', len(J_matching_indices_AUG))
print('  ', len(NN_AUG))
print('  ', len(V))
print()

print('SEP 2025:')
print('  ', len(V_matching_indices_SEP))
print('  ', len(J_matching_indices_SEP))
print('  ', len(NN_SEP))
print('  ', len(V))
print()

print('OCT 2025:')
print('  ', len(V_matching_indices_OCT))
print('  ', len(J_matching_indices_OCT))
print('  ', len(NN_OCT))
print('  ', len(V))
print()