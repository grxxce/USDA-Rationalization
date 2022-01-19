# Find reports unique in Tanium and SCCM
import matplotlib.pyplot as plt
import pandas as pd
import os

os.makedirs('./data', exist_ok=True)

# Read Tanium input data as DataFrame df
print('Reading in Tanium and SCCM data')
df_t = pd.read_excel('Tanium_data.xlsx')
df_s = pd.read_excel('SCCM_data.xlsx')

# Indicate all columns where AgencyID may be located for Tanium
idColumns = [
    'Asset - Custom Tags.2.1',
    'Asset - Custom Tags.2.2.1',
    'Asset - Custom Tags.2.2.2.1',
    'Asset - Custom Tags.2.2.2.2.1',
    'Asset - Custom Tags.2.2.2.2.2.1',
    'Asset - Custom Tags.2.2.2.2.2.2.1',
    'Asset - Custom Tags.2.2.2.2.2.2.2.1',
    'Asset - Custom Tags.2.2.2.2.2.2.2.2.2.1'
]

# Clean all Agency IDs from Tanium data
for column in idColumns:
    df_t[column] = df_t[column].str.replace("AgencyID-", "")

# REPORT 1: Mismatch Report (SCCM reported Agency ID != at least 1 Tanium reported Agency ID)
# Schema: Workstation Name, Operating System, SCCM Agency ID, Tanium Agency ID 1, Tanium Agency ID 2, …

# Keep only relevant columns from Tanium and SCCM datasets
tCols = ['Encrypted Workstation Name', 'Operating System']
tCols.extend(idColumns)
df_t = df_t[tCols].groupby('Encrypted Workstation Name').first()

sCols = ['Encrypted Workstation Name', 'Agency', 'OS']
df_s = df_s[sCols].groupby('Encrypted Workstation Name').first()

print('Exporting grouped data to Excel')
df_t.to_excel('./data/df_t_grouped.xlsx')
df_s.to_excel('./data/df_s_grouped.xlsx')

# Merge DataFrames on workstation
print('Merging datasets and exporting to Excel')
df_workShared = df_t.merge(df_s, how='inner', on='Encrypted Workstation Name')
df_workShared.to_excel('./data/df_workShared.xlsx')

# print('Finding matching AgencyID classifications')
# df_workShared['matching'] = True
# for col in idColumns:
#     df_workShared['matching'] = df_workShared.loc[
#         df_workShared['matching']
#         & (~df_workShared[col].str.startswith("AgencyID")
#         | df_workShared[col].str.replace('AgencyID-', '') == df_workShared['Agency'])
#         ]
# print('Exporting to Excel')
# df_workShared.to_excel('./data/matching_indicated.xlsx')

# REPORT 2: Joint Report (SCCM reported Agency ID == all Tanium reported Agency ID)
# Schema: Workstation Name, Operating System, SCCM Agency ID, Tanium Agency ID 1, Tanium Agency ID 2, …

# REPORT 3: Only Tanium Report (Workstation Name is only in Tanium)
# Schema: Workstation Name, Operating System, Tanium Agency ID 1, …

# REPORT 4: Only SCCM Report (Workstation Name is only in SCCM)
# Schema: Workstation Name, OS, OS Version, SCCM Agency ID
