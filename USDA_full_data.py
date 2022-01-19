# Find reports unique in Tanium and SCCM
import matplotlib.pyplot as plt
import pandas as pd
import os
os.makedirs('./data', exist_ok=True)

# Read Tanium input data as DataFrame df
df_t = pd.read_excel('Tanium_data_sample.xlsx')
df_s = pd.read_excel('SCCM_data_sample.xlsx')

# # Clean data by removing baselining and duplicates
df_t = df_t[df_t['Usage'] != 'Baselining'].drop_duplicates()
df_t = df_t[df_t['Name'] != 'Adobe Reader and Acrobat Manager']

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

# Create a sub-dataframe with desired columns from Tanium and SCCM
tCols = ['Encrypted Workstation Name', 'Operating System']
tCols.extend(idColumns)
df_t_sub = df_t[tCols]

df_s_sub = df_s[['Encrypted Workstation Name', 'Agency', 'OS']]

df_t_sub.to_excel('./data/tanium_sub.xlsx')
df_s_sub.to_excel('./data/sccm_sub.xlsx')

# Merge dataframes on workstation and index with workstation
# df_workShared = df_t_sub.merge(
#     df_s_sub, on='Encrypted Workstation Name', how='inner', indicator=True)
# df_workShared.set_index('Encrypted Workstation Name')
# df_workShared.to_excel('./data/work_shared.xlsx')

# version 1: If Agency ID does not equal at least 1 Tanium, add to excel
# df_idMismatch = df_workShared.iloc[0:0]


# Do not iterate manually
# for index, row in df_workShared.iterrows():
#     print('This is the row:' + row['Encrypted Workstation Name'])
#     mistmatch = False
#     for col in idColumns:
#         if (row['Agency'] != row[col] and not pd.isnull(row[col])):
#             mismatch = True
#             if mismatch:
#                 break
#     if mismatch:
#         df_idMismatch.append(df_workShared)
#         print('row appended')
# print('done loading')

# print(df_idMismatch)
# df_idMismatch.to_excel('./data/mismatched_usage.xlsx')

# version 2: If Agency ID does not equal at least 1 Tanium, add to excel
# df_idMismatch = df_workShared.iloc[0:0]
# for col in idColumns:
#     df_idMismatch.append(df_workShared.loc[df_workShared['Agency'] != df_workShared[col]])
# print(df_idMismatch)
# df_idMismatch.to_excel('./data/mismatched_usage.xlsx')

# REPORT 2: Joint Report (SCCM reported Agency ID == all Tanium reported Agency ID)
# Merge dataframes on workstation and index with workstation


# REPORT 3: Only Tanium Report (Workstation Name is only in Tanium)
# Schema: Workstation Name, Operating System, Tanium Agency ID 1, …
df_t_unique = df_t_sub.merge(df_s_sub, indicator=True, how='outer').query(
    '_merge=="left_only"').drop('_merge', axis=1)
df_t_unique.to_excel('./data/tanium_unique.xlsx')

# REPORT 4: Only SCCM Report (Workstation Name is only in SCCM)
# Schema: Workstation Name, OS, OS Version, SCCM Agency ID
df_s_unique = df_s_sub.merge(df_t_sub, indicator=True, how='outer').query(
    '_merge=="left_only"').drop('_merge', axis=1)
df_s_unique.to_excel('./data/sccm_unique.xlsx')

# REPORT 5: Agency workstation reporting for SCCM and Tanium software
# Schema: Agency, Work #, SCCM Work #, SCCM Work %, Tanium Work #, Tanium Work %, Work # in Both, Both %
