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
for col in idColumns:
    df_t[col] = df_t[col].str.replace('AgencyID-', '')
    df_t.loc[df_t[col].str.startswith('MissionArea', na=True), col] = ''

# REPORT 1: Mismatch Report (SCCM reported Agency ID != at least 1 Tanium reported Agency ID)
# Schema: Workstation Name, Operating System, SCCM Agency ID, Tanium Agency ID 1, Tanium Agency ID 2, …

# Keep only relevant columns from Tanium and SCCM datasets
tCols = ['Encrypted Workstation Name', 'Operating System']
tCols.extend(idColumns)
df_t = df_t[tCols].groupby('Encrypted Workstation Name').first()

sCols = ['Encrypted Workstation Name', 'Agency', 'OS']
df_s = df_s[sCols].groupby('Encrypted Workstation Name').first()

print('Exporting grouped data to Excel')
# df_t.to_excel('./data/df_t_grouped.xlsx')
# df_s.to_excel('./data/df_s_grouped.xlsx')

# Merge DataFrames on workstation
print('Merging datasets and exporting to Excel')
df_workShared = df_t.merge(df_s, how='inner', on='Encrypted Workstation Name')
# df_workShared.to_excel('./data/df_workShared.xlsx')

# print('Finding matching AgencyID classifications')
# df_workShared['matching'] = True
# for col in idColumns:
#     df_workShared['matching'] = df_workShared.loc[
#         df_workShared['matching']
#         & (~df_workShared[col].str.startswith("AgencyID")
#            | df_workShared[col].str.replace('AgencyID-', '') == df_workShared['Agency'])
#     ]
# print('Exporting to Excel')
# df_workShared.to_excel('./data/matching_indicated.xlsx')

# REPORT 2: Joint Report (SCCM reported Agency ID == all Tanium reported Agency ID)
# Schema: Workstation Name, Operating System, SCCM Agency ID, Tanium Agency ID 1, Tanium Agency ID 2, …

# REPORT 3: Only Tanium Report (Workstation Name is only in Tanium)
# Schema: Workstation Name, Operating System, Tanium Agency ID 1, …
print('Generating Tanium-only workstations')
df_t_unique = df_t.merge(df_s, on='Encrypted Workstation Name', indicator=True, how='outer').query(
    '_merge=="left_only"').drop('_merge', axis=1)
# df_t_unique.to_excel('./data/tanium_unique.xlsx')

# REPORT 4: Only SCCM Report (Workstation Name is only in SCCM)
# Schema: Workstation Name, OS, OS Version, SCCM Agency ID
print('Generating SCCM-only workstations')
df_s_unique = df_s.merge(df_t, on='Encrypted Workstation Name', indicator=True, how='outer').query(
    '_merge=="left_only"').drop('_merge', axis=1)
# df_s_unique.to_excel('./data/sccm_unique.xlsx')

# REPORT 5: Agency workstation reporting for SCCM and Tanium software
# Schema: Agency, Work #, SCCM Work #, SCCM Work %, Tanium Work #, Tanium Work %, Work # in Both, Both %
print('Generating agency statistics')
df_idStats = pd.DataFrame(columns=['Agency', 'Unique Workstations',
                                   'Unique SCCM Workstations',
                                   'Percent Unique SCCM Workstations',
                                   'Unique Tanium Workstations',
                                   'Percent Unique Tanium Workstations',
                                   'Shared Work Stations',
                                   'Percent of Shared Workstations'])
# Find all uinque IDs
uniqueIds = set()
allIdCols = ['Agency']
allIdCols.extend(idColumns)
for col in allIdCols:
    uniqueIds.update(df_workShared[col].dropna().unique())

# Genegerate table
print('Creating indicator column')
i = 1
df_union = df_t.merge(df_s, how='outer', on='Encrypted Workstation Name')
for id in uniqueIds:
    # Create Shared Work indicator column
    df_union['id_indicator'] = False
    for col in allIdCols:
        df_union.loc[df_union[col] == id, 'id_indicator'] = True

    # Create SCCM Work indicator column
    df_s['id_indicator'] = False
    df_s.loc[df_s['Agency'] == id, 'id_indicator'] = True

    # Create Tanium Work indicator column
    df_t['id_indicator'] = False
    for col in idColumns:
        df_t.loc[df_t[col] == id, 'id_indicator'] = True

    # Calculate statistics for final table
    df_idStats.at[i, 'Agency'] = id

    uniqueWork = df_union['id_indicator'].values.sum()
    df_idStats.at[i, 'Unique Workstations'] = uniqueWork

    workSCCM = df_s['id_indicator'].values.sum()
    df_idStats.at[i, 'Unique SCCM Workstations'] = workSCCM
    df_idStats.at[i, 'Percent Unique SCCM Workstations'] = workSCCM/uniqueWork

    workTanium = df_t['id_indicator'].values.sum()
    df_idStats.at[i, 'Unique Tanium Workstations'] = workTanium
    df_idStats.at[i, 'Percent Unique Tanium Workstations'] = workTanium/uniqueWork

    sharedWork = workSCCM + workTanium - uniqueWork
    df_idStats.at[i, 'Shared Work Stations'] = sharedWork
    df_idStats.at[i, 'Percent of Shared Workstations'] = sharedWork/uniqueWork
    i += 1
df_idStats.sort_values(by=['Agency'], inplace=True)
df_idStats.to_excel('./data/workstation_statistics.xlsx')
