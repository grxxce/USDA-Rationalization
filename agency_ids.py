import matplotlib.pyplot as plt
import numpy as np
import os
import pandas as pd

# Import Tanium and SCCM input data as DataFrames
print('Importing Tanium and SCCM data')
df_t = pd.read_excel('tanium.xlsx')
df_s = pd.read_excel('sccm.xlsx')

print('Cleaning and preparing data')

os.makedirs('./data', exist_ok=True)

# Remove Tanium entries with invalid 'Usage' values
valid_usages = {
    'Baselining',
    'Usage not detected',
    'Limited',
    'Normal',
    'High'
}
df_t = df_t[df_t['Usage'].isin(valid_usages)]

# Indicate all Tanium columns containing AgencyIDs
id_columns = [
    'Asset - Custom Tags.2.1',
    'Asset - Custom Tags.2.2.1',
    'Asset - Custom Tags.2.2.2.1',
    'Asset - Custom Tags.2.2.2.2.1',
    'Asset - Custom Tags.2.2.2.2.2.1',
    'Asset - Custom Tags.2.2.2.2.2.2.1',
    'Asset - Custom Tags.2.2.2.2.2.2.2.1',
    'Asset - Custom Tags.2.2.2.2.2.2.2.2.2.1'
]

# Clean Agency IDs and Mission Areas from Tanium data
for col in id_columns:
    df_t[col] = df_t[col].str.replace('AgencyID-', '')
    df_t.loc[df_t[col].str.startswith('MissionArea', na=True), col] = ''

# Keep only relevant columns and group by workstation
t_cols = ['Encrypted Workstation Name', 'Operating System']
df_t = df_t[t_cols + id_columns].groupby('Encrypted Workstation Name').first()

s_cols = ['Encrypted Workstation Name', 'Agency']
df_s = df_s[s_cols].groupby('Encrypted Workstation Name').first()

# Merge and only keep workstations in both datasets
print('Merging Tanium and SCCM data')
df_inner = df_t.merge(df_s, how='inner', on='Encrypted Workstation Name')

# 'Matching' indicates SCCM Agency ID matches all Tanium Agency IDs
print('Comparing Tanium and SCCM Agency ID classifications')
df_inner['Matching'] = True
for col in id_columns:
    # If Tanium entry is non-empty and not equal to SCCM entry, set False
    df_inner.loc[
        (df_inner[col].str.isalpha()) & (df_inner[col] != df_inner['Agency']),
        'Matching'
        ] = False

# Rearrange SCCM Agency ID to front
df_inner.insert(1, 'SCCM Agency ID', df_inner.pop('Agency'))

# Fill all NaN SCCM Agency IDs with 'None', for grouping
df_inner['SCCM Agency ID'].fillna('None', inplace=True)

# Alphabetize and concatenate all Tanium Agency IDs
df_inner[id_columns] = np.sort(df_inner[id_columns], axis=1)
df_inner['Tanium Agency IDs'] = \
    df_inner[id_columns].agg(lambda row: '-'.join(filter(None, row)), axis=1)
df_inner.drop(columns=id_columns, inplace=True)

# REPORT 1: SCCM and all Tanium classifications match
print('Exporting matching classification report (1/7)')
df_inner.loc[df_inner['Matching']].drop(columns='Matching') \
    .to_excel('./data/matching_raw.xlsx')

# REPORT 2: At least one classification in Tanium does not match SCCM
print('Exporting mismatching classification report (2/7)')
df_mismatch = df_inner.loc[~df_inner['Matching']].drop(columns='Matching')
df_mismatch.to_excel('./data/mismatching_raw.xlsx')

# REPORT 3: Mismatches grouped by SCCM Agency ID classification
print('Exporting SCCM-grouped mismatching classification report (3/7)')
df_mismatch.groupby(['SCCM Agency ID', 'Tanium Agency IDs']).count() \
    .rename(columns={'Operating System': 'Count'}) \
    .to_excel('./data/mismatching_sccm_grouped.xlsx')

# REPORT 4: Mismatches grouped by Tanium Agency ID classifications
print('Exporting Tanium-grouped mismatching classification report (4/7)')
df_mismatch.groupby(['Tanium Agency IDs', 'SCCM Agency ID']).count() \
    .rename(columns={'Operating System': 'Count'}) \
    .to_excel('./data/mismatching_tanium_grouped.xlsx')

df_outer = df_t.merge(
    df_s, how='outer', on='Encrypted Workstation Name', indicator=True
    )

# REPORT 5: Only Tanium Report (Workstation Name is only in Tanium)
print('Exporting Tanium-only workstations report (5/7)')
df_outer[df_outer['_merge'] == 'left_only'].drop(columns='_merge') \
    .to_excel('./data/tanium_unique.xlsx')

# REPORT 6: Only SCCM Report (Workstation Name is only in SCCM)
print('Exporting SCCM-only workstations report (6/7)')
df_outer[df_outer['_merge'] == 'right_only'].drop(columns='_merge') \
    .to_excel('./data/sccm_unique.xlsx')

# REPORT 7: Agency workstation reporting for SCCM and Tanium
print('Generating Tanium and SCCM reporting statistics by agency')
df_stats = pd.DataFrame(columns=['Agency ID', 
                                 'Total Workstations',
                                 'SCCM Workstations',
                                 'SCCM Workstations Proportion',
                                 'Tanium Workstations',
                                 'Tanium Workstations Proportion',
                                 'Shared Workstations',
                                 'Shared Workstations Proportion'])

# Find all unique Agency IDs
all_id_cols = id_columns + ['Agency']
unique_ids = set()
for col in all_id_cols:
    unique_ids.update(df_outer[col].dropna().unique())

# Genegerate statistics DataFrame
for i, id in enumerate(unique_ids):
    row = i + 1

    # Indicate if each workstation belongs to agency 'id'
    df_outer['id_indicator'] = False
    for col in all_id_cols:
        df_outer.loc[df_outer[col] == id, 'id_indicator'] = True

    # Indicate if Tanium workstations belong to agency 'id'
    df_t['id_indicator'] = False
    for col in id_columns:
        df_t.loc[df_t[col] == id, 'id_indicator'] = True

    # Calculate statistics for final table
    df_stats.at[row, 'Agency ID'] = id

    work_all = df_outer['id_indicator'].values.sum()
    df_stats.at[row, 'Total Workstations'] = work_all

    work_s = (df_s['Agency'] == id).values.sum()
    df_stats.at[row, 'SCCM Workstations'] = work_s
    df_stats.at[row, 'SCCM Workstations Proportion'] = work_s / work_all

    work_t = df_t['id_indicator'].values.sum()
    df_stats.at[row, 'Tanium Workstations'] = work_t
    df_stats.at[row, 'Tanium Workstations Proportion'] = work_t / work_all

    work_both = work_s + work_t - work_all
    df_stats.at[row, 'Shared Workstations'] = work_both
    df_stats.at[row, 'Shared Workstations Proportion'] = work_both / work_all

# Sort Agency IDs alphabetically
df_stats.sort_values(by=['Agency ID'], inplace=True)

print('Exporting workstation reporting statistics (7/7)')
df_stats.to_excel('./data/workstation_statistics.xlsx', index=False)
