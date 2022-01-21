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

# Clear all tags without 'AgencyID-' prefix. Remove prefix
for col in id_columns:
    df_t.loc[~df_t[col].str.startswith('AgencyID-', na=False), col] = np.nan
    df_t[col] = df_t[col].str.replace('AgencyID-', '')

# Keep only relevant columns and group by workstation
t_cols = ['Encrypted Workstation Name', 'Operating System']
df_t = df_t[t_cols + id_columns].groupby('Encrypted Workstation Name').first()

s_cols = ['Encrypted Workstation Name', 'Agency']
df_s = df_s[s_cols].groupby('Encrypted Workstation Name').first()

# Merge and only keep workstations in both datasets
print('Merging Tanium and SCCM data')
df_inner = df_t.merge(df_s, how='inner', on='Encrypted Workstation Name')

# If SCCM does not indicate Agency ID, mark with 'None'
df_inner['Agency'].fillna('None', inplace=True)

# If Tanium does not indicate Agency ID, mark with 'None'
df_inner.loc[df_inner[id_columns].isnull().all(axis=1), id_columns[0]] = 'None'

# 'Matching' indicates SCCM Agency ID matches all Tanium Agency IDs
print('Comparing Tanium and SCCM Agency ID classifications')
df_inner['Matching'] = True
for col in id_columns:
    # If Tanium entry does not equal SCCM entry, set False
    df_inner.loc[
        df_inner[col].notna() & (df_inner[col] != df_inner['Agency']),
        'Matching'
        ] = False

# Rearrange SCCM Agency ID to front
df_inner.insert(1, 'SCCM Agency ID', df_inner.pop('Agency'))

# Alphabetize and concatenate all Tanium Agency IDs
def tanium_concat(df):
    df = df.copy()
    df.loc[:, id_columns] = np.sort(df[id_columns].fillna(''), axis=1)
    df.loc[:, 'Tanium Agency IDs'] = \
        df[id_columns].agg(lambda row: '-'.join(filter(None, row)), axis=1)
    return df.drop(columns=id_columns)
df_inner = tanium_concat(df_inner)

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

# REPORT 5: Workstation Name is only in Tanium dataset
print('Exporting Tanium-only workstations report (5/7)')
tanium_concat(df_outer[df_outer['_merge'] == 'left_only']) \
    .drop(columns=['_merge', 'Agency']) \
    .to_excel('./data/tanium_only.xlsx')

# REPORT 6: Workstation Name is only in SCCM dataset
print('Exporting SCCM-only workstations report (6/7)')
df_outer[df_outer['_merge'] == 'right_only'] \
    .drop(columns=id_columns + ['_merge', 'Operating System']) \
    .to_excel('./data/sccm_only.xlsx')

# REPORT 7: Agency workstation coverage for SCCM and Tanium
print('Generating Tanium and SCCM coverage statistics by agency')
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

# Generate coverage statistics DataFrame
for row, id in enumerate(unique_ids, 1):
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

print('Exporting workstation coverage statistics (7/7)')
df_stats.to_excel('./data/coverage_statistics.xlsx', index=False)
