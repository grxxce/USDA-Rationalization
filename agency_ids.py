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
df_joint = df_t.merge(df_s, how='inner', on='Encrypted Workstation Name')

# 'Matching' indicates SCCM Agency ID matches all Tanium Agency IDs
print('Comparing Tanium and SCCM Agency ID classifications')
df_joint['Matching'] = True
for col in id_columns:
    # If Tanium entry is non-empty and not equal to SCCM entry, set False
    df_joint.loc[
        (df_joint[col].str.isalpha()) & (df_joint[col] != df_joint['Agency']),
        'Matching'
        ] = False

# Rearrange SCCM Agency ID to front
df_joint.insert(1, 'SCCM Agency ID', df_joint.pop('Agency'))

# Fill all NaN SCCM Agency IDs with 'None', for grouping
df_joint['SCCM Agency ID'].fillna('None', inplace=True)

# Alphabetize and concatenate all Tanium Agency IDs
df_joint[id_columns] = np.sort(df_joint[id_columns], axis=1)
df_joint['Tanium Agency IDs'] = \
    df_joint[id_columns].agg(lambda row: '-'.join(filter(None, row)), axis=1)
df_joint.drop(columns=id_columns, inplace=True)

# REPORT 1: SCCM and all Tanium classifications match
print('Exporting matching report to Excel')
df_joint.loc[df_joint['Matching']].drop(columns='Matching') \
    .to_excel('./data/matching_raw.xlsx')

# REPORT 2: At least one classification in Tanium does not match SCCM
print('Exporting mismatching report to Excel')
df_mismatch = df_joint.loc[~df_joint['Matching']].drop(columns='Matching')
df_mismatch.to_excel('./data/mismatching_raw.xlsx')

# REPORT 3: Mismatches grouped by SCCM Agency ID classification
print('Exporting SCCM-grouped mismatching report to Excel')
df_mismatch.groupby(['SCCM Agency ID', 'Tanium Agency IDs']).count() \
    .rename(columns={'Operating System': 'Count'}) \
    .to_excel('./data/mismatching_sccm_grouped.xlsx')

# REPORT 4: Mismatches grouped by Tanium Agency ID classifications
print('Exporting Tanium-grouped mismatching report to Excel')
df_mismatch.groupby(['Tanium Agency IDs', 'SCCM Agency ID']).count() \
    .rename(columns={'Operating System': 'Count'}) \
    .to_excel('./data/mismatching_tanium_grouped.xlsx')

# REPORT 5: Only Tanium Report (Workstation Name is only in Tanium)
# Schema: Workstation Name, Operating System, Tanium Agency ID 1, …

# REPORT 6: Only SCCM Report (Workstation Name is only in SCCM)
# Schema: Workstation Name, OS, OS Version, SCCM Agency ID
