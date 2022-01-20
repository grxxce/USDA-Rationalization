# Find reports unique in Tanium and SCCM
import matplotlib.pyplot as plt
import pandas as pd
import os

os.makedirs('./data', exist_ok=True)

# Import Tanium and SCCM input data as DataFrames
print('Importing Tanium and SCCM data')
df_t = pd.read_excel('Tanium_data.xlsx')
df_s = pd.read_excel('SCCM_data.xlsx')

print('Cleaning and preparing data')

# Indicate all columns where AgencyID may be located for Tanium
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

# Keep only relevant columns from Tanium and SCCM datasets
# Group by workstation so each workstation is present only once
t_cols = ['Encrypted Workstation Name', 'Operating System']
t_cols.extend(id_columns)
df_t = df_t[t_cols].groupby('Encrypted Workstation Name').first()

s_cols = ['Encrypted Workstation Name', 'Agency', 'OS']
df_s = df_s[s_cols].groupby('Encrypted Workstation Name').first()

# print('Exporting grouped data to Excel') # TODO: remove
# df_t.to_excel('./data/df_t_grouped.xlsx')
# df_s.to_excel('./data/df_s_grouped.xlsx')

# Merge DataFrames on workstation
print('Merging Tanium and SCCM data')
df_joint = df_t.merge(df_s, how='inner', on='Encrypted Workstation Name')
# df_joint.to_excel('./data/df_joint.xlsx')

# 'Matching' indicates SCCM classification matches all Tanium classifications 
print('Finding matching AgencyID classifications')
df_joint['Matching'] = True
for col in id_columns:
    df_joint.loc[
        (df_joint[col].str.isalpha()) &
        (df_joint[col] != df_joint['Agency']), 'Matching'
        ] = False

# REPORT 1: Matching Report (all classifications match)
print('Exporting matching report to Excel')
df_joint.loc[df_joint['Matching']] \
    .to_excel('./data/matching_workstations.xlsx')

# REPORT 2: Mismatch Report (at least one classification does not match)
print('Exporting mismatching report to Excel')
df_joint.loc[~df_joint['Matching']] \
    .to_excel('./data/mismatching_workstations.xlsx')

# REPORT 3: Only Tanium Report (Workstation Name is only in Tanium)
# Schema: Workstation Name, Operating System, Tanium Agency ID 1, …

# REPORT 4: Only SCCM Report (Workstation Name is only in SCCM)
# Schema: Workstation Name, OS, OS Version, SCCM Agency ID
