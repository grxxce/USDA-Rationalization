# Find reports unique in Tanium and SCCM
import matplotlib.pyplot as plt
import pandas as pd
import os
os.makedirs('./data', exist_ok=True)

# Read Tanium input data as DataFrame df
df_t = pd.read_excel('Tanium_data.xlsx')
df_s = pd.read_excel('SCCM_data.xlsx')

# Clean data by removing baselining and duplicates
df_t = df_t[df_t['Usage'] != 'Baselining'].drop_duplicates()
df_t = df_t[df_t['Name'] != 'Adobe Reader and Acrobat Manager']

# Indicate all columns where AgencyID may be located for Tanium
idColumns = [
    'Asset - Custom Tags.2.1',
    'Asset - Custom Tags.2.2.1'
]

# Save all unique Tanium Agency IDs into set idT
idT_all = set()
idT = set()
for column in idColumns:
    idT_all.update(df_t[column].dropna().unique())
idT_all = set([id for id in list(idT_all) if 'AgencyID' in id])
for id in idT_all:
    idT.add(id.split('-')[1])

# Save all unique SCCM Agency IDs into set idS
idS = set()
idS.update(df_s['Agency'].dropna().unique())

# Store work stations in respective set
workT = set(df_t['Encrypted Workstation Name'])
workS = set(df_s['Encrypted Workstation Name'])

# Store shared and unique work stations
workShared = workT | workS
workUniqueT = workT - workShared
workUniqueS = workS - workShared

# REPORT 1: Create a shared work stations dataframe and export to excel
df_workShared = df_t.merge(
    df_s, on='Encrypted Workstation Name', how='inner', indicator=True)
df_workShared.to_excel('./data/shared_work_stations.xlsx')

# REPORT 2: Create a share work and ID dataframe adn export to to excel
df_workIdShared = pd.DataFrame()
for column in idColumns:
    df_workIdShared.append(df_workShared.loc[(
        df_workShared['Agency'] == 'AgencyID-' + df_workShared[column])])
df_workIdShared.to_excel('./data/shared_work_and_IDs.xlsx')
