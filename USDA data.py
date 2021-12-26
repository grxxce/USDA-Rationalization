import pandas as pd
df = pd.read_csv('USDA_data.csv')

# Part 2: Creating columns for each Agency and Mission Area
# Create a column called AgencyID and insert the Agency ID and repeat for columns L, M, N, and O
df['AgencyID'] = ""
IDColumns = ['Asset - Custom Tags - Copy.2', 'Asset - Custom Tags - Copy.3', 'Asset - Custom Tags - Copy.4', 'Asset - Custom Tags - Copy.5']
for column in IDColumns:
    df['AgencyID'][df[column].str.contains("AgencyID", na=False)] = df[column].str.split('-').str[1]

# Create a column called Mission Area and insert the correct mission area for columns L, M, N, and O
df['MissionArea'] = ""
areaColumns = ['Asset - Custom Tags - Copy.2', 'Asset - Custom Tags - Copy.3', 'Asset - Custom Tags - Copy.4', 'Asset - Custom Tags - Copy.5']
for column in areaColumns:
    df['MissionArea'][df[column].str.contains("MissionArea", na=False)] = df[column].str.split('-').str[1]
## print(df[['AgencyID', 'MissionArea', 'Asset - Custom Tags - Copy.2', 'Asset - Custom Tags - Copy.3', 'Asset - Custom Tags - Copy.4', 'Asset - Custom Tags - Copy.5']])

# Create a new column for every AgencyID, where value is True when AgencyID matches
uAgencyID = []
for index, row in df.iterrows():
    if row['AgencyID'] not in uAgencyID:
        uAgencyID.append(row['AgencyID'])
        df[row['AgencyID']] = False
    df[row['AgencyID']][index] = True  
## print(uAgencyID)
## print(df[['AgencyID', 'NRCS', 'Asset - Custom Tags - Copy.2', 'Asset - Custom Tags - Copy.3', 'Asset - Custom Tags - Copy.4', 'Asset - Custom Tags - Copy.5']])

# Create a new column for every MissionArea, where value is True when MissionArea matches
uMissionArea = []
for index, row in df.iterrows():
    if row['MissionArea'] not in uMissionArea:
        uMissionArea.append(row['MissionArea'])
        df[row['MissionArea']] = False
    df[row['MissionArea']][index] = True  
## print(uMissionArea)
## print(df[['MissionArea', 'FPAC', 'Asset - Custom Tags - Copy.2', 'Asset - Custom Tags - Copy.3', 'Asset - Custom Tags - Copy.4', 'Asset - Custom Tags - Copy.5']])

# Part 3: Create new dataframes by Agency and Mission Area
# Filtering by AgencyID, count the number of occurances for each software at each usage level
df2 = {}
for agency in uAgencyID:
    df2[agency]=df.loc[df['AgencyID']==agency].groupby(['Name', 'Usage']).size()
    ## print("This is a new Agency: {}".format(agency))
    ## print(df2[agency])

# Filtering by MissionArea, count the number of occurances for each software at each usage level
df3 = {}
for area in uMissionArea:
    df3[area]=df.loc[df['MissionArea']==area].groupby(['Name', 'Usage']).size()
#     print("This is a new Mission Area: {}".format(area))
#     print(df3[area])