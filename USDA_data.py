import matplotlib.pyplot as plt
import pandas as pd
import os

# Read input data into dataframe df
df = pd.read_csv('USDA_data.csv', low_memory=False)

# Save all unique Mission Areas and Agency IDs into set tags
tags = set()
# Indicate all columns containing tags and save unique into a set
IDColumns = ['Asset - Custom Tags - Copy.2',
             'Asset - Custom Tags - Copy.3',
             'Asset - Custom Tags - Copy.4',
             'Asset - Custom Tags - Copy.5'
             ]
for column in IDColumns:
    tags.update(df[column].dropna().unique())
# Make folders to store figures and data
os.makedirs('./data', exist_ok=True)
os.makedirs('./figures', exist_ok=True)

# Iterate through tags and use'tag_indicator' to signal data is within domain
for tag in tags:
    df['tag_indicator'] = False
    for column in IDColumns:
        df.loc[df[column] == tag, 'tag_indicator'] = True
    # Group df_usage tags by descending usage and export as a new excel
    df_usage = df.loc[df['tag_indicator'] == True].groupby(
        ['Name', 'Usage']).size().reset_index()
    # Pivot the dataframe to display the software usage by usage levels
    df_usage.rename(columns={df_usage.columns[2]: 'Values'}, inplace=True)
    df_usage = df_usage.pivot(index='Name', columns='Usage', values='Values')
    new_index = ['Baselining', 'Usage not detected', 'Normal', 'High']
    df_usage = df_usage.reindex(columns=new_index)
    df_usage.to_excel('./data/' + tag + '_usage.xlsx')
    # Create bar graphs and save as PNGs
    fig = df_usage.plot(kind='bar', rot=45, figsize=(20, 30))
    plt.close()
    fig.figure.savefig('./figures/' + tag + '_bar.png')
