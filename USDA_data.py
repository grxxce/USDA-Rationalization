import matplotlib.pyplot as plt
import os
import pandas as pd

# Read input data as DataFrame df
df = pd.read_excel('USDA_data.xlsx')

# Clean data by removing baselining and duplicates
df = df[df['Usage'] != 'Baselining'].drop_duplicates()

# Ignore application that skews visualizations
df = df[df['Name'] != 'Adobe Reader and Acrobat Manager']

# Indicate all columns where tag data may be located
id_columns = [
    'Asset - Custom Tags - Copy.2',
    'Asset - Custom Tags - Copy.3',
    'Asset - Custom Tags - Copy.4',
    'Asset - Custom Tags - Copy.5'
]

# Save all unique Mission Areas and Agency IDs into Set tags
tags = set()
for column in id_columns:
    tags.update(df[column].dropna().unique())

# Make folders to store figures and data
os.makedirs('./data', exist_ok=True)
os.makedirs('./figures', exist_ok=True)

# Iterate through tags
for tag in tags:
    # 'tag_indicator' signals if the data has the given tag
    df['tag_indicator'] = False
    for column in id_columns:
        df.loc[df[column] == tag, 'tag_indicator'] = True
    # Group df_usage tags by descending usage and export as a new excel
    df_usage = df.loc[df['tag_indicator']].groupby(['Name', 'Usage'])\
        .size().reset_index()
    # Pivot the dataframe to display the software usage by usage levels
    df_usage.rename(columns={df_usage.columns[2]: 'Values'}, inplace=True)
    df_usage = df_usage.pivot(index='Name', columns='Usage', values='Values')
    new_index = ['Usage not detected', 'Limited', 'Normal', 'High']
    df_usage = df_usage.reindex(columns=new_index)
    df_usage.to_excel('./data/' + tag + '_usage.xlsx')
    # Create bar graphs and save as PNGs
    fig = df_usage.plot(kind='bar', rot=45, figsize=(
        20, 30), color=['b', 'r', 'y', 'g'])
    plt.close()
    fig.figure.savefig('./figures/' + tag + '_bar.png')
