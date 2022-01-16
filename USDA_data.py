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

# Iterate through tags and create new df for every region
for tag in tags:
    # 'tag_indicator' signals if the data has the given tag
    df['tag_indicator'] = False
    for column in id_columns:
        df.loc[df[column] == tag, 'tag_indicator'] = True

    # Group df by tags by descending usage and export as a new excel
    df_tag = df.loc[df['tag_indicator']].groupby(['Name', 'Usage'])\
        .size().reset_index()
    df_tag.rename(columns={df_tag.columns[2]: 'Values'}, inplace=True)
    # TO DO: Reindex 'Usage' to ascending order

    # Set the right axes
    softwares = set()
    softwares.update(df_tag['Name'].dropna().unique())
    softwares = list(softwares)
    colors = {'High': 'g', 'Normal': 'y',
              'Limited': 'r', 'Usage not detected': 'k'}

    # Generate the correct number of subplots
    numRows = len(softwares) // 3
    if len(softwares) < 3 and len(softwares) > 0:
        numRows = 1
    fig, ax = plt.subplots(nrows=numRows, ncols=3,
                           squeeze=False, figsize=(8, ((len(softwares) // 3) + 1) * 3))

    if len(softwares) != 0:
        # Create bar graphs and export as pngs
        num = 0
        i = 0
        while (i < (len(softwares) // 3) or (i == 0)):
            j = 0
            while j < 3 and num < len(softwares):
                usage = df_tag.loc[df_tag['Name'] == softwares[num], 'Usage']
                value = df_tag.loc[df_tag['Name'] == softwares[num], 'Values']
                ax[i, j].bar(usage, value, color=[colors[k] for k in usage])
                ax[i, j].set_title(softwares[num], fontsize=10)
                ax[i, j].tick_params(axis='x', labelrotation=30, labelsize=8)
                ax[i, j].set_xlabel('Usage', fontsize=8)
                ax[i, j].set_ylabel('Frequency', fontsize=8)
                num += 1
                j += 1
            i += 1
        fig.suptitle(tag + ' usage by software', fontsize=12)
        plt.tight_layout()
        plt.subplots_adjust(wspace=0.5, hspace=2)
        plt.savefig('./figures/' + tag + '_bar.png')
        plt.close()

        # Create pie charts and export as pngs
        num = 0
        i = 0
        while (i < (len(softwares) // 3) or (i == 0)):
            j = 0
            while j < 3 and num < len(softwares):
                usage = df_tag.loc[df_tag['Name'] == softwares[num], 'Usage']
                value = df_tag.loc[df_tag['Name'] == softwares[num], 'Values']
                ax[i, j].set_title(softwares[num], fontsize=10)
                ax[i, j].set_xlabel('Usage', fontsize=8)
                num += 1
                j += 1
            i += 1
        # Label each plot accordingly
        fig.suptitle(tag + ' usage by software', fontsize=12)
        plt.tight_layout()
        plt.subplots_adjust(wspace=0.5, hspace=2)
        plt.savefig('./figures/' + tag + '_pie.png')
        plt.close()

    # Pivot the dataframe to display the software usage by usage levels
    df_tag = df_tag.pivot(index='Name', columns='Usage', values='Values')
    new_index = ['Usage not detected', 'Limited', 'Normal', 'High']
    df_tag = df_tag.reindex(columns=new_index)
    df_tag.to_excel('./data/' + tag + '_usage.xlsx')
