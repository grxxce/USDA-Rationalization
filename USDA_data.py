import matplotlib.pyplot as plt
import os
import pandas as pd

print('Importing Tanium dataset')

# Read Tanium data as DataFrame df
df = pd.read_excel('tanium.xlsx')

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
    # 'Asset - Custom Tags.2.1',
    # 'Asset - Custom Tags.2.2.1',
    # 'Asset - Custom Tags.2.2.2.1',
    # 'Asset - Custom Tags.2.2.2.2.1',
    # 'Asset - Custom Tags.2.2.2.2.2.1',
    # 'Asset - Custom Tags.2.2.2.2.2.2.1',
    # 'Asset - Custom Tags.2.2.2.2.2.2.2.1',
    # 'Asset - Custom Tags.2.2.2.2.2.2.2.2.2.1'
]

# Save all unique Mission Areas and Agency IDs into Set tags
tags = set()
for column in id_columns:
    tags.update(df[column].dropna().unique())

# Make folders to store figures and data
os.makedirs('./data', exist_ok=True)
os.makedirs('./figures', exist_ok=True)

# Iterate through tags and create DataFrames and visualizations
for tag in tags:
    print('Processing tag: ' + tag)

    # 'tag_indicator' signals if a row has the given tag
    df['tag_indicator'] = False
    for column in id_columns:
        df.loc[df[column] == tag, 'tag_indicator'] = True

    # Group df by tags and usage
    df_tag = df.loc[df['tag_indicator']].groupby(['Name', 'Usage'])\
        .size().reset_index()
    df_tag.rename(columns={df_tag.columns[2]: 'Values'}, inplace=True)

    # Get usages and values for each software, required for visualizations
    softwares = list(df_tag['Name'].dropna().unique())
    usages, values = [], []
    for software in softwares:
        software_mask = df_tag['Name'] == software
        usages.append(df_tag.loc[software_mask, 'Usage'])
        values.append(df_tag.loc[software_mask, 'Values'])
    colors = {
        'High': 'green', 
        'Normal': 'yellow', 
        'Limited': 'orange', 
        'Usage not detected': 'red'
    }

    if len(softwares) > 0:
        # Generate figure with appropriate grid of subplots
        # Divides number of applications by 3 and rounds up
        numRows = len(softwares) // 3 + (len(softwares) % 3 > 0)
        fig, axs = plt.subplots(nrows=numRows, ncols=3,
                                squeeze=False, figsize=(8, numRows * 3))
        axs = axs.flat

        # Remove all extra subplots
        for ax in axs[len(softwares):]:
            ax.remove()

        # Set shared visualization features
        fig.suptitle(tag + ' usage by software', fontsize=12)
        fig.subplots_adjust(wspace=1, hspace=1)

        # Create bar charts and export as PNGs
        for i in range(len(softwares)):
            axs[i].set_title(softwares[i], fontsize=10)
            axs[i].bar(usages[i], values[i], color=[colors[k] for k in usages[i]])
            axs[i].tick_params(axis='x', labelrotation=30, labelsize=8)
            axs[i].set_xlabel('Usage', fontsize=8)
            axs[i].set_ylabel('Frequency', fontsize=8)
        fig.savefig('./figures/' + tag + '_bar.png')

        # Create pie charts and export as PNGs
        for i in range(len(softwares)):
            axs[i].clear()
            axs[i].set_title(softwares[i], fontsize=10)
            axs[i].pie(values[i], colors=[colors[k] for k in usages[i]])
            axs[i].axis('Equal')
        fig.savefig('./figures/' + tag + '_pie.png')
        plt.close()
    
    # Pivot the DataFrame to display the software usage by usage levels
    df_tag = df_tag.pivot(index='Name', columns='Usage', values='Values')
    new_index = ['Usage not detected', 'Limited', 'Normal', 'High']
    df_tag = df_tag.reindex(columns=new_index)
    df_tag.to_excel('./data/' + tag + '_usage.xlsx')
