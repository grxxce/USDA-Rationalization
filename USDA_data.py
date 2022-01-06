import matplotlib.pyplot as plt
import pandas as pd
import os
 
# Read input data into dataframe df
df = pd.read_csv('USDA_data.csv', low_memory=False)

# Part 2: Save all unique Mission Areas and Agency IDs into set tags
tags = set()
# Indicate all columns where tag data may be located and save unique tags into a set 'tags'
IDColumns = ['Asset - Custom Tags - Copy.2', 'Asset - Custom Tags - Copy.3', 'Asset - Custom Tags - Copy.4', 'Asset - Custom Tags - Copy.5']
for column in IDColumns:
    tags.update(df[column].dropna().unique())
# Make folders to store figures and data
os.mkdir("./data")
os.mkdir("./figures")

# Part 3: Iterate through tags and create a 'tag_indicator' to signal if the data is within the tag's domain
for tag in tags:
    df['tag_indicator'] = False
    for column in IDColumns:
        df.loc[df[column] == tag, 'tag_indicator'] = True
    # Create a df_usage dataframe grouped by the 'Name' and 'Usage' of every 'tag' and export as a new 'tag_usage' excel
    df_usage = df.loc[df['tag_indicator']==True].groupby(['Name', 'Usage']).size().reset_index()
    # Pivot the dataframe to display the software usage by usage levels
    df_usage.rename(columns={df_usage.columns[2]: "Values" }, inplace = True)
    df_usage = df_usage.pivot(index='Name', columns='Usage', values="Values")
    df_usage.to_excel("./data/" + tag + "_usage.xlsx")
    # Create bar grahs and save as PNGs
    fig = df_usage.plot(kind="bar")
    fig.figure.savefig("./figures/" + tag + "_image.png")
