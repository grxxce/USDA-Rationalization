import os
import pandas as pd

# Data will be outputed to ./results directory
os.makedirs('./results', exist_ok=True)

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

# Iterate through tags
for tag in tags:
    # 'tag_indicator' signals if the data has the given tag
    df['tag_indicator'] = False
    for column in id_columns:
        df.loc[df[column] == tag, 'tag_indicator'] = True
    # Create DataFrame grouped by the 'Name' and 'Usage' for the given tag
    df_usage = df.loc[df['tag_indicator']].groupby(['Name', 'Usage'], as_index=False).size()
    # Export to '/results/TAG_usage.xlsx' Excel file
    df_usage.to_excel("./results/" + tag + "_usage.xlsx", index=False)
