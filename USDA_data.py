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

# Create DataFrame grouped by 'Name' and 'Usage'
df_usage = df.groupby(['Name', 'Usage'], as_index=False).size()

# Export to '/results/aggregate_usage.xlsx' Excel file
df_usage.to_excel("./results/aggregate_usage.xlsx", index=False)
