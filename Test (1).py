#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd

# Load the Excel file into a DataFrame
data_path = "./data.xlsx"
df = pd.read_excel(data_path, sheet_name='Sheet1')


df['Column Name C'] = df['Column Name B'].str.upper().str.replace(" ", "_")


column_counts = df['Column Name C'].value_counts()

# Set a threshold for highlighting the rows (rows with counts greater than or equal to this threshold will be highlighted)
threshold = 1  # Change this value to your desired threshold

# Define a function to highlight rows based on the threshold
def highlight_rows(row):
    name = row['Column Name C']
    count = column_counts.get(name, 0)
    return ['background-color: lightgray' if count <= threshold else '' for _ in row]

# Apply the conditional formatting to the DataFrame by row
highlighted_df = df.style.apply(highlight_rows, axis=1)


unique_df = df[['Column Name', 'Column Name C']].drop_duplicates(subset='Column Name', keep='first')

# Reapply the row highlights to the unique DataFrame
highlighted_unique_df = unique_df.style.apply(highlight_rows, axis=1)

