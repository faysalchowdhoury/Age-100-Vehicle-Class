#!/usr/bin/env python
# coding: utf-8

# In[3]:


import pandas as pd
import re

def calculate_age(year):
    current_year = 2024
    return current_year - year

def extract_year(text):
    match = re.search(r'\b(19\d{2}|20\d{2})\b', text)
    if match:
        return int(match.group())
    return None

def transpose_rows_to_columns(input_file, output_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file, header=None)

    # Transpose rows to columns
    df_transposed = df.transpose()

    # Check if 'Number of Vehicles' is in the first row
    if 'Number of Vehicles' in df_transposed.iloc[0].values:
        # Set the first row as column headers
        df_transposed.columns = df_transposed.iloc[0]

        # Drop the first row (which is now the header)
        df_transposed = df_transposed.iloc[1:]

        # Filter rows where 'Number of Vehicles' is less than 100
        df_filtered = df_transposed[df_transposed['Number of Vehicles'].astype(float) >= 100]

        # Drop rows with empty values in the "Number of Vehicles" column
        df_filtered = df_filtered.dropna(subset=['Number of Vehicles'])

        # Set 'Serial Number' column as index temporarily to add the sequence
        df_filtered.set_index(pd.Index(range(1, len(df_filtered) + 1)), inplace=True)

        # Add 'Serial Number' column
        df_filtered.insert(0, 'Serial Number', range(1, len(df_filtered) + 1))

        # Reset the index to default numerical index
        df_filtered.reset_index(drop=True, inplace=True)

        # Extract year from 'Vehicle Class Name' column and calculate age
        df_filtered['Manufacturing Year'] = df_filtered['Vehicle Class Name'].apply(extract_year)
        df_filtered['Age'] = df_filtered['Manufacturing Year'].apply(calculate_age)

        # Print the transposed and filtered DataFrame to the console
        print(df_filtered)

        # Save the transposed and filtered DataFrame to an Excel file
        df_filtered.to_excel(output_file, index=False)
        print(f"Filtered data saved to {output_file}")
    else:
        print("Error: 'Number of Vehicles' column not found in the transposed DataFrame.")

# Replace 'Vehicle Class Template.xlsx' with your actual file path
input_excel_file = 'Vehicle Class Template.xlsx'
output_excel_file = 'Age 100 Vehicle Class.xlsx'

transpose_rows_to_columns(input_excel_file, output_excel_file)


# In[ ]:




