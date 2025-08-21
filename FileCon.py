import pandas as pd
import numpy as np

#Reading 
df = pd.read_excel("TestCon.xlsx", header=None)

#Appending based on common column ID in Col#1
grouped_data = {}
for _, row in df.iterrows():
    id_val = row[0]
    values = row[1:3].tolist()

    if id_val not in grouped_data:
        grouped_data[id_val] = values
    else:
        grouped_data[id_val].extend(values)

# Convert it back to a DataFrame
result_df = pd.DataFrame.from_dict(grouped_data, orient="index")
result_df.reset_index(inplace=True)

# Write the result into the Excel file
result_df.to_excel("input.xlsx", index=False)