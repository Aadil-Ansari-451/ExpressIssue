import pandas as pd
import os

# Define folder paths
raw_folder = r"C:\path\to\raw\data"
output_folder = r"C:\path\to\processed\data"

# Define file and sheet
input_file = os.path.join(raw_folder, "TestCon.xlsx")
sheet_name = "Sheet1"  # specify the worksheet name

# Read the sheet
df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)

# Filter only rows where column A == "Credit"
df_filtered = df[df[0] == "Credit"]

# Append values based on common ID in column B (index 1)
grouped_data = {}
for _, row in df_filtered.iterrows():
    id_val = row[1]  # assuming the ID is in column B
    values = row[2:4].tolist()  # columns C and D (adjust if needed)

    if id_val not in grouped_data:
        grouped_data[id_val] = values
    else:
        grouped_data[id_val].extend(values)

# Convert to DataFrame
result_df = pd.DataFrame.from_dict(grouped_data, orient="index")
result_df.reset_index(inplace=True)
result_df.rename(columns={"index": "ID"}, inplace=True)

# Convert to CSV with trailing comma at the end of each row
output_file = os.path.join(output_folder, "input.csv")
result_df.to_csv(output_file, index=False, header=False)

# Add trailing comma manually
with open(output_file, "r") as f:
    lines = f.readlines()

lines = [line.strip() + "," + "\n" for line in lines]  # append a comma at the end

with open(output_file, "w") as f:
    f.writelines(lines)

print(f"Processed CSV saved at: {output_file}")
