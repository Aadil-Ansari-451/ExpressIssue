import pandas as pd
import os

def process_excel_data(filter_type, raw_folder, output_folder, input_file, sheet_name):
    """
    Process Excel data and generate CSV with trailing commas
    
    Args:
        filter_type (str): "Charge" or "Credit" to filter by
        raw_folder (str): Path to raw data folder
        output_folder (str): Path to output folder
        input_file (str): Excel file name
        sheet_name (str): Sheet name in Excel file
    """
    # Construct full file path
    excel_path = os.path.join(raw_folder, input_file)
    
    # Read Excel file
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    
    # Filter rows by type
    df_filtered = df[df[0] == filter_type]
    
    if df_filtered.empty:
        print(f"No {filter_type} records found in the Excel file")
        return
    
    # Group data by ID (column B) and collect values from columns C and D
    grouped_data = {}
    for _, row in df_filtered.iterrows():
        id_val = row[1]  # ID in column B
        values = row[2:4].tolist()  # Values from columns C and D
        
        if id_val not in grouped_data:
            grouped_data[id_val] = values
        else:
            grouped_data[id_val].extend(values)
    
    # Convert to DataFrame
    result_df = pd.DataFrame.from_dict(grouped_data, orient="index")
    result_df.reset_index(inplace=True)
    result_df.rename(columns={"index": "ID"}, inplace=True)
    
    # Generate output file path
    output_file = os.path.join(output_folder, "input.csv")
    
    # Ensure output directory exists
    os.makedirs(output_folder, exist_ok=True)
    
    # Write CSV with trailing commas in one step
    with open(output_file, 'w', newline='') as f:
        for _, row in result_df.iterrows():
            # Convert row to string and add trailing comma
            row_str = ','.join(str(val) for val in row.values) + ','
            f.write(row_str + '\n')
    
    print(f"Processed CSV saved at: {output_file}")
    print(f"Total {filter_type} records processed: {len(result_df)}")

if __name__ == "__main__":
    # Default configuration (will be overridden by Flask app)
    raw_folder = r"C:\path\to\raw\data"
    output_folder = r"C:\path\to\processed\data"
    input_file = "TestCon.xlsx"
    sheet_name = "Sheet1"
    filter_type = "Credit"  # Default filter type
    
    process_excel_data(filter_type, raw_folder, output_folder, input_file, sheet_name)