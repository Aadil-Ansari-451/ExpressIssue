from flask import Flask, render_template, jsonify, request
import subprocess
import sys
import os
import json

app = Flask(__name__, static_folder='static')

def load_config():
    """Load configuration from JSON file"""
    try:
        with open('config.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        # Default configuration
        default_config = {
            'raw_folder': r"C:\path\to\raw\data",
            'output_folder': r"C:\path\to\processed\data",
            'input_file': "TestCon.xlsx",
            'sheet_name': "Sheet1"
        }
        save_config(default_config)
        return default_config

def save_config(config):
    """Save configuration to JSON file"""
    with open('config.json', 'w') as f:
        json.dump(config, f, indent=4)

CONFIG = load_config()

@app.route('/')
def index():
    """Main page with Charge and Credit buttons"""
    return render_template('index.html')

def execute_processing(filter_type):
    """Execute the Excel processing script"""
    try:
        # Update the FileCon.py script with current config
        update_script_config(filter_type)
        
        # Execute the script
        result = subprocess.run([sys.executable, 'FileCon.py'], 
                              capture_output=True, text=True, timeout=60)
        
        if result.returncode == 0:
            return {
                'success': True,
                'message': f'{filter_type} processing completed successfully! CSV saved to {CONFIG["output_folder"]}',
                'output': result.stdout
            }
        else:
            return {
                'success': False,
                'message': 'Error occurred during processing',
                'error': result.stderr
            }
            
    except subprocess.TimeoutExpired:
        return {
            'success': False,
            'message': 'Processing timed out. Please try again.'
        }
    except Exception as e:
        return {
            'success': False,
            'message': f'Unexpected error: {str(e)}'
        }

@app.route('/execute_charge', methods=['POST'])
def execute_charge():
    """Execute the charge processing script"""
    result = execute_processing('Charge')
    return jsonify(result)

@app.route('/execute_credit', methods=['POST'])
def execute_credit():
    """Execute the credit processing script"""
    result = execute_processing('Credit')
    return jsonify(result)

@app.route('/update_config', methods=['POST'])
def update_config():
    """Update configuration settings"""
    try:
        data = request.get_json()
        CONFIG.update(data)
        save_config(CONFIG)
        return jsonify({'success': True, 'message': 'Configuration updated successfully'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error updating config: {str(e)}'})

def update_script_config(filter_type):
    """Update the FileCon.py script with current configuration"""
    script_content = f'''import pandas as pd
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
        print(f"No {{filter_type}} records found in the Excel file")
        return
    
    # Group data by ID (column B) and collect values from columns C and D
    grouped_data = {{}}
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
    result_df.rename(columns={{"index": "ID"}}, inplace=True)
    
    # Generate output file path
    output_file = os.path.join(output_folder, "input.csv")
    
    # Ensure output directory exists
    os.makedirs(output_folder, exist_ok=True)
    
    # Write CSV with trailing commas in one step
    with open(output_file, 'w', newline='') as f:
        for _, row in result_df.iterrows():
            # Convert row to string and add trailing comma
            row_str = ','.join(str(val) for val in row.values) + ','
            f.write(row_str + '\\n')
    
    print(f"Processed CSV saved at: {{output_file}}")
    print(f"Total {{filter_type}} records processed: {{len(result_df)}}")

if __name__ == "__main__":
    # Configuration from Flask app
    raw_folder = r"{CONFIG['raw_folder']}"
    output_folder = r"{CONFIG['output_folder']}"
    input_file = "{CONFIG['input_file']}"
    sheet_name = "{CONFIG['sheet_name']}"
    filter_type = "{filter_type}"
    
    process_excel_data(filter_type, raw_folder, output_folder, input_file, sheet_name)
'''
    
    with open('FileCon.py', 'w') as f:
        f.write(script_content)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
