from flask import Flask, render_template, jsonify, request, send_from_directory
import subprocess
import sys
import os
import traceback
import json
from datetime import datetime

app = Flask(__name__, static_folder='static')

# Load configuration from file
def load_config():
    try:
        with open('config.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        # Default configuration if file doesn't exist
        default_config = {
            'raw_folder': r"C:\path\to\raw\data",
            'output_folder': r"C:\path\to\processed\data",
            'input_file': "TestCon.xlsx",
            'sheet_name': "Sheet1"
        }
        save_config(default_config)
        return default_config

def save_config(config):
    with open('config.json', 'w') as f:
        json.dump(config, f, indent=4)

CONFIG = load_config()

@app.route('/')
def index():
    """Main page with the Charge button"""
    return render_template('index.html')

@app.route('/execute_charge', methods=['POST'])
def execute_charge():
    """Execute the charge processing script"""
    try:
        # Update the FileCon.py script with current config
        update_script_config('Charge')
        
        # Execute the script
        result = subprocess.run([sys.executable, 'FileCon.py'], 
                              capture_output=True, text=True, timeout=60)
        
        if result.returncode == 0:
            return jsonify({
                'success': True,
                'message': f'Charge processing completed successfully! CSV saved to {CONFIG["output_folder"]}',
                'output': result.stdout
            })
        else:
            return jsonify({
                'success': False,
                'message': 'Error occurred during processing',
                'error': result.stderr
            })
            
    except subprocess.TimeoutExpired:
        return jsonify({
            'success': False,
            'message': 'Processing timed out. Please try again.'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Unexpected error: {str(e)}',
            'error': traceback.format_exc()
        })

@app.route('/execute_credit', methods=['POST'])
def execute_credit():
    """Execute the credit processing script"""
    try:
        # Update the FileCon.py script with current config
        update_script_config('Credit')
        
        # Execute the script
        result = subprocess.run([sys.executable, 'FileCon.py'], 
                              capture_output=True, text=True, timeout=60)
        
        if result.returncode == 0:
            return jsonify({
                'success': True,
                'message': f'Credit processing completed successfully! CSV saved to {CONFIG["output_folder"]}',
                'output': result.stdout
            })
        else:
            return jsonify({
                'success': False,
                'message': 'Error occurred during processing',
                'error': result.stderr
            })
            
    except subprocess.TimeoutExpired:
        return jsonify({
            'success': False,
            'message': 'Processing timed out. Please try again.'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Unexpected error: {str(e)}',
            'error': traceback.format_exc()
        })

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

def update_script_config(filter_type='Charge'):
    """Update the FileCon.py script with current configuration"""
    script_content = f'''import pandas as pd
import os

# Define folder paths
raw_folder = r"{CONFIG['raw_folder']}"
output_folder = r"{CONFIG['output_folder']}"

# Define file and sheet
input_file = os.path.join(raw_folder, "{CONFIG['input_file']}")
sheet_name = "{CONFIG['sheet_name']}"  # specify the worksheet name

# Read the sheet
df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)

# Filter only rows where column A == "{filter_type}"
df_filtered = df[df[0] == "{filter_type}"]

# Append values based on common ID in column B (index 1)
grouped_data = {{}}
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
result_df.rename(columns={{"index": "ID"}}, inplace=True)

# Convert to CSV with trailing comma at the end of each row
output_file = os.path.join(output_folder, "input.csv")
result_df.to_csv(output_file, index=False, header=False)

# Add trailing comma manually
with open(output_file, "r") as f:
    lines = f.readlines()

lines = [line.strip() + "," + "\\n" for line in lines]  # append a comma at the end

with open(output_file, "w") as f:
    f.writelines(lines)

print(f"Processed CSV saved at: {{output_file}}")
'''
    
    with open('FileCon.py', 'w') as f:
        f.write(script_content)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
