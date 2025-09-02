# Excel to CSV Processing Tool

A clean, efficient web-based application for processing Excel data and generating CSV output. This tool can be embedded in SharePoint pages and works across Windows and Mac platforms.

## Features

- **Clean, Professional UI**: Modern interface with customizable background
- **Dual Processing Options**: Charge and Credit buttons for different data types
- **Cross-Platform**: Works on Windows and Mac through web browsers
- **SharePoint Compatible**: Can be embedded in SharePoint pages
- **Efficient Processing**: Optimized code for faster execution
- **Error Handling**: Comprehensive error messages and status updates

## What It Does

The application processes Excel files by:
1. Reading data from a specified Excel file and sheet
2. Filtering rows where column A contains "Charge" or "Credit"
3. Grouping data by ID (column B)
4. Extracting values from columns C and D
5. Generating a CSV file with trailing commas
6. Saving the output to a specified folder

## Setup Instructions

### Prerequisites

- Python 3.7 or higher
- Required Python packages (see requirements.txt)

### Installation

1. **Clone or download** this project to your server
2. **Create virtual environment**:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure paths**:
   - Edit `config.json` or use the web interface
   - Update the following paths:
     - `raw_folder`: Path to your input Excel files
     - `output_folder`: Path where CSV files will be saved
     - `input_file`: Name of your Excel file (e.g., "TestCon.xlsx")
     - `sheet_name`: Name of the worksheet to process

### Running the Application

1. **Start the server**:
   ```bash
   python app.py
   ```

2. **Access the application**:
   - Open your web browser
   - Navigate to `http://localhost:5001`
   - You should see the Excel to CSV interface

## SharePoint Integration

### Direct Embedding
1. Host the Flask application on a server accessible from your SharePoint environment
2. In SharePoint, add a "Script Editor" web part
3. Embed the application using an iframe:
   ```html
   <iframe src="http://your-server:5001" width="100%" height="600px" frameborder="0"></iframe>
   ```

## Configuration

### File Structure
```
Project_ExpressIssue/
├── app.py              # Flask web application
├── FileCon.py          # Excel processing script
├── config.json         # Configuration settings
├── requirements.txt    # Python dependencies
├── templates/
│   └── index.html      # Web interface
├── static/             # Static files (images, etc.)
└── README.md           # This file
```

### Configuration Options

- **raw_folder**: Directory containing input Excel files
- **output_folder**: Directory where processed CSV files will be saved
- **input_file**: Name of the Excel file to process
- **sheet_name**: Name of the worksheet within the Excel file

## Usage

1. **Click Charge or Credit**: Press either button to execute the processing
2. **Monitor Status**: Watch for success/error messages
3. **Check Output**: Find the generated CSV file in your output folder

## Error Handling

The application handles various error scenarios:
- **File not found**: If the input Excel file doesn't exist
- **Invalid file format**: If the Excel file is corrupted or in wrong format
- **Permission errors**: If the application can't read/write to specified folders
- **Processing timeouts**: If the script takes too long to execute
- **Empty results**: If no matching records are found

## Code Optimization

### Recent Improvements

- **Eliminated redundant code**: Removed duplicate processing logic
- **Optimized CSV generation**: Single-step CSV writing with trailing commas
- **Reduced dependencies**: Only essential packages included
- **Improved error handling**: Better error messages and validation
- **Cleaner code structure**: Modular functions with clear documentation

### Performance Benefits

- **Faster execution**: Optimized data processing algorithms
- **Reduced memory usage**: Efficient data structures
- **Smaller package size**: Minimal dependencies
- **Better maintainability**: Clean, documented code

## Troubleshooting

### Common Issues

1. **"Module not found" errors**:
   - Ensure virtual environment is activated
   - Install dependencies: `pip install -r requirements.txt`

2. **Permission denied errors**:
   - Check file and folder permissions
   - Ensure the application has read/write access

3. **Configuration not saving**:
   - Verify the `config.json` file is writable
   - Check for disk space issues

4. **Port already in use**:
   - Change port in `app.py` if needed
   - Check for other running applications

### Logs and Debugging

- Check the Flask application logs for detailed error information
- Monitor the browser's developer console for JavaScript errors
- Verify network connectivity between SharePoint and the application server

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Review the application logs
3. Verify configuration settings
4. Test with a simple Excel file first

## License

This application is provided as-is for internal use.
