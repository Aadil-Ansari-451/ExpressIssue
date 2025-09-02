# Charge Processing Tool

A web-based application for processing Excel data and generating CSV output. This tool features a Matrix-themed interface and can be embedded in SharePoint pages.

## Features

- **Matrix-Themed UI**: Dark interface with your Matrix image as background
- **Two Processing Options**: Charge and Credit buttons for different data types
- **Cross-Platform**: Works on Windows and Mac through web browsers
- **SharePoint Compatible**: Can be embedded in SharePoint pages
- **Simple Configuration**: Easy setup of input/output folders and files
- **Error Handling**: Comprehensive error messages and status updates

## What It Does

The application processes Excel files by:
1. Reading data from a specified Excel file and sheet
2. Filtering rows where column A contains "Charge"
3. Grouping data by ID (column B)
4. Generating a CSV file with trailing commas
5. Saving the output to a specified folder

## Setup Instructions

### Prerequisites

- Python 3.7 or higher
- Required Python packages (see requirements.txt)

### Installation

1. **Clone or download** this project to your server
2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure paths**:
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
   - Navigate to `http://localhost:5000`
   - You should see the Charge Processing Tool interface

## SharePoint Integration

### Option 1: Direct Embedding
1. Host the Flask application on a server accessible from your SharePoint environment
2. In SharePoint, add a "Script Editor" web part
3. Embed the application using an iframe:
   ```html
   <iframe src="http://your-server:5001" width="100%" height="600px" frameborder="0"></iframe>
   ```

### Option 2: SharePoint App
1. Package the application as a SharePoint app
2. Deploy to your SharePoint environment
3. Add the app to your SharePoint page

### Option 3: Web Part Development
1. Create a custom SharePoint web part
2. Integrate the application's API endpoints
3. Deploy the web part to SharePoint

## Configuration

### File Structure
```
Project_ExpressIssue/
├── app.py              # Flask web application
├── FileCon.py          # Original processing script
├── config.json         # Configuration settings
├── requirements.txt    # Python dependencies
├── templates/
│   └── index.html      # Matrix-themed web interface
└── README.md           # This file
```

### Configuration Options

- **raw_folder**: Directory containing input Excel files
- **output_folder**: Directory where processed CSV files will be saved
- **input_file**: Name of the Excel file to process
- **sheet_name**: Name of the worksheet within the Excel file

## Usage

1. **Configure Settings**: Click "Show/Hide Configuration" to set up paths and file names
2. **Click Charge**: Press the "Charge" button to execute the processing
3. **Monitor Status**: Watch for success/error messages
4. **Check Output**: Find the generated CSV file in your output folder

## Error Handling

The application handles various error scenarios:
- **File not found**: If the input Excel file doesn't exist
- **Invalid file format**: If the Excel file is corrupted or in wrong format
- **Permission errors**: If the application can't read/write to specified folders
- **Processing timeouts**: If the script takes too long to execute

## Security Considerations

- **Network Security**: Ensure the server is properly secured
- **File Permissions**: Configure appropriate read/write permissions
- **Input Validation**: Validate file paths and names
- **Error Logging**: Monitor application logs for security issues

## Troubleshooting

### Common Issues

1. **"Module not found" errors**:
   - Ensure all dependencies are installed: `pip install -r requirements.txt`

2. **Permission denied errors**:
   - Check file and folder permissions
   - Ensure the application has read/write access

3. **Configuration not saving**:
   - Verify the `config.json` file is writable
   - Check for disk space issues

4. **SharePoint embedding issues**:
   - Ensure the server is accessible from SharePoint
   - Check iframe security policies
   - Verify CORS settings if needed

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
