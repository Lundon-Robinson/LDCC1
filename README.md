# LDCC1 Data Processor

A comprehensive, GUI-based automation tool for processing client cash management data with seamless EQ online banking integration.

## Features

- **Professional GUI Interface**: User-friendly interface with file selection, progress tracking, and comprehensive logging
- **CSV/Excel File Support**: Process both CSV and Excel files with automatic format detection
- **Payment Processing Option**: Checkbox to enable payment processing workflow
- **EQ Online Integration**: Processes data up to the point of EQ online banking integration
- **Comprehensive Logging**: Detailed logging with GUI display and file export options
- **Data Validation**: Robust input validation and error handling
- **Report Generation**: Automatic generation of processing summaries and reports
- **Ready-to-Run**: Fully polished script ready for immediate deployment

## Quick Start

### Prerequisites

- Python 3.6 or higher
- Windows, macOS, or Linux operating system

### Installation

1. **Clone the repository** (if not already done):
   ```bash
   git clone https://github.com/Lundon-Robinson/LDCC1.git
   cd LDCC1
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**:
   ```bash
   python ldcc1_processor.py
   ```

## Usage Guide

### Starting the Application

1. Double-click `ldcc1_processor.py` or run from command line:
   ```bash
   python ldcc1_processor.py
   ```

2. The GUI will open with the following interface:

### Using the Interface

#### 1. File Selection
- Click **"Browse..."** to select your CSV or Excel file
- Supported formats: `.csv`, `.xlsx`, `.xls`
- The selected file path will appear in the text field

#### 2. Payment Processing Option
- **Check the box** if you want to process payments
- **When checked**: The script will process all data and prepare it for EQ online banking, then stop
- **When unchecked**: The script will process data without payment preparation

#### 3. Processing
- Click **"Start Processing"** to begin
- Monitor progress via the progress bar and status updates
- View detailed logs in the scrollable log area

#### 4. Results
- **With Payments**: Script stops before EQ online - you'll get a notification to proceed to EQ banking
- **Without Payments**: Complete processing with summary report
- All results are saved to `reports/` and `payment_output/` directories

### Output Files

The script creates several output directories:

```
LDCC1/
├── logs/                    # Processing logs
├── reports/                 # Summary reports (JSON format)
├── payment_output/          # Payment files for EQ online (when payments enabled)
└── ldcc1_processor.py       # Main application
```

## Features in Detail

### Data Processing Workflow

1. **Data Loading**: Automatic detection and loading of CSV/Excel files
2. **Data Validation**: Comprehensive validation of data structure and content
3. **Benefits Processing**: Processing of benefits data and calculations
4. **Reconciliation**: Balance reconciliation and validation
5. **Payment Preparation**: (If selected) Preparation of payment data for EQ online
6. **Report Generation**: Creation of processing summaries and audit trails

### Error Handling

- Comprehensive error catching and logging
- User-friendly error messages
- Graceful handling of file format issues
- Data validation with clear feedback

### Logging System

- **GUI Logging**: Real-time log display in the application
- **File Logging**: Automatic saving of detailed logs with timestamps
- **Log Export**: Save current session logs to file
- **Log Clearing**: Clear display for new processing sessions

## Configuration

The script is designed to work out-of-the-box with minimal configuration. Key settings can be modified in the script if needed:

- **Log Level**: Modify the logging configuration in `setup_logging()`
- **Output Directories**: Change directory names in the processing functions
- **Data Validation**: Customize column detection in `validate_data_structure()`

## Troubleshooting

### Common Issues

1. **"Module not found" errors**:
   ```bash
   pip install -r requirements.txt
   ```

2. **File permission errors**:
   - Ensure the script has write permissions to create `logs/`, `reports/`, and `payment_output/` directories

3. **GUI not displaying**:
   - Ensure your system supports tkinter (usually included with Python)
   - On Linux: `sudo apt-get install python3-tk`

4. **Data loading errors**:
   - Verify file format (CSV or Excel)
   - Check file is not open in another application
   - Ensure file contains data

### Support

For additional support or questions:
1. Check the processing logs for detailed error information
2. Review the generated reports for processing summaries
3. Ensure input data follows expected format

## System Requirements

- **Operating System**: Windows 10+, macOS 10.12+, or Linux
- **Python**: 3.6 or higher
- **Memory**: 512MB RAM minimum (more for large datasets)
- **Storage**: 100MB free space for logs and reports

## Security and Privacy

- All processing is done locally on your machine
- No data is sent to external servers
- Logs contain processing information but can be cleared as needed
- Generated files are stored locally in the application directory

## Version History

- **v1.0.0**: Initial release with full GUI, CSV/Excel support, payment processing, and EQ online integration

---

**Ready to run straight away!** Simply install dependencies and execute the script.