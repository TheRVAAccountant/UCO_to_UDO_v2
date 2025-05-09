# UCO to UDO Reconciliation Tool - User Instructions

## Overview

The UCO to UDO Reconciliation Tool automates the process of comparing and reconciling Unfilled Customer Orders (UCO) and Undelivered Orders (UDO) financial data across multiple Excel spreadsheets. This tool helps financial analysts verify data consistency, identify discrepancies, and produce a consolidated reconciliation report.

## System Requirements

- Windows operating system
- Python 3.10 or higher
- Microsoft Excel installed (required for formula recalculation)
- Required Python packages (automatically installed during setup)

## Installation

### 1. Install Python

If you don't have Python installed:
1. Download Python 3.10 or higher from [python.org](https://www.python.org/downloads/)
2. During installation, check "Add Python to PATH"
3. Complete the installation following the on-screen instructions

### 2. Install the UCO to UDO Reconciliation Tool

1. Open Command Prompt (cmd) or PowerShell
2. Navigate to the UCO_to_UDO_v2 directory:
   ```
   cd path\to\UCO_to_UDO_v2
   ```
3. Create a virtual environment (recommended):
   ```
   python -m venv venv
   ```
4. Activate the virtual environment:
   ```
   venv\Scripts\activate
   ```
5. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Running the Tool

### Starting the Application

1. Ensure your virtual environment is activated:
   ```
   venv\Scripts\activate
   ```

2. Launch the application:
   ```
   python -m src.uco_to_udo_recon.main
   ```

3. The UCO to UDO Reconciliation Tool GUI will open.

### Using the Application

1. **Select Component Name**: Choose the relevant component (e.g., CBP, CG, CIS, etc.) from the dropdown menu.

2. **Select Input Files**:
   - **UCO to UDO Reconciliation File**: Click "Browse..." to select the main reconciliation file where results will be stored.
   - **Trial Balance File**: Click "Browse..." to select the trial balance Excel file.
   - **UCO to UDO TIER File**: Click "Browse..." to select the UCO to UDO TIER Excel file.

3. **Start Processing**: Click the "Start" button to begin the reconciliation process.

4. **Monitor Progress**:
   - The progress bar shows the current status of the operation.
   - The log window displays detailed information about each step.

5. **View Results**: When the processing is complete, the tool will:
   - Save the results to a new Excel file (original filename with "- DO" added)
   - Automatically open the resulting Excel file

### Understanding the Output

The tool creates a new Excel file with several key sheets:

1. **Original Sheets**: Preserved from the original UCO to UDO Reconciliation File.
2. **DO TB**: Added sheet containing trial balance data.
3. **DO UCO to UDO**: Added sheet containing UCO to UDO TIER data.

The tool adds the following elements to the reconciled file:

- **Tickmarks**: Special symbols that indicate whether values match or differ.
- **Yellow Highlighting**: Indicates important cells that have been verified.
- **DO Comments**: Column for any automated or manual comments about the reconciliation.

### Interpreting Tickmarks

- **✓** (checkmark symbol): Values match correctly.
- **X**: Values do not match, indicating a discrepancy.
- Other special characters may be used to indicate specific matching conditions.

## Troubleshooting

### Common Issues

1. **Excel File Access Errors**:
   - Ensure all Excel files are closed before running the tool.
   - If errors persist, restart Excel and try again.

2. **Component Sheet Not Found**:
   - Verify that the component name selected matches the sheet names in your Excel files.
   - Check that component sheets are formatted correctly with "UCO total reported in TIER" and "UDO total reported in TIER" cells.

3. **Missing Financial Data**:
   - Ensure Trial Balance file contains the component total sheet (e.g., "WMD Total").
   - Ensure TIER file contains the "UCO to UDO" sheet.

### Log Files

The application creates detailed log files in the "logs" directory. These logs are helpful for diagnosing issues and are named with the date and time of execution (e.g., "UCOtoUDORecon_Log_2025-05-09_10-40-46.txt").

If you encounter problems, please check these logs for error messages that can help identify the source of the issue.

## Contact and Support

For assistance with the UCO to UDO Reconciliation Tool, please contact your system administrator or financial systems support team.