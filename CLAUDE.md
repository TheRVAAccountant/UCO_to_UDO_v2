# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

UCO_to_UDO_v2 is a Python application that reconciles and compares data between multiple Excel spreadsheets for financial reconciliation of Undelivered Orders (UDO) and Unfilled Customer Orders (UCO) data across different government components. The application:

1. Copies specific sheets from source Excel files to a target file
2. Compares financial data between these sheets
3. Applies formatting and adds tickmarks to indicate matches or discrepancies
4. Performs calculations and verifications on the financial data

## Key Components

- **gui_excel_tool.py**: Main GUI application built with Tkinter that allows users to select component names and Excel files for processing
- **find_table_range.py**: Core logic for finding and processing financial data tables in Excel sheets
- **compare_ranges.py**: Functions for comparing data between different Excel sheets and marking matches/discrepancies
- **excel_utils.py**: Utility functions for handling Excel cells and values

## Running the Application

To run the UCO to UDO reconciliation tool:

```bash
# Navigate to the UCO_to_UDO_v2 directory
cd UCO_to_UDO_v2

# Install dependencies
pip install -r requirements.txt

# Run the GUI application
python gui_excel_tool.py
```

## Dependencies

The application requires the following Python packages:
- openpyxl (3.0.10) - For Excel file manipulation
- Pillow (8.4.0) - For image processing (used in the GUI)
- pywin32 (302) - For Excel COM automation
- pythoncom (1.0) - For COM communication
- tkinter - For the GUI interface
- PyQt6 (6.2.2) - Alternative GUI framework

## Technical Implementation Details

### Excel Data Handling

- **File Handle Management**: The application includes special handling to ensure Excel file handles are properly released using `ensure_file_handle_release()` function
- **Formula Preservation**: When copying sheets, the tool preserves formulas by setting `data_only=False` when loading workbooks
- **Dual Workbook Loading**: The application loads workbooks twice, once with `data_only=False` to preserve formulas and once with `data_only=True` to access calculated values
- **COM Automation**: Uses `win32com.client` to trigger Excel recalculations through COM automation

### Financial Data Processing

- **Decimal Handling**: Financial calculations use the Decimal type with ROUND_HALF_UP to ensure precise financial calculations
- **Value Conversion**: The `safe_convert_to_decimal()` function handles conversion of various string and numeric formats to Decimal
- **Comparison Logic**: Financial comparisons use a tolerance of 0.01 to account for potential rounding differences

### Reconciliation Process

1. The application copies sheets from source files to a target file
2. It recalculates all formulas using Excel COM automation
3. It identifies and processes key financial tables in each sheet
4. It compares values between sheets, looking for matches and discrepancies
5. It applies "tickmarks" (special formatted characters) to indicate verification status
6. It adds custom formulas to perform additional calculations and verification checks

## Common Issues and Solutions

1. **File Access Errors**: If Excel files are locked or inaccessible, ensure all Excel instances are closed and retry the operation
2. **COM Automation Errors**: The application includes retry logic for COM operations to handle potential COM initialization issues
3. **Missing Components**: If component names in spreadsheets don't match exactly, the application uses fallback strategies to find the correct sheets