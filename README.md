# UCO to UDO Reconciliation Tool v2

A Python application for reconciling Unfilled Customer Orders (UCO) and Undelivered Orders (UDO) data across government component Excel files.

## Overview

The UCO to UDO Reconciliation Tool automates the process of copying, comparing, and reconciling financial data between different Excel spreadsheets. It provides a user-friendly interface for selecting files, processing data, and viewing results.

## Features

- Copy specific sheets between Excel files while preserving formatting and formulas
- Identify and compare matching data ranges between sheets
- Apply formatting and tickmarks to indicate matches or discrepancies
- Perform calculations and verify financial data across different sources
- Background processing to keep the UI responsive during operations

## Project Structure

```
UCO_to_UDO_v2/
│
├── src/                     # Source code directory
│   └── uco_to_udo_recon/    # Main package
│       ├── core/            # Core business logic
│       │   ├── comparison.py       # Data comparison operations
│       │   ├── excel_operations.py # Excel file operations
│       │   └── reconciliation.py   # Reconciliation logic
│       │
│       ├── modules/         # Application modules
│       │   ├── background_worker.py # Threaded background processing
│       │   └── gui.py              # Graphical user interface 
│       │
│       ├── utils/           # Utility functions
│       │   ├── excel_utils.py  # Excel helper functions
│       │   └── file_utils.py   # File handling utilities
│       │
│       └── main.py          # Application entry point
│
├── tests/                   # Test directory
│   ├── test_excel_utils.py    # Tests for Excel utilities
│   └── test_background_worker.py # Tests for background worker
│
├── logs/                    # Log files directory
├── forest-dark/             # UI theme files
├── forest-light/            # UI theme files
├── forest-dark.tcl          # Dark theme definition
├── forest-light.tcl         # Light theme definition
├── requirements.txt         # Project dependencies
└── README.md                # Project documentation
```

## Key Components

- **GUI Module**: User interface for file selection and operation control
- **Background Worker**: Handles threaded operations to keep the UI responsive
- **Excel Operations**: Core functions for manipulating Excel workbooks
- **Reconciliation Logic**: Business logic for comparing and reconciling data

## Background Worker

The `background_worker` module provides robust threading support for long-running operations:

- **BackgroundWorker**: Base class for running tasks in background threads with progress updates
- **ProgressTracker**: Manages progress across multiple sequential tasks
- **TaskManager**: Handles complex workflows with task dependencies

Example usage:

```python
from src.uco_to_udo_recon.modules.background_worker import BackgroundWorker

# Create a worker with callbacks
worker = BackgroundWorker(
    on_progress=update_progress_ui,
    on_complete=handle_task_completion,
    on_message=handle_status_message,
    logger=logger
)

# Start the worker
worker.start()

# Queue a task
worker.queue_task(
    my_long_running_function,
    args=(arg1, arg2),
    kwargs={"param1": value1},
    task_name="Excel Processing"
)
```

See the [background worker documentation](docs/background_worker.md) for detailed usage instructions and examples.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/username/UCO_to_UDO_v2.git
   cd UCO_to_UDO_v2
   ```

2. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Run the application using:

```bash
python -m src.uco_to_udo_recon.main
```

1. Select the component from the dropdown menu
2. Choose the UCO to UDO Reconciliation File
3. Select the Trial Balance File
4. Select the UCO to UDO TIER File
5. Click "Start Reconciliation"

## Testing

Run tests using:

```bash
python -m unittest discover tests
```

## Requirements

- Python 3.7 or higher
- openpyxl 3.0.10
- Pillow 8.4.0
- pywin32 302 (for Windows)
- tkinter