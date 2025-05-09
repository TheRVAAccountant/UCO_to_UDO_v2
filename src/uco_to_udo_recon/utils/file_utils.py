"""
File handling utilities for the UCO to UDO Reconciliation tool.

This module provides helper functions for file operations and Excel file handling.
"""

import os
import sys
import subprocess
import gc
import time
import logging
from pathlib import Path
from typing import Optional


def ensure_file_handle_release(file_path: str, logger: logging.Logger) -> None:
    """
    Ensures Python releases the file handle before Excel operations.
    
    Args:
        file_path: Path to the file that needs handle release
        logger: Logger instance for operation tracking
        
    Returns:
        None
    """
    try:
        # Force Python garbage collection
        gc.collect()
        
        # Give Windows time to release file handles
        time.sleep(2)
        
        logger.info(f"Released file handle for: {file_path}")
    except Exception as e:
        logger.warning(f"Error during file handle release: {e}")


def open_excel_file(file_path: str, logger: logging.Logger) -> None:
    """
    Open the Excel file using the default system application.
    
    Args:
        file_path: Path to the Excel file to open
        logger: Logger instance for operation tracking
        
    Returns:
        None
    """
    try:
        if os.name == 'nt':  # Windows
            os.startfile(file_path)
        elif os.name == 'posix':  # macOS and Linux
            opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
            subprocess.call([opener, file_path])
        logger.info(f"Opened Excel file: {file_path}")
    except Exception as e:
        logger.error(f"Failed to open Excel file: {e}", exc_info=True)