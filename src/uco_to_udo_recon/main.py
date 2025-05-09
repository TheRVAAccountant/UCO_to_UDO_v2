"""
Main entry point for the UCO to UDO Reconciliation application.

This module provides the main entry point for the application, handling
startup, logging configuration, and the primary execution flow.
"""

import logging
import os
import sys
from datetime import datetime
from pathlib import Path

# Set up paths
ROOT_DIR = Path(__file__).parent.parent.parent
sys.path.insert(0, str(ROOT_DIR))

# Import application modules
from src.uco_to_udo_recon.modules.gui import MainWindow


def setup_logging() -> logging.Logger:
    """
    Configure and set up the application's logging system.
    
    Returns:
        logging.Logger: Configured logger instance
    """
    # Create logs directory if it doesn't exist
    log_dir = ROOT_DIR / "logs"
    log_dir.mkdir(exist_ok=True)
    
    # Set up logger
    log_filename = log_dir / f"UCOtoUDORecon_Log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
    
    logger = logging.getLogger("MainLogger")
    logger.setLevel(logging.DEBUG)
    
    # Set up file handler
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler = logging.FileHandler(str(log_filename))
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    return logger


def main():
    """
    Main entry point function.
    
    Initializes the application, sets up logging, and launches the GUI.
    """
    # Set up logging
    logger = setup_logging()
    logger.info("Application starting...")
    
    try:
        # Launch the GUI
        app = MainWindow()
        app.mainloop()
    except Exception as e:
        logger.critical(f"Critical error in application: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()