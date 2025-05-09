#!/usr/bin/env python3
"""
UCO to UDO Reconciliation Tool - Main entry point
"""
import os
import sys
import tkinter as tk
import logging
import traceback
from datetime import datetime
from gui_excel_tool import MainWindow

def setup_logger():
    """Set up a logger for uncaught exceptions and general logging."""
    logger = logging.getLogger("MainLogger")
    logger.setLevel(logging.DEBUG)
    
    # Create logs directory if it doesn't exist
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)
    
    # Create a file handler for the log file
    log_filename = os.path.join(log_dir, f"UCOtoUDORecon_Log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt")
    file_handler = logging.FileHandler(log_filename)
    file_handler.setLevel(logging.DEBUG)
    
    # Create a formatter and add it to the handler
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    
    # Add the handler to the logger
    logger.addHandler(file_handler)
    
    return logger

def handle_exception(exc_type, exc_value, exc_traceback):
    """Handle uncaught exceptions and log them."""
    if issubclass(exc_type, KeyboardInterrupt):
        # Let the default handler handle keyboard interrupts
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    
    logger = logging.getLogger("MainLogger")
    logger.critical("Unhandled exception", exc_info=(exc_type, exc_value, exc_traceback))
    
    # Show error message to user
    error_msg = f"An unhandled error occurred:\n{exc_value}\n\nSee log for details."
    try:
        tk.messagebox.showerror("Unhandled Error", error_msg)
    except:
        print(error_msg)

def main():
    """Main entry point for the application."""
    # Set up logging
    logger = setup_logger()
    logger.info("Application starting up")
    
    # Set up global exception handler
    sys.excepthook = handle_exception
    
    try:
        # Initialize and run the main application
        app = MainWindow()
        logger.info("GUI initialized successfully")
        app.mainloop()
        logger.info("Application closed normally")
    except Exception as e:
        logger.critical(f"Fatal error during application startup: {e}", exc_info=True)
        error_msg = f"An error occurred while starting the application:\n{e}\n\nSee log for details."
        try:
            tk.messagebox.showerror("Startup Error", error_msg)
        except:
            print(error_msg)
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())