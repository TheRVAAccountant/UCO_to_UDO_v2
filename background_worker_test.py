import os
import sys
import time
import random
import logging
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from background_worker import BackgroundWorker

class TextHandler(logging.Handler):
    """Custom handler to redirect logging output to a tkinter Text widget"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.see(tk.END)
        self.text_widget.update_idletasks()

class ExcelSimulator:
    """Class to simulate Excel operations"""
    def __init__(self, logger, update_progress=None, cancel_event=None):
        self.logger = logger
        self.update_progress = update_progress
        self.cancel_event = cancel_event

    def simulate_loading_file(self, file_name, duration=3):
        """Simulate the loading of an Excel file"""
        self.logger.info(f"Loading Excel file: {file_name}")
        for i in range(10):
            if self.cancel_event and self.cancel_event.is_set():
                self.logger.info("Loading cancelled")
                return False
            time.sleep(duration/10)
            if self.update_progress:
                progress_value = (i+1) * 10
                self.update_progress(progress_value)
                self.logger.info(f"Loading progress: {progress_value}%")
        return True

    def simulate_processing(self, sheet_name, rows=1000, duration=5):
        """Simulate processing operations on Excel data"""
        self.logger.info(f"Processing sheet: {sheet_name} with {rows} rows")
        start_progress = 10
        for i in range(9):  # 10% to 100%
            if self.cancel_event and self.cancel_event.is_set():
                self.logger.info("Processing cancelled")
                return False
            time.sleep(duration/9)
            if self.update_progress:
                progress_value = start_progress + (i+1) * 10
                self.update_progress(progress_value)
                self.logger.info(f"Processing progress: {progress_value}%")
                
            # Simulate occasional errors
            if random.random() < 0.05:  # 5% chance of error
                if random.random() < 0.5:  # 50% of errors are recoverable
                    self.logger.warning(f"Recovered from minor error in row {random.randint(1, rows)}")
                else:
                    error_row = random.randint(1, rows)
                    error_msg = f"Error processing row {error_row} in sheet {sheet_name}"
                    self.logger.error(error_msg)
                    if random.random() < 0.2:  # 20% of errors are fatal
                        raise ValueError(error_msg)
        return True

    def simulate_formatting(self, duration=2):
        """Simulate formatting operations in Excel"""
        self.logger.info("Applying formatting to worksheet")
        for i in range(10):
            if self.cancel_event and self.cancel_event.is_set():
                self.logger.info("Formatting cancelled")
                return False
            time.sleep(duration/10)
        return True

    def simulate_saving(self, file_name, duration=2):
        """Simulate saving an Excel file"""
        self.logger.info(f"Saving file: {file_name}")
        for i in range(10):
            if self.cancel_event and self.cancel_event.is_set():
                self.logger.info("Saving cancelled")
                return False
            time.sleep(duration/10)
        return True

class TestWindow(tk.Tk):
    """Test GUI window to demonstrate background worker integration"""
    def __init__(self):
        super().__init__()
        self.title('Background Worker Test')
        self.geometry('600x500')
        self.configure(bg='#f0f0f0')

        # Make the window resizable
        self.resizable(True, True)

        # Initialize the background worker
        self.background_worker = BackgroundWorker(
            update_progress_callback=self.update_progress,
            complete_callback=self.on_task_complete,
            error_callback=self.on_task_error
        )

        self.init_ui()
        
        # Bind the window close event
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def init_ui(self):
        """Initialize the user interface"""
        # Configure grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # Test parameters frame
        param_frame = ttk.LabelFrame(self, text="Test Parameters")
        param_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        param_frame.grid_columnconfigure(1, weight=1)

        # File name
        ttk.Label(param_frame, text="File Name:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.file_name_var = tk.StringVar(value="example.xlsx")
        ttk.Entry(param_frame, textvariable=self.file_name_var).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Sheet name
        ttk.Label(param_frame, text="Sheet Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.sheet_name_var = tk.StringVar(value="Sheet1")
        ttk.Entry(param_frame, textvariable=self.sheet_name_var).grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # Row count
        ttk.Label(param_frame, text="Row Count:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.row_count_var = tk.IntVar(value=1000)
        ttk.Entry(param_frame, textvariable=self.row_count_var).grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        # Error simulation
        ttk.Label(param_frame, text="Simulate Error:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.simulate_error_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(param_frame, variable=self.simulate_error_var).grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # Log text area
        log_frame = ttk.LabelFrame(self, text="Log")
        log_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, width=70, height=10, bg='#ffffff', fg='black')
        self.log_text.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # Progress bar
        progress_frame = ttk.Frame(self)
        progress_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        progress_frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=580, mode="determinate")
        self.progress_bar.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        # Status label
        self.status_label = ttk.Label(progress_frame, text="Ready")
        self.status_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        # Buttons frame
        button_frame = ttk.Frame(self)
        button_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        button_frame.grid_columnconfigure(2, weight=1)

        # Start Button
        self.start_button = ttk.Button(button_frame, text="Start", command=self.start_simulation)
        self.start_button.grid(row=0, column=0, padx=5, pady=5, sticky="e")

        # Cancel Button
        self.cancel_button = ttk.Button(button_frame, text="Cancel", command=self.cancel_simulation, state="disabled")
        self.cancel_button.grid(row=0, column=1, padx=5, pady=5)
        
        # Exit Button
        self.exit_button = ttk.Button(button_frame, text="Exit", command=self.on_closing)
        self.exit_button.grid(row=0, column=2, padx=5, pady=5, sticky="w")

        # Set up logging
        self.logger = self.setup_logging()

    def setup_logging(self):
        """Set up logging with file and text handlers"""
        logger = logging.getLogger("TestLogger")
        logger.setLevel(logging.DEBUG)

        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

        # File handler
        log_dir = "logs"
        os.makedirs(log_dir, exist_ok=True)
        log_filename = os.path.join(log_dir, f"BackgroundWorkerTest_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt")
        file_handler = logging.FileHandler(log_filename)
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        # GUI Text handler
        text_handler = TextHandler(self.log_text)
        text_handler.setLevel(logging.DEBUG)
        text_handler.setFormatter(formatter)
        logger.addHandler(text_handler)

        return logger

    def start_simulation(self):
        """Start the Excel simulation in a background thread"""
        file_name = self.file_name_var.get()
        sheet_name = self.sheet_name_var.get()
        row_count = self.row_count_var.get()
        simulate_error = self.simulate_error_var.get()

        # Reset UI
        self.progress_bar['value'] = 0
        self.update_idletasks()
        self.status_label.config(text="Processing...")

        # Disable start button, enable cancel button
        self.start_button.config(state="disabled")
        self.cancel_button.config(state="normal")

        # Start background worker
        self.logger.info("Excel simulation started...")
        try:
            self.background_worker.run_task(
                self.run_simulation_task,
                file_name=file_name,
                sheet_name=sheet_name,
                row_count=row_count,
                simulate_error=simulate_error
            )
        except Exception as e:
            self.logger.error(f"Failed to start background worker: {e}", exc_info=True)
            self.on_task_error(e)

    def run_simulation_task(self, file_name, sheet_name, row_count, simulate_error, cancel_event):
        """Main simulation task to run in background"""
        try:
            # Initialize Excel simulator
            simulator = ExcelSimulator(
                self.logger,
                self.background_worker.update_progress,
                cancel_event
            )

            # Step 1: Load file
            self.logger.info("STEP 1: Loading Excel file")
            if not simulator.simulate_loading_file(file_name):
                return None

            # Check for cancellation
            if cancel_event.is_set():
                self.logger.info("Simulation cancelled after loading file.")
                return None

            # Step 2: Process data
            self.logger.info("STEP 2: Processing Excel data")
            if not simulator.simulate_processing(sheet_name, row_count):
                return None

            # Check for cancellation
            if cancel_event.is_set():
                self.logger.info("Simulation cancelled after processing data.")
                return None

            # Step 3: Format data
            self.logger.info("STEP 3: Formatting Excel data")
            if not simulator.simulate_formatting():
                return None

            # Simulate an error if requested
            if simulate_error:
                self.logger.error("Simulated error triggered by user")
                raise ValueError("This is a simulated error requested by user")

            # Check for cancellation
            if cancel_event.is_set():
                self.logger.info("Simulation cancelled after formatting data.")
                return None

            # Step 4: Save file
            self.logger.info("STEP 4: Saving Excel file")
            output_file = f"output_{file_name}"
            if not simulator.simulate_saving(output_file):
                return None

            self.logger.info("Excel simulation completed successfully.")
            return output_file

        except Exception as e:
            self.logger.error(f"Error during Excel simulation: {e}", exc_info=True)
            raise

    def cancel_simulation(self):
        """Cancel the ongoing simulation"""
        if self.background_worker.is_running():
            if messagebox.askyesno("Cancel Simulation", "Are you sure you want to cancel the current simulation?"):
                self.logger.info("User requested simulation cancellation.")
                self.background_worker.cancel()
                self.status_label.config(text="Cancelling...")

    def on_task_complete(self, result):
        """Handle successful completion of the background task"""
        self.progress_bar['value'] = 100
        self.status_label.config(text="Complete")
        self.start_button.config(state="normal")
        self.cancel_button.config(state="disabled")

        if result:  # If we have a valid result (file path)
            self.logger.info(f"Simulation completed successfully. Output file: {result}")
            messagebox.showinfo("Complete", f"Excel simulation completed successfully!\nOutput file: {result}")

    def on_task_error(self, error):
        """Handle errors from the background task"""
        self.progress_bar['value'] = 0
        self.status_label.config(text="Error")
        self.start_button.config(state="normal")
        self.cancel_button.config(state="disabled")

        error_message = str(error)
        self.logger.error(f"Error during simulation: {error_message}")
        messagebox.showerror("Error", f"An error occurred during Excel simulation:\n{error_message}")

    def update_progress(self, value):
        """Update the progress bar value with thread safety"""
        # Use after() to update from the main thread
        self.after(0, lambda: self._safe_update_progress(value))

    def _safe_update_progress(self, value):
        """Thread-safe progress bar update"""
        self.progress_bar['value'] = value
        self.update_idletasks()

    def on_closing(self):
        """Handle window closing event with cleanup"""
        if self.background_worker.is_running():
            if messagebox.askyesno("Cancel Simulation",
                               "A simulation is still running. Are you sure you want to exit?"):
                self.logger.info("Closing application while simulation is running.")
                self.background_worker.cancel()
                self.destroy()
        else:
            self.logger.info("Application closed normally.")
            self.destroy()

def main():
    app = TestWindow()
    app.mainloop()

if __name__ == "__main__":
    main()