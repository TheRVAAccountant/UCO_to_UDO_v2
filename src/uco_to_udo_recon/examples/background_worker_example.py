"""
Example usage of the background worker module.

This module demonstrates how to use the BackgroundWorker and ProgressTracker
classes for handling long-running operations in the background.
"""

import time
import logging
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Any, List, Tuple

# Import the background worker classes
from src.uco_to_udo_recon.modules.background_worker import BackgroundWorker, ProgressTracker


class ExampleApp(tk.Tk):
    """Example application demonstrating background worker usage."""
    
    def __init__(self):
        """Initialize the example application."""
        super().__init__()
        
        # Setup basic UI
        self.title("Background Worker Example")
        self.geometry("600x400")
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Configure logger
        self.logger = logging.getLogger("ExampleLogger")
        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)
        self.logger.setLevel(logging.DEBUG)
        
        # Initialize background worker
        self.worker = BackgroundWorker(
            on_progress=self.update_progress_from_worker,
            on_complete=self.on_task_complete,
            on_message=self.on_worker_message,
            logger=self.logger
        )
        
        # Track if operation is in progress
        self.processing = False
        self.task_cancel_requested = False
        
        # Create UI components
        self.create_ui()
    
    def create_ui(self):
        """Create the user interface components."""
        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Task type selection
        ttk.Label(frame, text="Task Type:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.task_type_var = tk.StringVar(value="simple")
        task_types = [
            ("Simple Task", "simple"),
            ("Multi-stage Task", "multi_stage"),
            ("Error-prone Task", "error_prone"),
            ("Cancellable Task", "cancellable")
        ]
        
        task_frame = ttk.Frame(frame)
        task_frame.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        for i, (text, value) in enumerate(task_types):
            ttk.Radiobutton(
                task_frame, 
                text=text, 
                value=value, 
                variable=self.task_type_var
            ).grid(row=0, column=i, padx=5)
        
        # Duration settings
        ttk.Label(frame, text="Duration (seconds):").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        self.duration_var = tk.IntVar(value=5)
        duration_scale = ttk.Scale(
            frame, 
            from_=1, 
            to=20, 
            orient=tk.HORIZONTAL, 
            variable=self.duration_var, 
            length=200
        )
        duration_scale.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(frame, textvariable=self.duration_var).grid(row=1, column=2, padx=5)
        
        # Status display
        ttk.Label(frame, text="Status:").grid(row=2, column=0, sticky=tk.W, pady=5)
        
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(frame, textvariable=self.status_var).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Progress bar
        ttk.Label(frame, text="Progress:").grid(row=3, column=0, sticky=tk.W, pady=5)
        
        self.progress_bar = ttk.Progressbar(frame, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.progress_bar.grid(row=3, column=1, sticky=tk.W, pady=5)
        
        self.progress_label = ttk.Label(frame, text="")
        self.progress_label.grid(row=3, column=2, sticky=tk.W, pady=5)
        
        # Log output
        ttk.Label(frame, text="Log:").grid(row=4, column=0, sticky=tk.NW, pady=5)
        
        self.log_text = tk.Text(frame, height=10, width=50)
        self.log_text.grid(row=4, column=1, columnspan=2, sticky=tk.NSEW, pady=5)
        
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=4, column=3, sticky=tk.NS, pady=5)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # Button frame
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=10)
        
        self.start_button = ttk.Button(button_frame, text="Start Task", command=self.start_task)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.cancel_button = ttk.Button(
            button_frame, 
            text="Cancel Task", 
            command=self.cancel_task, 
            state=tk.DISABLED
        )
        self.cancel_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Quit", command=self.on_close).pack(side=tk.LEFT, padx=5)
        
        # Configure grid weights
        frame.rowconfigure(4, weight=1)
        frame.columnconfigure(1, weight=1)
    
    def log(self, message):
        """Add a message to the log display."""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
    
    def start_task(self):
        """Start the selected task in the background."""
        if self.processing:
            messagebox.showinfo("Task in Progress", "A task is already running.")
            return
        
        # Get task parameters
        task_type = self.task_type_var.get()
        duration = self.duration_var.get()
        
        # Update UI state
        self.processing = True
        self.task_cancel_requested = False
        self.start_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.progress_bar['value'] = 0
        self.progress_label.config(text="")
        self.status_var.set("Processing...")
        
        # Log start
        self.log(f"Starting {task_type} task with duration {duration}s")
        
        # Select appropriate task function based on type
        if task_type == "simple":
            self.start_simple_task(duration)
        elif task_type == "multi_stage":
            self.start_multi_stage_task(duration)
        elif task_type == "error_prone":
            self.start_error_prone_task(duration)
        elif task_type == "cancellable":
            self.start_cancellable_task(duration)
    
    def start_simple_task(self, duration):
        """Start a simple background task."""
        def simple_task(progress_callback):
            """A simple task that just sleeps and reports progress."""
            start_time = time.time()
            end_time = start_time + duration
            
            while time.time() < end_time:
                elapsed = time.time() - start_time
                progress = min(100, int((elapsed / duration) * 100))
                progress_callback(progress, f"Processing: {progress}%")
                time.sleep(0.1)
            
            return "Simple task completed successfully!"
        
        # Queue the task
        self.worker.queue_task(simple_task, task_name="Simple Task")
    
    def start_multi_stage_task(self, duration):
        """Start a multi-stage background task using ProgressTracker."""
        # Calculate stage durations
        stage_time = duration / 3
        
        # Define stages
        stages = [
            ("Preparation", 20),  # 20% of total
            ("Processing", 50),   # 50% of total
            ("Finalization", 30)  # 30% of total
        ]
        
        # Create progress tracker
        self.progress_tracker = ProgressTracker(stages, self.update_progress)
        
        def multi_stage_task():
            """A multi-stage task that reports progress through ProgressTracker."""
            # Stage 1: Preparation
            self.progress_tracker.update(0, "Starting preparation...")
            
            stage_start = time.time()
            stage_end = stage_start + stage_time
            
            while time.time() < stage_end:
                if self.task_cancel_requested:
                    return "Task cancelled during preparation"
                
                elapsed = time.time() - stage_start
                progress = min(100, int((elapsed / stage_time) * 100))
                self.progress_tracker.update(progress, f"Preparation: {progress}%")
                time.sleep(0.1)
            
            # Move to stage 2
            self.progress_tracker.next_stage()
            self.progress_tracker.update(0, "Starting processing...")
            
            stage_start = time.time()
            stage_end = stage_start + stage_time
            
            while time.time() < stage_end:
                if self.task_cancel_requested:
                    return "Task cancelled during processing"
                
                elapsed = time.time() - stage_start
                progress = min(100, int((elapsed / stage_time) * 100))
                self.progress_tracker.update(progress, f"Processing: {progress}%")
                time.sleep(0.1)
            
            # Move to stage 3
            self.progress_tracker.next_stage()
            self.progress_tracker.update(0, "Starting finalization...")
            
            stage_start = time.time()
            stage_end = stage_start + stage_time
            
            while time.time() < stage_end:
                if self.task_cancel_requested:
                    return "Task cancelled during finalization"
                
                elapsed = time.time() - stage_start
                progress = min(100, int((elapsed / stage_time) * 100))
                self.progress_tracker.update(progress, f"Finalization: {progress}%")
                time.sleep(0.1)
            
            self.progress_tracker.update(100, "Multi-stage task complete!")
            return "Multi-stage task completed successfully!"
        
        # Queue the task
        self.worker.queue_task(multi_stage_task, task_name="Multi-stage Task")
    
    def start_error_prone_task(self, duration):
        """Start a task that might generate errors."""
        def error_task(progress_callback):
            """A task that will raise an exception if duration > 5."""
            start_time = time.time()
            halfway_point = start_time + (duration / 2)
            end_time = start_time + duration
            
            # First half works normally
            while time.time() < halfway_point:
                elapsed = time.time() - start_time
                progress = min(50, int((elapsed / duration) * 100))
                progress_callback(progress, f"Processing: {progress}%")
                time.sleep(0.1)
            
            # Check if we should trigger an error
            if duration > 5:
                raise ValueError(f"Task failed because duration ({duration}s) is > 5s")
            
            # Continue if duration <= 5
            while time.time() < end_time:
                elapsed = time.time() - start_time
                progress = min(100, int((elapsed / duration) * 100))
                progress_callback(progress, f"Processing: {progress}%")
                time.sleep(0.1)
            
            return "Error-prone task completed successfully!"
        
        # Queue the task
        self.worker.queue_task(error_task, task_name="Error-prone Task")
    
    def start_cancellable_task(self, duration):
        """Start a task that checks for cancellation."""
        def cancellable_task(progress_callback, cancellation_check):
            """A task that periodically checks if it should be cancelled."""
            start_time = time.time()
            end_time = start_time + duration
            
            while time.time() < end_time:
                # Check for cancellation
                if cancellation_check():
                    progress_callback(0, "Task cancelled")
                    return "Task cancelled by user request"
                
                elapsed = time.time() - start_time
                progress = min(100, int((elapsed / duration) * 100))
                progress_callback(progress, f"Processing: {progress}%")
                time.sleep(0.1)
            
            return "Cancellable task completed without cancellation!"
        
        # Queue the task
        self.worker.queue_task(cancellable_task, task_name="Cancellable Task")
    
    def cancel_task(self):
        """Request cancellation of the current task."""
        if not self.processing:
            return
        
        if messagebox.askyesno("Cancel Task", "Are you sure you want to cancel the task?"):
            self.log("Requesting task cancellation...")
            self.task_cancel_requested = True
            self.status_var.set("Cancelling...")
            self.cancel_button.config(state=tk.DISABLED)
    
    def update_progress_from_worker(self, value: int, message: Optional[str] = None):
        """Update progress from worker thread."""
        # Schedule UI update on the main thread
        self.after(0, lambda: self.update_progress(value, message))
    
    def update_progress(self, value: int, message: Optional[str] = None):
        """Update the progress bar and label."""
        self.progress_bar['value'] = value
        
        if message is not None:
            self.progress_label.config(text=message)
            self.log(f"Progress: {message}")
    
    def on_worker_message(self, message: str, level: str = "info"):
        """Handle messages from worker thread."""
        self.log(f"[{level.upper()}] {message}")
    
    def on_task_complete(self, success: bool, result: Any, error: Optional[Exception]):
        """Handle task completion from worker thread."""
        self.after(0, lambda: self._handle_task_completion(success, result, error))
    
    def _handle_task_completion(self, success: bool, result: Any, error: Optional[Exception]):
        """Process task completion and update UI accordingly."""
        self.processing = False
        
        # Reset the UI state
        self.start_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        
        if success:
            self.progress_bar['value'] = 100
            self.status_var.set("Complete")
            self.log(f"Task completed: {result}")
            messagebox.showinfo("Task Complete", str(result))
        else:
            self.progress_bar['value'] = 0
            self.status_var.set("Error")
            error_message = str(error)
            self.log(f"Task failed: {error_message}")
            messagebox.showerror("Task Failed", f"Error: {error_message}")
    
    def on_close(self):
        """Handle application closing."""
        if self.processing:
            if messagebox.askyesno(
                "Task in Progress",
                "A task is currently running. Close anyway?"
            ):
                self.task_cancel_requested = True
                self.worker.stop()
                self.destroy()
        else:
            self.worker.stop()
            self.destroy()


def main():
    """Run the example application."""
    app = ExampleApp()
    app.mainloop()


if __name__ == "__main__":
    main()