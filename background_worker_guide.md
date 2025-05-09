# Background Worker Integration Guide

This guide explains how to use the BackgroundWorker class to safely perform long-running Excel operations in the background while keeping the GUI responsive.

## Overview

The BackgroundWorker class is designed to:

1. Run tasks in a separate thread to keep the GUI responsive
2. Provide progress updates to the GUI
3. Support cancellation of long-running operations
4. Handle error reporting and cleanup
5. Ensure clean application exit

## Key Components

### BackgroundWorker Class

The `BackgroundWorker` class (in `background_worker.py`) is responsible for:

- Managing background threads
- Handling exceptions
- Providing progress updates
- Managing cancellation
- Routing results back to the GUI

### Integration with GUI

The main application integrates with the BackgroundWorker using callbacks:

- **Progress callback**: Updates the GUI progress bar
- **Completion callback**: Handles successful task completion
- **Error callback**: Handles errors and displays them to the user

## How to Use the Background Worker

### 1. Initialize the Background Worker

```python
self.background_worker = BackgroundWorker(
    update_progress_callback=self.update_progress,
    complete_callback=self.on_task_complete,
    error_callback=self.on_task_error
)
```

### 2. Define Task Functions

Create a function that performs your long-running task:

```python
def run_processing_task(self, param1, param2, cancel_event):
    try:
        # Initialize required objects
        
        # Step 1: First operation
        if cancel_event.is_set():
            return None  # Check for cancellation
            
        # Update progress
        self.background_worker.update_progress(10)
        
        # Step 2: Next operation
        if cancel_event.is_set():
            return None  # Check for cancellation
            
        # Complete the task
        return result  # Return any data needed by the completion handler
        
    except Exception as e:
        logger.error(f"Error: {e}", exc_info=True)
        raise  # Re-raise to let the background worker handle it
```

### 3. Start the Background Task

```python
def start_operations(self):
    # Collect parameters
    param1 = self.some_input.get()
    param2 = self.another_input.get()
    
    # Update UI state
    self.progress_bar['value'] = 0
    self.start_button.config(state="disabled")
    self.cancel_button.config(state="normal")
    
    # Start worker
    try:
        self.background_worker.run_task(
            self.run_processing_task,
            param1=param1,
            param2=param2
        )
    except Exception as e:
        logger.error(f"Failed to start: {e}")
        self.on_task_error(e)
```

### 4. Handle Completion, Cancellation, and Errors

```python
def on_task_complete(self, result):
    # Update UI
    self.progress_bar['value'] = 100
    self.start_button.config(state="normal")
    self.cancel_button.config(state="disabled")
    
    # Process result
    if result:
        # Do something with the result
        messagebox.showinfo("Complete", "Task completed successfully!")

def cancel_operations(self):
    if self.background_worker.is_running():
        if messagebox.askyesno("Cancel?", "Are you sure?"):
            self.background_worker.cancel()
            
def on_task_error(self, error):
    # Update UI
    self.progress_bar['value'] = 0
    self.start_button.config(state="normal")
    self.cancel_button.config(state="disabled")
    
    # Show error to user
    messagebox.showerror("Error", f"An error occurred: {error}")
```

### 5. Clean Up on Exit

```python
def on_closing(self):
    if self.background_worker.is_running():
        if messagebox.askyesno("Exit?", "Task still running. Exit anyway?"):
            self.background_worker.cancel()
            self.destroy()
    else:
        self.destroy()
```

## Best Practices

1. **Check for Cancellation Frequently**: Add `if cancel_event.is_set(): return None` checks at logical points in your task function, especially after time-consuming operations.

2. **Thread Safety**: Use the `after()` method to update the GUI from background threads:
   ```python
   def update_progress(self, value):
       self.after(0, lambda: self._safe_update_progress(value))
   ```

3. **Release Resources**: Make sure all resources (especially Excel file handles) are properly released, even if the operation is cancelled.

4. **Handle Excel Operations Safely**: When working with Excel files:
   - Close workbooks properly
   - Use `ensure_file_handle_release()` after operations
   - Catch and handle Excel-specific exceptions

5. **Granular Progress Updates**: Provide regular progress updates (0-100) to show the user that the application is still responsive.

## Example Usage

The `background_worker_test.py` file demonstrates a complete example of background worker integration that simulates Excel operations. Run it to see:

1. How progress updates work
2. How cancellation is handled
3. How errors are reported
4. How cleanup works on exit

To run the test:
```
python background_worker_test.py
```

## Common Issues and Solutions

1. **GUI Freezes**: If the GUI still freezes, make sure all long-running operations are being properly moved to the background thread.

2. **Resources Not Released**: Use `try/finally` blocks to ensure resources are released even if an error occurs.

3. **Excel COM Interaction**: When using COM automation with Excel, always handle COM exceptions specifically and ensure the Excel application is properly closed.

4. **Memory Leaks**: For large Excel files, monitor memory usage and implement proper cleanup to avoid memory leaks.