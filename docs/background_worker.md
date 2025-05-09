# Background Worker Documentation

The Background Worker module provides a robust solution for running long operations in separate threads while keeping the GUI responsive. It handles progress updates, completion callbacks, error reporting, and task cancellation.

## Core Components

### BackgroundWorker

The `BackgroundWorker` class manages a background thread that processes tasks from a queue:

```python
from src.uco_to_udo_recon.modules.background_worker import BackgroundWorker

# Create worker with callbacks
worker = BackgroundWorker(
    on_progress=update_progress_ui,  # Callback for progress updates
    on_complete=handle_task_completion,  # Callback for task completion
    on_message=handle_status_message,  # Callback for log messages
    logger=logger  # Logger instance
)

# Start the worker thread
worker.start()

# Queue a task
worker.queue_task(
    task_function,  # Function to execute in background
    args=(arg1, arg2),  # Positional arguments
    kwargs={"param1": value1},  # Keyword arguments
    task_name="My Task"  # Task name for logging
)

# Later, stop the worker when done
worker.stop()
```

#### Key Methods

- `start()`: Start the worker thread if not already running
- `stop()`: Stop the worker thread and clean up resources
- `queue_task(task_func, args=None, kwargs=None, task_name=None, task_id=None)`: Queue a task for execution
- `cancel_task(task_id)`: Cancel a specific task
- `is_task_cancelled(task_id)`: Check if a task is marked for cancellation
- `clear_cancelled_tasks()`: Remove all cancelled tasks from the queue

#### Callback Functions

- `on_progress(value, message)`: Called with progress updates (0-100)
- `on_complete(success, result, error)`: Called when a task completes
- `on_message(message, level)`: Called when the worker wants to log a message

### ProgressTracker

The `ProgressTracker` manages progress reporting across multiple sequential stages:

```python
from src.uco_to_udo_recon.modules.background_worker import ProgressTracker

# Define stages with weights
stages = [
    ("Copy Files", 10),        # 10% of total progress
    ("Process Data", 60),      # 60% of total progress
    ("Generate Report", 30)    # 30% of total progress
]

# Create tracker with callback
tracker = ProgressTracker(stages, update_progress_callback)

# Stage 1: Report progress within first stage
tracker.update(50, "Copying files...")  # Results in 5% overall progress

# Move to next stage
tracker.next_stage()

# Stage 2: Report progress within second stage
tracker.update(25, "Processing data...")  # Results in 10% + (25% of 60%) = 25% overall
```

#### Key Methods

- `next_stage()`: Move to the next stage in the sequence
- `update(stage_progress, message=None)`: Update progress within current stage (0-100)

### TaskManager

The `TaskManager` manages complex workflows with task dependencies:

```python
from src.uco_to_udo_recon.modules.background_worker import TaskManager, BackgroundWorker

# Create worker and task manager
worker = BackgroundWorker(...)
task_manager = TaskManager(worker)

# Add tasks with dependencies
task_manager.add_task(
    "task1", 
    task_func1
)

task_manager.add_task(
    "task2", 
    task_func2, 
    dependencies=["task1"]  # This task depends on task1
)

task_manager.add_task(
    "task3", 
    task_func3, 
    dependencies=["task1"]  # This also depends on task1
)

# Execute the workflow
task_manager.execute_workflow()
```

## Integration with GUI

To integrate with a GUI, you need to handle thread synchronization carefully:

```python
class MyApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # Create UI components
        self.progress_bar = ttk.Progressbar(self)
        self.progress_bar.pack()
        
        self.status_label = ttk.Label(self, text="Ready")
        self.status_label.pack()
        
        self.start_button = ttk.Button(self, text="Start", command=self.start_task)
        self.start_button.pack()
        
        self.cancel_button = ttk.Button(
            self, text="Cancel", command=self.cancel_task, state=tk.DISABLED
        )
        self.cancel_button.pack()
        
        # Initialize worker
        self.worker = BackgroundWorker(
            on_progress=self.update_progress_from_worker,
            on_complete=self.on_task_complete,
            on_message=self.on_worker_message,
            logger=logger
        )
        
        # State tracking
        self.processing = False
        self.task_cancel_requested = False
        
        # Set up window close handling
        self.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def update_progress_from_worker(self, value, message=None):
        """Update progress from worker thread safely."""
        # Schedule UI update on the main thread
        self.after(0, lambda: self.update_progress(value, message))
    
    def update_progress(self, value, message=None):
        """Update UI with progress."""
        self.progress_bar['value'] = value
        if message:
            self.status_label.config(text=message)
    
    def on_worker_message(self, message, level="info"):
        """Handle messages from worker thread."""
        log_method = getattr(logger, level.lower(), logger.info)
        log_method(message)
    
    def on_task_complete(self, success, result, error):
        """Handle task completion from worker thread."""
        # Schedule UI update on the main thread
        self.after(0, lambda: self._handle_task_completion(success, result, error))
    
    def _handle_task_completion(self, success, result, error):
        """Process task completion and update UI."""
        self.processing = False
        self.start_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        
        if success:
            messagebox.showinfo("Complete", "Task completed successfully!")
        else:
            messagebox.showerror("Error", f"Task failed: {error}")
    
    def start_task(self):
        """Start a task in the background."""
        if self.processing:
            return
        
        self.processing = True
        self.task_cancel_requested = False
        self.start_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        
        # Queue the task
        self.worker.queue_task(
            self.my_long_running_task,
            task_name="Long Operation"
        )
    
    def my_long_running_task(self, progress_callback=None, cancellation_check=None):
        """Sample long-running task."""
        total_steps = 100
        for i in range(total_steps):
            # Check for cancellation if function was provided
            if cancellation_check and cancellation_check():
                return "Task was cancelled"
            
            # Perform work...
            time.sleep(0.1)
            
            # Report progress if callback was provided
            if progress_callback:
                progress = (i + 1) / total_steps * 100
                progress_callback(progress, f"Step {i+1}/{total_steps}")
        
        return "Task completed successfully"
    
    def cancel_task(self):
        """Cancel the current task."""
        if self.processing and messagebox.askyesno(
            "Cancel Task",
            "Are you sure you want to cancel the current task?"
        ):
            self.task_cancel_requested = True
            self.cancel_button.config(state=tk.DISABLED)
    
    def on_close(self):
        """Handle window closing."""
        if self.processing:
            if messagebox.askyesno(
                "Task in Progress",
                "A task is running. Close anyway?"
            ):
                self.worker.stop()
                self.destroy()
        else:
            self.worker.stop()
            self.destroy()
```

## Important Patterns

### Progress Reporting

Tasks can report progress in two ways:

1. **Direct callback**: Use a progress callback function provided by the worker
   ```python
   def my_task(progress_callback=None):
       # Report 50% progress
       if progress_callback:
           progress_callback(50, "Halfway done")
   ```

2. **ProgressTracker**: For multi-stage tasks, use the ProgressTracker
   ```python
   def my_task():
       # Using a tracker initialized elsewhere
       tracker.update(50, "Stage 1: 50% complete")
   ```

### Cancellation Support

Tasks can support cancellation by checking a cancellation flag:

```python
def cancellable_task(progress_callback=None, cancellation_check=None):
    while not_done:
        # Check if cancellation was requested
        if cancellation_check and cancellation_check():
            cleanup_resources()
            return "Operation cancelled"
        
        # Continue processing...
```

### Error Handling

The worker handles exceptions automatically:

```python
def risky_task():
    try:
        # Risky operation
        result = perform_operation()
        return result
    except Exception as e:
        # Optional: do cleanup before re-raising
        cleanup_resources()
        raise  # Worker will catch and report this exception
```

## Examples

See the `/src/uco_to_udo_recon/examples/background_worker_example.py` for a complete, working example of background worker usage.

## Best Practices

1. **Always update UI from main thread**: Use `after(0, lambda: update_ui())` to ensure UI updates happen on the main thread
2. **Handle task completion properly**: Always reset UI state after task completion
3. **Clean up resources**: Always stop the worker when the application closes
4. **Support cancellation**: Make long-running tasks cancellable for better user experience
5. **Report progress regularly**: Update progress frequently to keep users informed
6. **Throttle progress updates**: Avoid flooding the UI with too many updates