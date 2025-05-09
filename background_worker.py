import threading
import queue
import logging
import traceback
from typing import Dict, Any, Callable

class BackgroundWorker:
    """
    Worker class to handle background operations with progress updates and error handling.
    """
    def __init__(self, update_progress_callback=None, complete_callback=None, error_callback=None):
        """
        Initialize a background worker with callbacks.
        
        Args:
            update_progress_callback: Function to call with progress updates (0-100)
            complete_callback: Function to call when the operation completes successfully
            error_callback: Function to call when an error occurs
        """
        self.update_progress_callback = update_progress_callback
        self.complete_callback = complete_callback
        self.error_callback = error_callback
        self.thread = None
        self.running = False
        self.logger = logging.getLogger("BackgroundWorker")
        self.results_queue = queue.Queue()
        self.cancel_event = threading.Event()

    def run_task(self, task_function, *args, **kwargs):
        """
        Run a task in the background.
        
        Args:
            task_function: The function to run in the background
            *args, **kwargs: Arguments to pass to the task function
        """
        if self.thread and self.thread.is_alive():
            self.logger.warning("A task is already running. Please wait for it to complete.")
            return False
        
        # Reset cancel event
        self.cancel_event.clear()
        
        # Start the thread
        self.thread = threading.Thread(
            target=self._thread_wrapper,
            args=(task_function, args, kwargs),
            daemon=True
        )
        self.running = True
        self.thread.start()
        return True

    def _thread_wrapper(self, task_function, args, kwargs):
        """Wrapper that handles exceptions and callbacks."""
        try:
            # Add cancel_event to kwargs
            kwargs['cancel_event'] = self.cancel_event
            
            # Run the task
            result = task_function(*args, **kwargs)
            
            # Put the result in the queue
            self.results_queue.put({'status': 'success', 'result': result})
            
            # Call complete callback on success
            if self.complete_callback and not self.cancel_event.is_set():
                self.complete_callback(result)
                
        except Exception as e:
            self.logger.error(f"Error in background task: {e}", exc_info=True)
            error_info = {
                'error': str(e),
                'traceback': traceback.format_exc()
            }
            self.results_queue.put({'status': 'error', 'error': error_info})
            
            # Call error callback
            if self.error_callback:
                self.error_callback(e)
        finally:
            self.running = False

    def cancel(self):
        """Cancel the currently running operation."""
        if self.running:
            self.logger.info("Canceling background operation...")
            self.cancel_event.set()
            return True
        return False

    def update_progress(self, value):
        """Update progress value (0-100)."""
        if self.update_progress_callback:
            self.update_progress_callback(value)
            
    def get_results(self):
        """Get the results of the operation if available."""
        if not self.results_queue.empty():
            return self.results_queue.get()
        return None

    def is_running(self):
        """Check if the worker is currently running."""
        return self.running