"""
Background worker module for handling long-running tasks.

This module provides a background worker implementation using 
threading to keep the GUI responsive during long-running operations.
"""

import threading
import queue
import time
import traceback
import logging
import sys
from typing import Any, Callable, Dict, List, Optional, Tuple, Union


class BackgroundWorker:
    """
    A background worker that runs operations in a separate thread.
    
    This allows the GUI to remain responsive during long-running operations.
    The worker handles task queuing, status updates, and error reporting.
    """
    
    def __init__(self, on_progress: Optional[Callable[[int, str], None]] = None,
                on_complete: Optional[Callable[[bool, Any, Optional[Exception]], None]] = None,
                on_message: Optional[Callable[[str, str], None]] = None,
                logger: Optional[logging.Logger] = None):
        """
        Initialize the background worker.
        
        Args:
            on_progress: Callback function for progress updates (value, message)
            on_complete: Callback function for task completion (success, result, error)
            on_message: Callback function for status messages (message, level)
            logger: Logger instance for logging
        """
        self.task_queue = queue.Queue()
        self.on_progress = on_progress
        self.on_complete = on_complete
        self.on_message = on_message
        self.logger = logger or logging.getLogger(__name__)
        self.running = False
        self.stopping = False
        self.current_task = None
        self.worker_thread = None
        self.last_progress_time = 0
        self.progress_throttle = 0.05  # 50ms minimum between progress updates
        self.task_cancellation_flags = {}  # Track cancellation flags by task ID

    def start(self) -> None:
        """
        Start the worker thread if it's not already running.
        
        Returns:
            None
        """
        if not self.running:
            self.running = True
            self.stopping = False
            self.worker_thread = threading.Thread(target=self._process_queue, daemon=True)
            self.worker_thread.start()
            self.logger.debug("Background worker started")

    def stop(self) -> None:
        """
        Stop the worker thread.
        
        Returns:
            None
        """
        if self.running:
            self.stopping = True
            self.running = False
            if self.worker_thread and self.worker_thread.is_alive():
                self.worker_thread.join(timeout=1.0)
            self.logger.debug("Background worker stopped")

    def queue_task(self, task_func: Callable[..., Any], 
                  args: Optional[Tuple] = None, 
                  kwargs: Optional[Dict[str, Any]] = None,
                  task_name: Optional[str] = None,
                  task_id: Optional[str] = None) -> str:
        """
        Queue a task to be executed in the background.
        
        Args:
            task_func: The function to execute
            args: Positional arguments to pass to the function
            kwargs: Keyword arguments to pass to the function
            task_name: Optional name for the task (for logging)
            task_id: Optional unique identifier for the task
            
        Returns:
            str: The task ID assigned to this task
        """
        args = args or ()
        kwargs = kwargs or {}
        task_name = task_name or task_func.__name__
        
        # Generate a task ID if not provided
        if task_id is None:
            task_id = f"task_{time.time()}_{id(task_func)}"
        
        # Create cancellation flag for this task
        self.task_cancellation_flags[task_id] = False
        
        self.task_queue.put((task_func, args, kwargs, task_name, task_id))
        self.logger.debug(f"Queued task: {task_name} (ID: {task_id})")
        
        # Make sure the worker is running
        if not self.running:
            self.start()
            
        return task_id

    def cancel_task(self, task_id: str) -> bool:
        """
        Mark a specific task for cancellation.
        
        Args:
            task_id: The ID of the task to cancel
            
        Returns:
            bool: True if the task was found and marked for cancellation, False otherwise
        """
        if task_id in self.task_cancellation_flags:
            self.task_cancellation_flags[task_id] = True
            self.logger.debug(f"Task marked for cancellation: {task_id}")
            return True
        return False
    
    def is_task_cancelled(self, task_id: str) -> bool:
        """
        Check if a task has been marked for cancellation.
        
        Args:
            task_id: The ID of the task to check
            
        Returns:
            bool: True if the task is marked for cancellation, False otherwise
        """
        return self.task_cancellation_flags.get(task_id, False)
    
    def clear_cancelled_tasks(self) -> None:
        """
        Remove all cancelled tasks from the queue.
        
        Returns:
            None
        """
        new_queue = queue.Queue()
        cancelled_count = 0
        
        # Move all non-cancelled tasks to a new queue
        while not self.task_queue.empty():
            try:
                task = self.task_queue.get(block=False)
                task_id = task[4]  # Extract task_id from tuple
                
                if not self.is_task_cancelled(task_id):
                    new_queue.put(task)
                else:
                    cancelled_count += 1
                    self.logger.debug(f"Removed cancelled task from queue: {task[3]} (ID: {task_id})")
            except queue.Empty:
                break
        
        self.task_queue = new_queue
        self.logger.debug(f"Cleared {cancelled_count} cancelled tasks from queue")

    def _process_queue(self) -> None:
        """
        Process tasks from the queue until stopped.
        
        Returns:
            None
        """
        while self.running:
            try:
                # Get a task from the queue with a timeout
                try:
                    task_func, args, kwargs, task_name, task_id = self.task_queue.get(timeout=0.5)
                except queue.Empty:
                    continue
                
                self.current_task = task_name
                self.logger.info(f"Starting task: {task_name} (ID: {task_id})")
                
                # Send starting message
                if self.on_message:
                    self.on_message(f"Starting task: {task_name}", "info")
                
                # Create a progress callback that will throttle updates
                def progress_callback(value: int, message: Optional[str] = None) -> None:
                    """Throttled progress callback to avoid GUI freezing."""
                    now = time.time()
                    if now - self.last_progress_time >= self.progress_throttle:
                        if self.on_progress:
                            self.on_progress(value, message)
                        self.last_progress_time = now
                
                # Add a cancellation check function
                def check_cancelled() -> bool:
                    """Check if this task has been cancelled."""
                    return self.is_task_cancelled(task_id)
                
                # Execute the task
                result = None
                error = None
                success = False
                
                try:
                    # Add progress_callback to kwargs if task supports it
                    if 'progress_callback' in kwargs or task_func.__code__.co_varnames.count('progress_callback') > 0:
                        kwargs['progress_callback'] = progress_callback
                    
                    # Add cancellation_check to kwargs if task supports it
                    if 'cancellation_check' in kwargs or task_func.__code__.co_varnames.count('cancellation_check') > 0:
                        kwargs['cancellation_check'] = check_cancelled
                    
                    # Execute the task
                    start_time = time.time()
                    result = task_func(*args, **kwargs)
                    elapsed_time = time.time() - start_time
                    success = True
                    
                    self.logger.info(f"Task completed: {task_name} in {elapsed_time:.2f}s")
                    if self.on_message:
                        self.on_message(f"Task completed: {task_name} in {elapsed_time:.2f}s", "info")
                        
                except Exception as e:
                    elapsed_time = time.time() - start_time
                    error = e
                    self.logger.error(f"Task failed: {task_name} after {elapsed_time:.2f}s - {str(e)}")
                    self.logger.error(traceback.format_exc())
                    if self.on_message:
                        self.on_message(f"Task failed: {task_name} - {str(e)}", "error")
                
                finally:
                    # Clean up cancellation flag
                    if task_id in self.task_cancellation_flags:
                        del self.task_cancellation_flags[task_id]
                    
                    # Mark the task as done
                    self.task_queue.task_done()
                    self.current_task = None
                    
                    # Send completion callback
                    if self.on_complete:
                        self.on_complete(success, result, error)
                    
                    # Send final progress update
                    if self.on_progress and success:
                        self.on_progress(100, "Completed")
            
            except Exception as e:
                # Log any unexpected errors in the worker thread
                self.logger.error(f"Unexpected error in worker thread: {str(e)}")
                self.logger.error(traceback.format_exc())
                
                if self.on_message:
                    self.on_message(f"Internal worker error: {str(e)}", "error")
        
        self.logger.debug("Worker thread exiting")


class ProgressTracker:
    """
    Tracks progress across multiple sequential tasks.
    
    This class allows dividing total progress (0-100) across multiple 
    sequential operations, each with their own progress range.
    """
    
    def __init__(self, stages: List[Tuple[str, int]], 
                on_progress: Optional[Callable[[int, str], None]] = None):
        """
        Initialize the progress tracker.
        
        Args:
            stages: List of (stage_name, weight) tuples, where weight is the 
                  relative importance of each stage in the overall progress
            on_progress: Callback function for progress updates (value, message)
        """
        self.stages = stages
        self.on_progress = on_progress
        self.current_stage = 0
        self.total_weight = sum(weight for _, weight in stages)
        self.completed_weight = 0
    
    def next_stage(self) -> None:
        """
        Move to the next stage.
        
        Returns:
            None
        """
        if self.current_stage < len(self.stages):
            _, weight = self.stages[self.current_stage]
            self.completed_weight += weight
            self.current_stage += 1
    
    def update(self, stage_progress: int, message: Optional[str] = None) -> None:
        """
        Update the progress for the current stage.
        
        Args:
            stage_progress: Progress within the current stage (0-100)
            message: Optional message to display
            
        Returns:
            None
        """
        if self.current_stage < len(self.stages):
            stage_name, stage_weight = self.stages[self.current_stage]
            
            # Calculate overall progress
            stage_contribution = (stage_progress / 100.0) * stage_weight
            overall_progress = int(
                ((self.completed_weight + stage_contribution) / self.total_weight) * 100
            )
            
            # Ensure progress is bounded
            overall_progress = max(0, min(100, overall_progress))
            
            # Update display message
            if message is None:
                if stage_progress == 100:
                    display_message = f"Completed: {stage_name}"
                else:
                    display_message = f"{stage_name}: {stage_progress}%"
            else:
                display_message = message
            
            # Report progress
            if self.on_progress:
                self.on_progress(overall_progress, display_message)


class TaskManager:
    """
    Manages multiple background tasks with dependencies.
    
    This class allows for defining a workflow of tasks with dependencies,
    where some tasks can only start after others have completed.
    """
    
    def __init__(self, worker: BackgroundWorker, logger: Optional[logging.Logger] = None):
        """
        Initialize the task manager.
        
        Args:
            worker: BackgroundWorker instance to execute tasks
            logger: Logger instance for logging
        """
        self.worker = worker
        self.logger = logger or logging.getLogger(__name__)
        self.tasks = {}  # task_id -> (task_func, args, kwargs, dependencies)
        self.results = {}  # task_id -> result
        self.status = {}  # task_id -> "pending", "running", "completed", "failed"
        self.dependencies = {}  # task_id -> [dependent_task_ids]
        self.lock = threading.Lock()  # For thread-safe operations on shared data
        
    def add_task(self, 
                task_id: str, 
                task_func: Callable[..., Any], 
                args: Optional[Tuple] = None, 
                kwargs: Optional[Dict[str, Any]] = None,
                dependencies: Optional[List[str]] = None) -> None:
        """
        Add a task to the workflow.
        
        Args:
            task_id: Unique identifier for the task
            task_func: The function to execute
            args: Positional arguments to pass to the function
            kwargs: Keyword arguments to pass to the function
            dependencies: List of task IDs that must complete before this task can run
            
        Returns:
            None
        """
        args = args or ()
        kwargs = kwargs or {}
        dependencies = dependencies or []
        
        with self.lock:
            self.tasks[task_id] = (task_func, args, kwargs)
            self.status[task_id] = "pending"
            
            # Register dependencies
            for dep_id in dependencies:
                if dep_id not in self.dependencies:
                    self.dependencies[dep_id] = []
                self.dependencies[dep_id].append(task_id)
                
    def on_task_complete(self, task_id: str, success: bool, result: Any, error: Optional[Exception]) -> None:
        """
        Handle task completion and trigger dependent tasks.
        
        Args:
            task_id: ID of the completed task
            success: Whether the task completed successfully
            result: Result of the task
            error: Any error that occurred during task execution
            
        Returns:
            None
        """
        with self.lock:
            if success:
                self.results[task_id] = result
                self.status[task_id] = "completed"
                self.logger.debug(f"Task completed successfully: {task_id}")
                
                # Check if any dependent tasks can now run
                if task_id in self.dependencies:
                    for dependent_id in self.dependencies[task_id]:
                        self._check_and_queue_task(dependent_id)
            else:
                self.status[task_id] = "failed"
                self.logger.error(f"Task failed: {task_id}, Error: {error}")
                
    def _check_and_queue_task(self, task_id: str) -> None:
        """
        Check if all dependencies for a task are met and queue it if they are.
        
        Args:
            task_id: ID of the task to check
            
        Returns:
            None
        """
        # Check if all dependencies are completed
        can_run = True
        for dep_id, deps in self.dependencies.items():
            if task_id in deps and self.status.get(dep_id) != "completed":
                can_run = False
                break
                
        if can_run and self.status.get(task_id) == "pending":
            task_func, args, kwargs = self.tasks[task_id]
            self.status[task_id] = "running"
            
            # Add result delivery from dependencies
            for dep_id, deps in self.dependencies.items():
                if task_id in deps and dep_id in self.results:
                    # Pass dependency results in kwargs with key = dependency_task_id_result
                    kwargs[f"{dep_id}_result"] = self.results[dep_id]
            
            # Queue the task
            self.worker.queue_task(
                task_func, 
                args, 
                kwargs, 
                task_name=task_id,
                task_id=task_id
            )
            self.logger.debug(f"Queued dependent task: {task_id}")
    
    def execute_workflow(self) -> None:
        """
        Start executing the workflow by queuing tasks with no dependencies.
        
        Returns:
            None
        """
        with self.lock:
            # Find all tasks without dependencies
            for task_id in self.tasks.keys():
                has_dependencies = False
                for deps in self.dependencies.values():
                    if task_id in deps:
                        has_dependencies = True
                        break
                        
                if not has_dependencies and self.status.get(task_id) == "pending":
                    task_func, args, kwargs = self.tasks[task_id]
                    self.status[task_id] = "running"
                    
                    # Create a wrapped function that will call our callback
                    def wrapped_task(*w_args, **w_kwargs):
                        try:
                            result = task_func(*w_args, **w_kwargs)
                            self.on_task_complete(task_id, True, result, None)
                            return result
                        except Exception as e:
                            self.on_task_complete(task_id, False, None, e)
                            raise
                    
                    # Queue the wrapped task
                    self.worker.queue_task(
                        wrapped_task, 
                        args, 
                        kwargs, 
                        task_name=task_id,
                        task_id=task_id
                    )
                    self.logger.debug(f"Queued initial task: {task_id}")
    
    def cancel_workflow(self) -> None:
        """
        Cancel all pending and running tasks in the workflow.
        
        Returns:
            None
        """
        with self.lock:
            for task_id, status in self.status.items():
                if status in ("pending", "running"):
                    self.worker.cancel_task(task_id)
                    self.status[task_id] = "cancelled"
            
            self.logger.debug("Workflow cancelled")
    
    def get_workflow_status(self) -> Dict[str, str]:
        """
        Get the current status of all tasks in the workflow.
        
        Returns:
            Dict[str, str]: Dictionary of task IDs to their status
        """
        with self.lock:
            return self.status.copy()