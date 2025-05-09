"""
Test module for the background worker implementation.

This module provides tests for the BackgroundWorker, ProgressTracker, and TaskManager
classes, demonstrating their usage in a simulated Excel operation workflow.
"""

import os
import sys
import time
import logging
import unittest
from unittest.mock import MagicMock, patch

# Add the src directory to the path for imports
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.uco_to_udo_recon.modules.background_worker import BackgroundWorker, ProgressTracker, TaskManager


# Sample functions for testing
def long_running_task(duration=1, progress_callback=None, cancellation_check=None):
    """A sample long-running task that reports progress."""
    start_time = time.time()
    end_time = start_time + duration
    
    while time.time() < end_time:
        # Check for cancellation if provided
        if cancellation_check and cancellation_check():
            raise Exception("Task was cancelled")
            
        # Calculate progress percentage
        elapsed = time.time() - start_time
        progress = min(int((elapsed / duration) * 100), 100)
        
        # Report progress if callback provided
        if progress_callback:
            progress_callback(progress, f"Processing... {progress}%")
            
        # Small sleep to avoid tight loop
        time.sleep(0.05)
        
    return f"Task completed in {duration} seconds"


def task_with_error():
    """A task that raises an exception."""
    raise ValueError("Simulated task error")


class TestBackgroundWorker(unittest.TestCase):
    """Test cases for the BackgroundWorker class."""
    
    def setUp(self):
        """Set up test environment."""
        self.progress_mock = MagicMock()
        self.complete_mock = MagicMock()
        self.message_mock = MagicMock()
        self.logger = logging.getLogger("test_logger")
        
        # Create worker with mocked callbacks
        self.worker = BackgroundWorker(
            on_progress=self.progress_mock,
            on_complete=self.complete_mock,
            on_message=self.message_mock,
            logger=self.logger
        )
        
        # Start the worker thread
        self.worker.start()
        
    def tearDown(self):
        """Clean up after tests."""
        # Stop the worker thread
        self.worker.stop()
        
    def test_successful_task(self):
        """Test executing a successful task."""
        # Queue a short task
        self.worker.queue_task(long_running_task, args=(0.5,))
        
        # Wait for it to complete (with timeout)
        timeout = time.time() + 2
        while self.complete_mock.call_count == 0 and time.time() < timeout:
            time.sleep(0.1)
            
        # Verify callbacks were called
        self.assertTrue(self.progress_mock.call_count > 0, "Progress callback should be called")
        self.complete_mock.assert_called_once()
        self.message_mock.assert_called()
        
        # Verify successful result
        success, result, error = self.complete_mock.call_args[0]
        self.assertTrue(success)
        self.assertIsNotNone(result)
        self.assertIsNone(error)
        
    def test_task_with_error(self):
        """Test executing a task that raises an error."""
        # Queue a task that will raise an exception
        self.worker.queue_task(task_with_error)
        
        # Wait for it to complete (with timeout)
        timeout = time.time() + 2
        while self.complete_mock.call_count == 0 and time.time() < timeout:
            time.sleep(0.1)
            
        # Verify error handling
        self.complete_mock.assert_called_once()
        success, result, error = self.complete_mock.call_args[0]
        self.assertFalse(success)
        self.assertIsNone(result)
        self.assertIsInstance(error, ValueError)
        self.assertEqual(str(error), "Simulated task error")
        
    def test_task_cancellation(self):
        """Test cancelling a task."""
        # Queue a long task
        task_id = self.worker.queue_task(long_running_task, args=(2,))
        
        # Wait for it to start
        time.sleep(0.2)
        
        # Cancel the task
        self.worker.cancel_task(task_id)
        
        # Wait for completion with timeout
        timeout = time.time() + 3
        while self.complete_mock.call_count == 0 and time.time() < timeout:
            time.sleep(0.1)
            
        # Verify task was cancelled
        self.complete_mock.assert_called_once()
        success, result, error = self.complete_mock.call_args[0]
        self.assertFalse(success)
        self.assertIsInstance(error, Exception)
        self.assertIn("cancelled", str(error).lower())


class TestProgressTracker(unittest.TestCase):
    """Test cases for the ProgressTracker class."""
    
    def setUp(self):
        """Set up test environment."""
        self.progress_mock = MagicMock()
        
        # Define stages with weights
        self.stages = [
            ("Stage 1", 10),
            ("Stage 2", 30),
            ("Stage 3", 60)
        ]
        
        # Create tracker
        self.tracker = ProgressTracker(self.stages, self.progress_mock)
        
    def test_progress_calculation(self):
        """Test progress calculation across stages."""
        # Stage 1 (10% weight)
        self.tracker.update(50)  # 50% of Stage 1
        self.progress_mock.assert_called_with(5, "Stage 1: 50%")  # 5% overall
        
        self.tracker.update(100)  # 100% of Stage 1
        self.progress_mock.assert_called_with(10, "Completed: Stage 1")  # 10% overall
        
        # Move to Stage 2 (30% weight)
        self.tracker.next_stage()
        self.tracker.update(50)  # 50% of Stage 2
        self.progress_mock.assert_called_with(25, "Stage 2: 50%")  # 10% + 15% = 25% overall
        
        self.tracker.update(100)
        self.progress_mock.assert_called_with(40, "Completed: Stage 2")  # 10% + 30% = 40% overall
        
        # Move to Stage 3 (60% weight)
        self.tracker.next_stage()
        self.tracker.update(50)  # 50% of Stage 3
        self.progress_mock.assert_called_with(70, "Stage 3: 50%")  # 40% + 30% = 70% overall
        
        self.tracker.update(100, "All done!")
        self.progress_mock.assert_called_with(100, "All done!")  # 100% overall with custom message


class TestTaskManager(unittest.TestCase):
    """Test cases for the TaskManager class."""
    
    def setUp(self):
        """Set up test environment."""
        self.logger = logging.getLogger("test_logger")
        self.worker = BackgroundWorker(logger=self.logger)
        self.worker.start()
        
        # Create task manager
        self.task_manager = TaskManager(self.worker, self.logger)
        
        # Define test tasks
        def task_a():
            return "Result A"
            
        def task_b(task_a_result=None):
            return f"Result B with {task_a_result}"
            
        def task_c(task_a_result=None):
            return f"Result C with {task_a_result}"
            
        def task_d(task_b_result=None, task_c_result=None):
            return f"Result D with {task_b_result} and {task_c_result}"
            
        self.task_a = task_a
        self.task_b = task_b
        self.task_c = task_c
        self.task_d = task_d
        
    def tearDown(self):
        """Clean up after tests."""
        self.worker.stop()
        
    def test_dependency_execution(self):
        """Test execution of tasks with dependencies."""
        # Add tasks with dependencies
        self.task_manager.add_task("task_a", self.task_a)
        self.task_manager.add_task("task_b", self.task_b, dependencies=["task_a"])
        self.task_manager.add_task("task_c", self.task_c, dependencies=["task_a"])
        self.task_manager.add_task("task_d", self.task_d, dependencies=["task_b", "task_c"])
        
        # Start workflow execution
        self.task_manager.execute_workflow()
        
        # Wait for all tasks to complete with timeout
        timeout = time.time() + 5
        while len([s for s in self.task_manager.get_workflow_status().values() 
                  if s in ("pending", "running")]) > 0 and time.time() < timeout:
            time.sleep(0.1)
        
        # Check results
        self.assertEqual(self.task_manager.status["task_a"], "completed")
        self.assertEqual(self.task_manager.status["task_b"], "completed")
        self.assertEqual(self.task_manager.status["task_c"], "completed")
        self.assertEqual(self.task_manager.status["task_d"], "completed")
        
        # Verify correct execution order through result dependencies
        self.assertEqual(self.task_manager.results["task_a"], "Result A")
        self.assertIn("Result A", self.task_manager.results["task_b"])
        self.assertIn("Result A", self.task_manager.results["task_c"])
        self.assertIn("Result B", self.task_manager.results["task_d"])
        self.assertIn("Result C", self.task_manager.results["task_d"])


if __name__ == '__main__':
    # Configure logging
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Run tests
    unittest.main()