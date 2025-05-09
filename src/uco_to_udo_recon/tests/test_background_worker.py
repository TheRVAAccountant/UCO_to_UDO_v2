"""
Tests for the background worker module.

This module contains tests for the BackgroundWorker and ProgressTracker classes.
"""

import unittest
import time
import threading
import logging
from unittest.mock import MagicMock, patch

# Import modules to test
from src.uco_to_udo_recon.modules.background_worker import BackgroundWorker, ProgressTracker


class TestBackgroundWorker(unittest.TestCase):
    """Tests for the BackgroundWorker class."""

    def setUp(self):
        """Set up the test environment."""
        # Set up mock callbacks
        self.on_progress = MagicMock()
        self.on_complete = MagicMock()
        self.on_message = MagicMock()
        self.logger = logging.getLogger("test_logger")
        
        # Create a worker instance
        self.worker = BackgroundWorker(
            on_progress=self.on_progress,
            on_complete=self.on_complete,
            on_message=self.on_message,
            logger=self.logger
        )
        
        # Start the worker thread
        self.worker.start()
    
    def tearDown(self):
        """Clean up after tests."""
        # Stop the worker thread
        self.worker.stop()
        
    def test_queue_and_execute_task(self):
        """Test queueing and executing a task."""
        # Define a simple test task
        def test_task(value):
            return value * 2
        
        # Create a semaphore to wait for task completion
        completion_event = threading.Event()
        
        # Override the on_complete callback to signal when done
        def on_complete_override(success, result, error):
            self.on_complete(success, result, error)
            completion_event.set()
            
        self.worker.on_complete = on_complete_override
        
        # Queue the task
        task_id = self.worker.queue_task(test_task, args=(5,))
        
        # Wait for task to complete (with timeout)
        completion_event.wait(timeout=5.0)
        
        # Check that the task completed successfully
        self.on_complete.assert_called_once()
        success, result, error = self.on_complete.call_args[0]
        self.assertTrue(success)
        self.assertEqual(result, 10)
        self.assertIsNone(error)
    
    def test_task_error_handling(self):
        """Test that errors in tasks are handled properly."""
        # Define a task that raises an exception
        def failing_task():
            raise ValueError("Test error")
        
        # Create a semaphore to wait for task completion
        completion_event = threading.Event()
        
        # Override the on_complete callback to signal when done
        def on_complete_override(success, result, error):
            self.on_complete(success, result, error)
            completion_event.set()
            
        self.worker.on_complete = on_complete_override
        
        # Queue the task
        task_id = self.worker.queue_task(failing_task)
        
        # Wait for task to complete (with timeout)
        completion_event.wait(timeout=5.0)
        
        # Check that the task failed and error was reported
        self.on_complete.assert_called_once()
        success, result, error = self.on_complete.call_args[0]
        self.assertFalse(success)
        self.assertIsNone(result)
        self.assertIsInstance(error, ValueError)
        self.assertEqual(str(error), "Test error")
    
    def test_progress_reporting(self):
        """Test that progress is reported correctly."""
        # Define a task that reports progress
        def progress_task(progress_callback):
            progress_callback(0, "Starting")
            time.sleep(0.1)
            progress_callback(50, "Halfway")
            time.sleep(0.1)
            progress_callback(100, "Done")
            return "Completed"
        
        # Create a semaphore to wait for task completion
        completion_event = threading.Event()
        
        # Override the on_complete callback to signal when done
        def on_complete_override(success, result, error):
            self.on_complete(success, result, error)
            completion_event.set()
            
        self.worker.on_complete = on_complete_override
        
        # Queue the task
        task_id = self.worker.queue_task(progress_task)
        
        # Wait for task to complete (with timeout)
        completion_event.wait(timeout=5.0)
        
        # Check progress was reported (at least once, might be more due to throttling)
        self.assertGreaterEqual(self.on_progress.call_count, 1)
        
        # Check that the task completed successfully
        self.on_complete.assert_called_once()
        success, result, error = self.on_complete.call_args[0]
        self.assertTrue(success)
        self.assertEqual(result, "Completed")
        self.assertIsNone(error)
    
    def test_task_cancellation(self):
        """Test that tasks can be cancelled."""
        # Define a task that checks cancellation
        def cancellable_task(cancellation_check):
            for i in range(10):
                if cancellation_check():
                    return "Cancelled"
                time.sleep(0.1)
            return "Completed"
        
        # Create a semaphore to wait for task completion
        completion_event = threading.Event()
        
        # Override the on_complete callback to signal when done
        def on_complete_override(success, result, error):
            self.on_complete(success, result, error)
            completion_event.set()
            
        self.worker.on_complete = on_complete_override
        
        # Queue the task
        task_id = self.worker.queue_task(cancellable_task)
        
        # Wait a moment then cancel
        time.sleep(0.2)
        self.worker.cancel_task(task_id)
        
        # Wait for task to complete (with timeout)
        completion_event.wait(timeout=5.0)
        
        # Check that the task completed with cancellation
        self.on_complete.assert_called_once()
        success, result, error = self.on_complete.call_args[0]
        self.assertTrue(success)
        self.assertEqual(result, "Cancelled")
        self.assertIsNone(error)


class TestProgressTracker(unittest.TestCase):
    """Tests for the ProgressTracker class."""
    
    def setUp(self):
        """Set up the test environment."""
        self.on_progress = MagicMock()
        
        # Define test stages
        self.stages = [
            ("Stage 1", 20),
            ("Stage 2", 30),
            ("Stage 3", 50)
        ]
        
        # Create progress tracker
        self.tracker = ProgressTracker(self.stages, self.on_progress)
    
    def test_progress_tracking(self):
        """Test that progress is tracked correctly across stages."""
        # Start at stage 0
        self.assertEqual(self.tracker.current_stage, 0)
        
        # Update progress for stage 0
        self.tracker.update(50)
        self.on_progress.assert_called_with(10, "Stage 1: 50%")  # 50% of 20% = 10%
        
        # Complete stage 0
        self.tracker.update(100, "Custom message")
        self.on_progress.assert_called_with(20, "Custom message")  # 100% of 20% = 20%
        
        # Move to stage 1
        self.tracker.next_stage()
        self.assertEqual(self.tracker.current_stage, 1)
        
        # Update progress for stage 1
        self.tracker.update(50)
        self.on_progress.assert_called_with(35, "Stage 2: 50%")  # 20% + (50% of 30%) = 35%
        
        # Complete stage 1
        self.tracker.update(100)
        self.on_progress.assert_called_with(50, "Completed: Stage 2")  # 20% + 30% = 50%
        
        # Move to stage 2
        self.tracker.next_stage()
        self.assertEqual(self.tracker.current_stage, 2)
        
        # Update progress for stage 2
        self.tracker.update(50)
        self.on_progress.assert_called_with(75, "Stage 3: 50%")  # 50% + (50% of 50%) = 75%
        
        # Complete stage 2
        self.tracker.update(100)
        self.on_progress.assert_called_with(100, "Completed: Stage 3")  # 50% + 50% = 100%


if __name__ == "__main__":
    unittest.main()