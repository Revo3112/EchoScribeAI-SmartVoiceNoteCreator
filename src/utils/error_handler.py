# -*- coding: utf-8 -*-
"""
Error Handling and Logging Utilities for EchoScribe AI
Centralized error handling and logging configuration.
"""

import logging
import logging.handlers
import sys
import traceback
import time
from pathlib import Path
from typing import Optional, Callable, Dict, Any
import threading

class ErrorHandler:
    """
    Centralized error handling with retry logic and user notifications.
    """

    def __init__(self, status_callback: Optional[Callable[[str], None]] = None):
        self.status_callback = status_callback or (lambda x: None)
        self.error_counts: Dict[str, int] = {}
        self.last_errors: Dict[str, float] = {}
        self.max_retry = 3
        self.retry_delays = {
            'network': [2, 5, 10],  # Exponential backoff for network errors
            'api_error': [5, 15, 60],  # Longer delays for API errors
            'transcription': [1, 3, 5],  # Quick retries for transcription
            'audio': [1, 2, 5],  # Audio device errors
            'file': [0.5, 1, 2]  # File I/O errors
        }

    def handle_error(self, error_type: str, exception: Exception, operation: str = None,
                    retry_func: Optional[Callable] = None, retry_args: tuple = None,
                    retry_kwargs: Dict[str, Any] = None) -> bool:
        """
        Handle error with centralized logic and optional retry.

        Args:
            error_type: Category of error (network, api_error, transcription, etc.)
            exception: The exception that was raised
            operation: Description of the failed operation
            retry_func: Function to retry if appropriate
            retry_args: Arguments for retry function
            retry_kwargs: Keyword arguments for retry function

        Returns:
            bool: True if retry was attempted, False otherwise
        """
        current_time = time.time()

        # Update error tracking
        if error_type not in self.error_counts:
            self.error_counts[error_type] = 1
            self.last_errors[error_type] = current_time
        else:
            self.error_counts[error_type] += 1
            self.last_errors[error_type] = current_time

        # Log the error
        error_msg = f"Error [{error_type}]"
        if operation:
            error_msg += f" in {operation}"
        error_msg += f": {exception}"

        logging.error(error_msg, exc_info=self.error_counts[error_type] <= 2)

        # Determine if retry is appropriate
        should_retry = False
        if retry_func and self.error_counts[error_type] <= self.max_retry:
            retry_delay = self._get_retry_delay(error_type, self.error_counts[error_type])
            should_retry = True

            # Schedule retry
            if retry_delay > 0:
                self.status_callback(f"Error: {str(exception)[:50]}... Retrying in {retry_delay}s")
                threading.Timer(retry_delay, self._execute_retry,
                              args=(retry_func, retry_args or (), retry_kwargs or {})).start()
            else:
                self._execute_retry(retry_func, retry_args or (), retry_kwargs or {})
        else:
            # Show final error message
            self.status_callback(f"Error: {str(exception)[:100]}")

        return should_retry

    def _get_retry_delay(self, error_type: str, attempt: int) -> float:
        """Get retry delay for specific error type and attempt number."""
        delays = self.retry_delays.get(error_type, [1, 2, 5])
        if attempt <= len(delays):
            return delays[attempt - 1]
        return delays[-1]  # Use last delay for subsequent attempts

    def _execute_retry(self, retry_func: Callable, retry_args: tuple, retry_kwargs: Dict[str, Any]):
        """Execute retry function with proper error handling."""
        try:
            retry_func(*retry_args, **retry_kwargs)
        except Exception as e:
            logging.error(f"Retry failed: {e}")

    def get_user_friendly_message(self, error_type: str, exception: Exception) -> Dict[str, str]:
        """
        Generate user-friendly error messages with actionable solutions.
        Integrated from monolithic system (lines 3200-3450).
        """
        error_mapping = {
            'network': {
                'title': '🌐 Network Connection Error',
                'message': 'Unable to connect to AI services. Please check your internet connection.',
                'solutions': [
                    'Check your internet connection',
                    'Try again in a few moments',
                    'Check if your firewall is blocking the application',
                    'Verify your API key is correct'
                ]
            },
            'api_error': {
                'title': '🔑 API Service Error',
                'message': 'Error with AI transcription service.',
                'solutions': [
                    'Verify your API key is valid',
                    'Check if you have remaining API credits',
                    'Try switching to economic mode',
                    'Contact Groq support if the issue persists'
                ]
            },
            'transcription': {
                'title': '🎙️ Transcription Error',
                'message': 'Failed to process audio for transcription.',
                'solutions': [
                    'Ensure audio file is not corrupted',
                    'Try recording shorter segments',
                    'Check if audio format is supported',
                    'Reduce background noise if possible'
                ]
            },
            'audio': {
                'title': '🔊 Audio System Error',
                'message': 'Problem with audio recording or playback.',
                'solutions': [
                    'Check if microphone is connected and working',
                    'Try switching audio devices in settings',
                    'Close other applications using audio',
                    'Restart the application'
                ]
            },
            'file': {
                'title': '📁 File Operation Error',
                'message': 'Error saving or loading files.',
                'solutions': [
                    'Check if you have write permissions',
                    'Ensure sufficient disk space',
                    'Try saving to a different location',
                    'Close the file if it\'s open in another program'
                ]
            },
            'device': {
                'title': '🎧 Device Configuration Error',
                'message': 'Audio device configuration issue.',
                'solutions': [
                    'Check device connections',
                    'Update audio drivers',
                    'Try selecting different audio device',
                    'Restart Windows audio service'
                ]
            }
        }

        # Get specific error info or use generic
        error_info = error_mapping.get(error_type, {
            'title': '⚠️ Application Error',
            'message': f'An unexpected error occurred: {str(exception)[:100]}',
            'solutions': [
                'Try the operation again',
                'Restart the application',
                'Check the log file for details',
                'Report this issue if it persists'
            ]
        })

        # Add specific details based on exception type
        if 'api' in str(exception).lower() or 'groq' in str(exception).lower():
            error_info['details'] = 'API service may be temporarily unavailable'
        elif 'permission' in str(exception).lower():
            error_info['details'] = 'File permission issue detected'
        elif 'network' in str(exception).lower() or 'connection' in str(exception).lower():
            error_info['details'] = 'Network connectivity issue'
        elif 'audio' in str(exception).lower() or 'device' in str(exception).lower():
            error_info['details'] = 'Audio system configuration issue'

        return error_info

    def create_fallback_strategy(self, error_type: str, context: Dict[str, Any] = None) -> Dict[str, Any]:
        """
        Create fallback strategies for different types of failures.
        Integrated from monolithic system (lines 3450-3600).
        """
        context = context or {}

        fallback_strategies = {
            'transcription': {
                'primary': 'Use alternative transcription method',
                'secondary': 'Use local speech recognition if available',
                'fallback': 'Manual text entry mode',
                'actions': [
                    {'type': 'switch_method', 'target': 'whisper-large-v3'},
                    {'type': 'reduce_quality', 'params': {'sample_rate': 16000}},
                    {'type': 'manual_mode', 'message': 'Switch to manual text entry'}
                ]
            },
            'ai_enhancement': {
                'primary': 'Use economic AI model',
                'secondary': 'Basic text formatting only',
                'fallback': 'No AI enhancement',
                'actions': [
                    {'type': 'switch_model', 'target': 'llama3-8b-8192'},
                    {'type': 'local_enhancement', 'method': 'basic_formatting'},
                    {'type': 'bypass_ai', 'message': 'Continue without AI enhancement'}
                ]
            },
            'audio_recording': {
                'primary': 'Switch to different audio device',
                'secondary': 'Use system audio only',
                'fallback': 'Manual audio file import',
                'actions': [
                    {'type': 'switch_device', 'scan_alternatives': True},
                    {'type': 'change_method', 'target': 'system_audio'},
                    {'type': 'file_import', 'message': 'Import audio file instead'}
                ]
            },
            'file_operations': {
                'primary': 'Try alternative save location',
                'secondary': 'Use temporary directory',
                'fallback': 'Copy to clipboard',
                'actions': [
                    {'type': 'change_location', 'target': 'user_documents'},
                    {'type': 'use_temp', 'cleanup': True},
                    {'type': 'clipboard_copy', 'message': 'Content copied to clipboard'}
                ]
            },
            'network': {
                'primary': 'Retry with exponential backoff',
                'secondary': 'Use cached/offline mode',
                'fallback': 'Local processing only',
                'actions': [
                    {'type': 'exponential_retry', 'max_attempts': 5},
                    {'type': 'offline_mode', 'features': ['local_enhancement']},
                    {'type': 'local_only', 'message': 'Continue in offline mode'}
                ]
            }
        }

        strategy = fallback_strategies.get(error_type, {
            'primary': 'Retry operation',
            'secondary': 'Use default settings',
            'fallback': 'Skip this step',
            'actions': [
                {'type': 'simple_retry', 'attempts': 2},
                {'type': 'use_defaults'},
                {'type': 'skip', 'message': 'Operation skipped due to error'}
            ]
        })

        # Customize strategy based on context
        if context.get('is_critical', False):
            strategy['priority'] = 'high'
            strategy['user_intervention'] = True

        if context.get('offline_mode', False):
            # Remove network-dependent actions
            strategy['actions'] = [action for action in strategy['actions']
                                 if action.get('type') not in ['network_retry', 'api_call']]

        return strategy

    def execute_fallback_action(self, action: Dict[str, Any], context: Dict[str, Any] = None) -> bool:
        """
        Execute a specific fallback action.
        Integrated from monolithic system recovery logic.
        """
        try:
            action_type = action.get('type')

            if action_type == 'switch_method':
                return self._switch_transcription_method(action.get('target'), context)
            elif action_type == 'switch_model':
                return self._switch_ai_model(action.get('target'), context)
            elif action_type == 'switch_device':
                return self._switch_audio_device(action.get('scan_alternatives', False), context)
            elif action_type == 'change_location':
                return self._change_save_location(action.get('target'), context)
            elif action_type == 'offline_mode':
                return self._enable_offline_mode(action.get('features', []), context)
            elif action_type == 'clipboard_copy':
                return self._copy_to_clipboard(context)
            elif action_type == 'simple_retry':
                return self._simple_retry(action.get('attempts', 1), context)
            else:
                logging.warning(f"Unknown fallback action: {action_type}")
                return False

        except Exception as e:
            logging.error(f"Error executing fallback action {action_type}: {e}")
            return False

    def _switch_transcription_method(self, target_method: str, context: Dict[str, Any]) -> bool:
        """Switch to alternative transcription method."""
        try:
            if context and 'ai_service' in context:
                ai_service = context['ai_service']
                # Update transcription method if service supports it
                if hasattr(ai_service, 'fallback_to_method'):
                    return ai_service.fallback_to_method(target_method)
            return False
        except Exception as e:
            logging.error(f"Error switching transcription method: {e}")
            return False

    def _switch_ai_model(self, target_model: str, context: Dict[str, Any]) -> bool:
        """Switch to alternative AI model."""
        try:
            if context and 'ai_service' in context:
                ai_service = context['ai_service']
                if hasattr(ai_service, 'switch_model'):
                    return ai_service.switch_model(target_model)
            return False
        except Exception as e:
            logging.error(f"Error switching AI model: {e}")
            return False

    def _switch_audio_device(self, scan_alternatives: bool, context: Dict[str, Any]) -> bool:
        """Switch to alternative audio device."""
        try:
            if context and 'audio_recorder' in context:
                recorder = context['audio_recorder']
                if hasattr(recorder, 'switch_to_alternative_device'):
                    return recorder.switch_to_alternative_device(scan_alternatives)
            return False
        except Exception as e:
            logging.error(f"Error switching audio device: {e}")
            return False

    def _change_save_location(self, target_location: str, context: Dict[str, Any]) -> bool:
        """Change file save location."""
        try:
            if context and 'document_processor' in context:
                processor = context['document_processor']
                if hasattr(processor, 'change_save_location'):
                    return processor.change_save_location(target_location)
            return False
        except Exception as e:
            logging.error(f"Error changing save location: {e}")
            return False

    def _enable_offline_mode(self, features: list, context: Dict[str, Any]) -> bool:
        """Enable offline mode with specified features."""
        try:
            self.status_callback("⚡ Switching to offline mode...")
            # This would integrate with app controller to disable network features
            return True
        except Exception as e:
            logging.error(f"Error enabling offline mode: {e}")
            return False

    def _copy_to_clipboard(self, context: Dict[str, Any]) -> bool:
        """Copy content to clipboard as fallback."""
        try:
            import tkinter as tk
            if context and 'content' in context:
                root = tk.Tk()
                root.withdraw()
                root.clipboard_clear()
                root.clipboard_append(context['content'])
                root.update()
                root.destroy()
                self.status_callback("📋 Content copied to clipboard")
                return True
            return False
        except Exception as e:
            logging.error(f"Error copying to clipboard: {e}")
            return False

    def _simple_retry(self, attempts: int, context: Dict[str, Any]) -> bool:
        """Perform simple retry operation."""
        try:
            if context and 'retry_function' in context:
                retry_func = context['retry_function']
                for attempt in range(attempts):
                    try:
                        if callable(retry_func):
                            retry_func()
                            return True
                    except Exception as e:
                        if attempt == attempts - 1:
                            raise e
                        time.sleep(0.5 * (attempt + 1))  # Progressive delay
            return False
        except Exception as e:
            logging.error(f"Error in simple retry: {e}")
            return False

    def log_system_info(self):
        """
        Log comprehensive system information for debugging.
        Integrated from monolithic system (lines 3600+).
        """
        try:
            import platform
            import psutil
            import subprocess

            logging.info("=== SYSTEM INFORMATION ===")
            logging.info(f"OS: {platform.system()} {platform.release()} {platform.version()}")
            logging.info(f"Python: {platform.python_version()}")
            logging.info(f"Architecture: {platform.architecture()}")

            # Memory info
            memory = psutil.virtual_memory()
            logging.info(f"Memory: {memory.total // (1024**3)}GB total, {memory.available // (1024**3)}GB available")

            # Disk space
            disk = psutil.disk_usage('/')
            logging.info(f"Disk: {disk.total // (1024**3)}GB total, {disk.free // (1024**3)}GB free")

            # Audio devices (if available)
            try:
                import soundcard as sc
                devices = sc.all_microphones()
                logging.info(f"Audio devices: {[d.name for d in devices]}")
            except:
                logging.info("Audio device info not available")

            # FFmpeg check
            try:
                result = subprocess.run(['ffmpeg', '-version'],
                                      capture_output=True, text=True, timeout=5)
                if result.returncode == 0:
                    version = result.stdout.split('\n')[0]
                    logging.info(f"FFmpeg: {version}")
                else:
                    logging.warning("FFmpeg not properly installed")
            except:
                logging.warning("FFmpeg not found")

            logging.info("=== END SYSTEM INFO ===")

        except Exception as e:
            logging.error(f"Error logging system info: {e}")

    def reset_error_counts(self, error_type: str = None):
        """Reset error counts for specific type or all types."""
        if error_type:
            self.error_counts.pop(error_type, None)
            self.last_errors.pop(error_type, None)
        else:
            self.error_counts.clear()
            self.last_errors.clear()

    def get_error_statistics(self) -> Dict[str, Any]:
        """Get error statistics for monitoring."""
        return {
            'error_counts': self.error_counts.copy(),
            'last_errors': self.last_errors.copy(),
            'total_errors': sum(self.error_counts.values()),
            'error_types': list(self.error_counts.keys())
        }

    def reset_error_count(self, error_type: str):
        """Reset error count for a specific error type."""
        if error_type in self.error_counts:
            del self.error_counts[error_type]
        if error_type in self.last_errors:
            del self.last_errors[error_type]

    def get_error_stats(self) -> Dict[str, Any]:
        """Get error statistics."""
        return {
            "error_counts": self.error_counts.copy(),
            "last_errors": self.last_errors.copy(),
            "total_errors": sum(self.error_counts.values())
        }

def setup_logging(log_level: str = "INFO", log_to_file: bool = True,
                 log_dir: Optional[Path] = None) -> logging.Logger:
    """
    Setup centralized logging configuration.

    Args:
        log_level: Logging level (DEBUG, INFO, WARNING, ERROR)
        log_to_file: Whether to log to file
        log_dir: Directory for log files (defaults to ~/.echoscribe/logs)

    Returns:
        Configured logger
    """
    # Create log directory
    if log_dir is None:
        log_dir = Path.home() / ".echoscribe" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    # Configure root logger
    logger = logging.getLogger()
    logger.setLevel(getattr(logging, log_level.upper()))

    # Clear existing handlers
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%H:%M:%S'
    )
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

    # File handler
    if log_to_file:
        log_file = log_dir / "echoscribe.log"
        file_handler = logging.handlers.RotatingFileHandler(
            log_file, maxBytes=10*1024*1024, backupCount=5, encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        file_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s'
        )
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)

    # Error file handler for critical errors
    if log_to_file:
        error_log_file = log_dir / "errors.log"
        error_handler = logging.handlers.RotatingFileHandler(
            error_log_file, maxBytes=5*1024*1024, backupCount=3, encoding='utf-8'
        )
        error_handler.setLevel(logging.ERROR)
        error_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d\n%(message)s\n%(exc_info)s\n' + '-'*80
        )
        error_handler.setFormatter(error_formatter)
        logger.addHandler(error_handler)

    # Suppress some noisy loggers
    logging.getLogger('matplotlib').setLevel(logging.WARNING)
    logging.getLogger('PIL').setLevel(logging.WARNING)
    logging.getLogger('urllib3').setLevel(logging.WARNING)

    logger.info("Logging system initialized")
    return logger

class ContextManager:
    """Context manager for safe resource handling with proper cleanup."""

    def __init__(self, resource_name: str, cleanup_func: Optional[Callable] = None):
        self.resource_name = resource_name
        self.cleanup_func = cleanup_func
        self.logger = logging.getLogger(__name__)

    def __enter__(self):
        self.logger.debug(f"Acquiring resource: {self.resource_name}")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.logger.debug(f"Releasing resource: {self.resource_name}")
        if self.cleanup_func:
            try:
                self.cleanup_func()
            except Exception as e:
                self.logger.error(f"Error during cleanup of {self.resource_name}: {e}")

        if exc_type:
            self.logger.error(f"Exception in {self.resource_name}: {exc_val}")

        return False  # Don't suppress exceptions

def safe_execute(func: Callable, *args, error_handler: Optional[ErrorHandler] = None,
                error_type: str = "general", operation: str = None, **kwargs) -> Any:
    """
    Safely execute a function with error handling.

    Args:
        func: Function to execute
        *args: Arguments for the function
        error_handler: ErrorHandler instance
        error_type: Type of error for categorization
        operation: Description of the operation
        **kwargs: Keyword arguments for the function

    Returns:
        Function result or None if error occurred
    """
    try:
        return func(*args, **kwargs)
    except Exception as e:
        if error_handler:
            error_handler.handle_error(error_type, e, operation)
        else:
            logging.error(f"Error in {operation or 'operation'}: {e}", exc_info=True)
        return None

def format_error_for_user(exception: Exception, operation: str = None) -> str:
    """
    Format an error message for user display.

    Args:
        exception: The exception to format
        operation: Description of the failed operation

    Returns:
        User-friendly error message
    """
    error_type = type(exception).__name__
    error_msg = str(exception)

    # Map technical errors to user-friendly messages
    error_mappings = {
        'ConnectionError': 'Koneksi internet bermasalah',
        'TimeoutError': 'Operasi timeout - coba lagi',
        'FileNotFoundError': 'File tidak ditemukan',
        'PermissionError': 'Tidak ada izin akses file',
        'OSError': 'Error sistem operasi',
        'ValueError': 'Data tidak valid',
        'KeyError': 'Konfigurasi tidak lengkap',
        'ImportError': 'Library yang diperlukan tidak tersedia',
        'ModuleNotFoundError': 'Module tidak ditemukan'
    }

    user_msg = error_mappings.get(error_type, error_msg)

    if operation:
        return f"Error saat {operation}: {user_msg}"
    else:
        return f"Error: {user_msg}"

def log_system_info():
    """Log system information for debugging."""
    import platform
    import sys

    logger = logging.getLogger(__name__)
    logger.info("=== System Information ===")
    logger.info(f"Platform: {platform.platform()}")
    logger.info(f"Python: {sys.version}")
    logger.info(f"Architecture: {platform.architecture()}")
    logger.info(f"Processor: {platform.processor()}")
    logger.info("==========================")
