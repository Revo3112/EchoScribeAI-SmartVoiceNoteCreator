# -*- coding: utf-8 -*-
"""
Main Application Controller for EchoScribe AI
Orchestrates all modules and handles the main application logic.
"""

import logging
import threading
import time
from pathlib import Path
from typing import Optional, Dict, Any, Callable, List
import tempfile
import os

# Import our modular components
from src.config.config_manager import ConfigManager
from src.audio.recorder import AudioRecorder
from src.ai.ai_service import AIService
from src.document.processor import DocumentProcessor
from src.utils.error_handler import ErrorHandler, setup_logging

logger = logging.getLogger(__name__)

class EchoScribeApp:
    """
    Main application controller that coordinates all modules.
    This replaces the monolithic VoiceToMarkdownApp class.
    """

    def __init__(self, status_callback: Optional[Callable[[str], None]] = None):
        # Setup logging
        setup_logging()

        # Initialize status callback
        self.status_callback = status_callback or (lambda x: print(f"Status: {x}"))

        # Initialize core components
        self.config = ConfigManager()
        self.error_handler = ErrorHandler(self.status_callback)

        # Initialize services
        self.audio_recorder = None
        self.ai_service = None
        self.document_processor = None

        # Application state
        self.recording = False
        self.processing = False

        # Recording data
        self.current_audio_file: Optional[str] = None
        self.current_transcript: Optional[str] = None
        self.current_enhanced_text: Optional[str] = None

        # Initialize services
        self._initialize_services()

        logger.info("EchoScribe AI initialized successfully")

    def _initialize_services(self):
        """Initialize all service modules with proper error handling and validation."""
        try:
            # Initialize audio recorder with enhanced config
            audio_config = self.config.get_audio_config()
            self.audio_recorder = AudioRecorder(audio_config, self.status_callback)

            # Validate audio devices
            available_mics = self.audio_recorder.get_available_microphones()
            if not available_mics:
                logger.warning("No microphones detected")
                self.status_callback("Warning: No microphones detected")

            # Initialize AI service with comprehensive config
            api_key = self.config.get_user_api_key()
            ai_config = {
                'use_economic_model': self.config.get('use_economic_model', False),
                'language': self.config.get('language', 'id-ID'),
                'engine': self.config.get('engine', 'Google'),
                'use_ai_enhancement': self.config.get('use_ai_enhancement', True),
                'max_tokens': self.config.get('max_tokens', 4000),
                'api_request_delay': self.config.get('api_request_delay', 10),
                'chunk_size': self.config.get('chunk_size', 600)
            }
            self.ai_service = AIService(api_key, ai_config, self.status_callback)

            # Initialize document processor with enhanced settings
            doc_config = {
                "output_folder": self.config.get("output_folder", str(Path.home() / "Documents")),
                "filename_prefix": self.config.get("filename_prefix", "catatan"),
                "output_formats": self.config.get("output_formats", ["markdown", "word"]),
                "heading_spacing_before": self.config.get("heading_spacing_before", 12),
                "heading_spacing_after": self.config.get("heading_spacing_after", 6),
                "paragraph_spacing": self.config.get("paragraph_spacing", 6)
            }
            self.document_processor = DocumentProcessor(doc_config, self.status_callback)

            # Validate services
            services_status = {
                "audio_recorder": self.audio_recorder is not None,
                "ai_service": self.ai_service is not None and self.ai_service.is_available(),
                "document_processor": self.document_processor is not None
            }

            logger.info(f"Services initialized: {services_status}")

            # Set default recording method
            self._setup_default_recording_method()

        except Exception as e:
            logger.error(f"Service initialization failed: {e}")
            self.error_handler.handle_error("initialization", e, "initializing services")

    def _setup_default_recording_method(self):
        """Setup default recording method based on platform and available devices."""
        try:
            recording_method = self.config.get("recording_method", "microphone")

            # Validate and set recording method
            if recording_method == "system":
                if not self.audio_recorder.test_device("system", 0):
                    logger.warning("System audio not available, falling back to microphone")
                    recording_method = "microphone"
                    self.config.set("recording_method", "microphone")

            elif recording_method == "dual":
                if not (self.audio_recorder.test_device("system", 0) and
                       self.audio_recorder.test_device("microphone", 0)):
                    logger.warning("Dual recording not available, falling back to microphone")
                    recording_method = "microphone"
                    self.config.set("recording_method", "microphone")

            logger.info(f"Recording method set to: {recording_method}")

        except Exception as e:
            logger.error(f"Failed to setup recording method: {e}")

    # =============================================================================
    # RECORDING METHODS (Based on monolithic system)
    # =============================================================================

    def start_recording(self, method: str = None) -> bool:
        """
        Start recording using specified method.
        Methods: 'microphone', 'system', 'dual'
        """
        if self.recording:
            self.status_callback("Recording is already in progress")
            return False

        try:
            method = method or self.config.get("recording_method", "microphone")
            self.recording = True
            self.current_audio_file = None
            self.current_transcript = None
            self.current_enhanced_text = None

            self.status_callback(f"Starting {method} recording...")
            logger.info(f"Starting {method} recording")

            # Route to appropriate recording method
            if method == "microphone":
                result = self.audio_recorder.record_microphone_audio()
            elif method == "system":
                result = self.audio_recorder.record_system_audio()
            elif method == "dual":
                result = self.audio_recorder.record_dual_audio()
            else:
                raise ValueError(f"Unknown recording method: {method}")

            if result:
                self.current_audio_file = result
                self.status_callback("Recording completed successfully")
                return True
            else:
                self.status_callback("Recording failed")
                return False

        except Exception as e:
            logger.error(f"Recording failed: {e}")
            self.error_handler.handle_error("audio", e, f"starting {method} recording")
            return False
        finally:
            self.recording = False

    def stop_recording(self) -> bool:
        """Stop current recording and return success status."""
        if not self.recording:
            self.status_callback("No recording in progress")
            return False

        try:
            self.recording = False
            self.audio_recorder.stop_recording_flag = True

            # Wait a moment for recording to cleanly stop
            time.sleep(0.5)

            self.status_callback("Recording stopped")
            logger.info("Recording stopped by user")
            return True

        except Exception as e:
            logger.error(f"Error stopping recording: {e}")
            self.error_handler.handle_error("audio", e, "stopping recording")
            return False

    def get_recording_status(self) -> Dict[str, Any]:
        """Get current recording status and metrics."""
        try:
            if self.audio_recorder:
                return {
                    "recording": self.recording,
                    "elapsed_time": getattr(self.audio_recorder, 'elapsed_time', 0),
                    "audio_file": self.current_audio_file,
                    "chunks_recorded": len(getattr(self.audio_recorder, 'temp_audio_files', [])),
                    "has_audio_data": bool(self.current_audio_file or
                                         getattr(self.audio_recorder, 'temp_audio_files', []))
                }
            return {"recording": False}
        except Exception as e:
            logger.error(f"Error getting recording status: {e}")
            return {"recording": False, "error": str(e)}

    # =============================================================================
    # PROCESSING METHODS (Based on monolithic system)
    # =============================================================================

    def process_audio_async(self) -> threading.Thread:
        """Start audio processing in a separate thread."""
        if self.processing:
            self.status_callback("Processing is already in progress")
            return None

        self.processing = True
        processing_thread = threading.Thread(
            target=self._process_audio_thread,
            daemon=True
        )
        processing_thread.start()
        return processing_thread

    def _process_audio_thread(self):
        """Main audio processing thread (based on monolithic system)."""
        try:
            self.status_callback("Starting audio processing...")

            # Check if we have audio data
            if not self.current_audio_file and not hasattr(self.audio_recorder, 'temp_audio_files'):
                self.status_callback("No audio data to process")
                return

            # Determine processing mode
            use_extended_recording = self.config.get("use_extended_recording", False)
            if use_extended_recording and hasattr(self.audio_recorder, 'temp_audio_files'):
                self._process_extended_recording()
            else:
                self._process_single_recording()

        except Exception as e:
            logger.error(f"Audio processing failed: {e}")
            self.error_handler.handle_error("processing", e, "audio processing")
        finally:
            self.processing = False

    def _process_single_recording(self):
        """Process a single audio recording."""
        self.status_callback("Starting transcription...")

        # Detect language if auto-detection is enabled
        language = self.config.get("language", "id")
        if self.config.get("auto_detect_language", False):
            self.status_callback("Detecting language...")
            quick_transcript = self.ai_service.transcribe_audio(
                self.current_audio_file, "auto", "whisper-large-v3"
            )
            if quick_transcript:
                detected_lang = self.ai_service.detect_language(quick_transcript[:500])
                language = detected_lang
                logger.info(f"Detected language: {language}")

        # Main transcription
        transcript = self.ai_service.transcribe_audio(
            self.current_audio_file,
            language,
            self._select_whisper_model()
        )

        if not transcript:
            self.status_callback("Transcription failed")
            return

        self.current_transcript = transcript
        self.status_callback("Transcription completed. Enhancing text...")

        # Enhance text if enabled
        if self.config.get("use_ai_enhancement", True):
            enhanced_text = self.ai_service.enhance_text(transcript)
            if enhanced_text:
                self.current_enhanced_text = enhanced_text
                self.status_callback("Text enhancement completed")
            else:
                self.current_enhanced_text = transcript
                self.status_callback("Enhancement failed, using raw transcript")
        else:
            self.current_enhanced_text = transcript

        # Generate documents
        self._generate_documents()

    def _process_extended_recording(self):
        """Process extended recording with multiple chunks."""
        self.status_callback("Processing extended recording...")

        chunks = getattr(self.audio_recorder, 'temp_audio_files', [])
        if not chunks:
            self.status_callback("No audio chunks found")
            return

        all_transcripts = []
        language = self.config.get("language", "id")

        for i, chunk_file in enumerate(chunks):
            self.status_callback(f"Processing chunk {i+1}/{len(chunks)}...")

            transcript = self.ai_service.transcribe_audio(
                chunk_file,
                language,
                self._select_whisper_model()
            )

            if transcript:
                all_transcripts.append(transcript)

                # Add delay between API calls
                delay = self.config.get("api_request_delay", 10)
                if i < len(chunks) - 1:  # Don't delay after last chunk
                    time.sleep(delay)

        # Combine transcripts
        combined_transcript = "\n\n".join(all_transcripts)
        self.current_transcript = combined_transcript

        # Enhance combined text
        if self.config.get("use_ai_enhancement", True):
            self.status_callback("Enhancing combined text...")
            enhanced_text = self.ai_service.enhance_text(combined_transcript)
            self.current_enhanced_text = enhanced_text or combined_transcript
        else:
            self.current_enhanced_text = combined_transcript

        # Generate documents
        self._generate_documents()

    def _select_whisper_model(self) -> str:
        """Select appropriate Whisper model based on configuration."""
        language = self.config.get("language", "id")
        use_economic = self.config.get("use_economic_model", False)

        if language.startswith("en") and use_economic:
            return "distil-whisper-large-v3-en"
        elif language.startswith("en"):
            return "whisper-large-v3"
        else:
            return "whisper-large-v3-turbo"

    def _generate_documents(self):
        """Generate output documents."""
        try:
            self.status_callback("Generating documents...")

            if not self.current_enhanced_text:
                self.status_callback("No text to generate documents")
                return

            # Generate filename
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            prefix = self.config.get("filename_prefix", "catatan")
            base_filename = f"{prefix}_{timestamp}"

            # Generate documents in configured formats
            output_formats = self.config.get("output_formats", ["word"])
            output_folder = self.config.get("output_folder", str(Path.home() / "Documents"))

            generated_files = []

            for format_type in output_formats:
                if format_type == "word":
                    doc_path = self.document_processor.create_word_document(
                        self.current_enhanced_text,
                        base_filename,
                        output_folder
                    )
                    if doc_path:
                        generated_files.append(doc_path)

                elif format_type == "markdown":
                    md_path = self.document_processor.create_markdown_document(
                        self.current_enhanced_text,
                        base_filename,
                        output_folder
                    )
                    if md_path:
                        generated_files.append(md_path)

            if generated_files:
                files_str = ", ".join([Path(f).name for f in generated_files])
                self.status_callback(f"✅ Documents created: {files_str}")
                logger.info(f"Generated documents: {generated_files}")
            else:
                self.status_callback("❌ Document generation failed")

        except Exception as e:
            logger.error(f"Document generation failed: {e}")
            self.error_handler.handle_error("document", e, "document generation")

    # =============================================================================
    # AUDIO VISUALIZATION AND MONITORING
    # =============================================================================

    def get_audio_queue(self):
        """Get audio queue for real-time visualization."""
        if self.audio_recorder:
            return self.audio_recorder.get_audio_queue()
        return None

    def set_visualization_mode(self, mode: str, sensitivity: float = 1.0):
        """Set audio visualization mode and sensitivity."""
        self.config.set("viz_mode", mode)
        self.config.set("viz_sensitivity", sensitivity)
        logger.info(f"Visualization mode set to: {mode} with sensitivity: {sensitivity}")

    def toggle_visualization(self, enabled: bool):
        """Enable or disable audio visualization."""
        self.config.set("viz_enabled", enabled)
        if self.audio_recorder:
            self.audio_recorder.set_visualization_enabled(enabled)
        logger.info(f"Audio visualization {'enabled' if enabled else 'disabled'}")

    # =============================================================================
    # CONFIGURATION MANAGEMENT
    # =============================================================================

    def update_api_key(self, api_key: str) -> bool:
        """Update the API key for AI services."""
        try:
            if self.config.save_user_api_key(api_key):
                if self.ai_service:
                    self.ai_service.update_api_key(api_key)
                self.status_callback("API key updated successfully")
                return True
            return False
        except Exception as e:
            self.error_handler.handle_error("config", e, "API key update")
            return False

    def get_config_value(self, key: str, default: Any = None) -> Any:
        """Get configuration value."""
        return self.config.get(key, default)

    def set_config_value(self, key: str, value: Any) -> None:
        """Set configuration value."""
        self.config.set(key, value)
        # Reinitialize services if audio config changed
        if key in ["sample_rate", "channels", "blocksize", "sample_width"]:
            self._reinitialize_audio_recorder()

    def _reinitialize_audio_recorder(self):
        """Reinitialize audio recorder with new config."""
        try:
            audio_config = self.config.get_audio_config()
            self.audio_recorder = AudioRecorder(audio_config, self.status_callback)
            logger.info("Audio recorder reinitialized with new config")
        except Exception as e:
            logger.error(f"Error reinitializing audio recorder: {e}")

    # =============================================================================
    # DEVICE MANAGEMENT
    # =============================================================================

    def get_available_microphones(self) -> List[Dict[str, Any]]:
        """Get list of available microphone devices."""
        if self.audio_recorder:
            return self.audio_recorder.get_available_microphones()
        return []

    def set_microphone_device(self, device_id: int) -> bool:
        """Set the microphone device."""
        if self.audio_recorder:
            return self.audio_recorder.set_microphone(device_id)
        return False

    def set_speaker_device(self, device_id: int) -> bool:
        """Set the speaker device for loopback recording."""
        if self.audio_recorder:
            return self.audio_recorder.set_speaker(device_id)
        return False

    def test_audio_device(self, device_type: str, device_id: int) -> bool:
        """Test if an audio device is working properly."""
        if self.audio_recorder:
            return self.audio_recorder.test_device(device_type, device_id)
        return False

    # =============================================================================
    # UTILITY METHODS
    # =============================================================================

    def get_current_transcript(self) -> Optional[str]:
        """Get the current transcript text."""
        return self.current_transcript

    def get_current_enhanced_text(self) -> Optional[str]:
        """Get the current enhanced text."""
        return self.current_enhanced_text

    def get_current_audio_file(self) -> Optional[str]:
        """Get the current audio file path."""
        return self.current_audio_file

    def is_processing(self) -> bool:
        """Check if processing is currently in progress."""
        return self.processing

    def is_recording(self) -> bool:
        """Check if recording is currently in progress."""
        return self.recording

    def cleanup(self):
        """Clean up resources and temporary files."""
        try:
            # Stop any ongoing operations
            if self.recording:
                self.stop_recording()

            # Clean up audio recorder
            if self.audio_recorder:
                if hasattr(self.audio_recorder, 'cleanup'):
                    self.audio_recorder.cleanup()

            # Clean up temporary files
            if hasattr(self.audio_recorder, 'temp_audio_files'):
                for temp_file in self.audio_recorder.temp_audio_files:
                    try:
                        if os.path.exists(temp_file):
                            os.remove(temp_file)
                    except Exception as e:
                        logger.warning(f"Could not remove temp file {temp_file}: {e}")

            logger.info("Application cleanup completed")

        except Exception as e:
            logger.error(f"Error during cleanup: {e}")

    def get_app_status(self) -> Dict[str, Any]:
        """Get comprehensive application status."""
        try:
            status = {
                "services": {
                    "audio_recorder": self.audio_recorder is not None,
                    "ai_service": self.ai_service is not None and self.ai_service.is_available(),
                    "document_processor": self.document_processor is not None
                },
                "state": {
                    "recording": self.recording,
                    "processing": self.processing,
                    "has_audio": bool(self.current_audio_file),
                    "has_transcript": bool(self.current_transcript),
                    "has_enhanced_text": bool(self.current_enhanced_text)
                },
                "config": {
                    "recording_method": self.config.get("recording_method"),
                    "language": self.config.get("language"),
                    "use_ai_enhancement": self.config.get("use_ai_enhancement"),
                    "output_folder": self.config.get("output_folder")
                }
            }
            return status
        except Exception as e:
            logger.error(f"Error getting app status: {e}")
            return {"error": str(e)}

    # =============================================================================
    # AUDIO VISUALIZATION AND MONITORING
    # =============================================================================

    def get_audio_queue(self):
        """Get audio queue for real-time visualization."""
        if self.audio_recorder:
            return self.audio_recorder.get_audio_queue()
        return None

    def set_visualization_mode(self, mode: str, sensitivity: float = 1.0):
        """Set audio visualization mode and sensitivity."""
        self.config.set("viz_mode", mode)
        self.config.set("viz_sensitivity", sensitivity)
        logger.info(f"Visualization mode set to: {mode} with sensitivity: {sensitivity}")

    def toggle_visualization(self, enabled: bool):
        """Enable or disable audio visualization."""
        self.config.set("viz_enabled", enabled)
        if self.audio_recorder:
            self.audio_recorder.set_visualization_enabled(enabled)
        logger.info(f"Audio visualization {'enabled' if enabled else 'disabled'}")

    # =============================================================================
    # CONFIGURATION MANAGEMENT
    # =============================================================================

    def update_api_key(self, api_key: str) -> bool:
        """Update the API key for AI services."""
        try:
            if self.config.save_user_api_key(api_key):
                if self.ai_service:
                    self.ai_service.update_api_key(api_key)
                self.status_callback("API key updated successfully")
                return True
            return False
        except Exception as e:
            self.error_handler.handle_error("config", e, "API key update")
            return False

    def get_config_value(self, key: str, default: Any = None) -> Any:
        """Get configuration value."""
        return self.config.get(key, default)

    def set_config_value(self, key: str, value: Any) -> None:
        """Set configuration value."""
        self.config.set(key, value)
        # Reinitialize services if audio config changed
        if key in ["sample_rate", "channels", "blocksize", "sample_width"]:
            self._reinitialize_audio_recorder()

    def _reinitialize_audio_recorder(self):
        """Reinitialize audio recorder with new config."""
        try:
            audio_config = self.config.get_audio_config()
            self.audio_recorder = AudioRecorder(audio_config, self.status_callback)
            logger.info("Audio recorder reinitialized with new config")
        except Exception as e:
            logger.error(f"Error reinitializing audio recorder: {e}")

    # =============================================================================
    # RECORDING OPERATIONS
    # =============================================================================

    def start_recording(self, mode: str = None) -> bool:
        """
        Start audio recording in specified mode.

        Args:
            mode: Recording mode ("microphone", "system", "dual")

        Returns:
            bool: True if recording started successfully
        """
        if self.recording:
            self.status_callback("Recording already in progress")
            return False

        if not self.audio_recorder:
            self.status_callback("Audio recorder not initialized")
            return False

        # Use configured mode if not specified
        if mode is None:
            mode = self.config.get("recording_mode", "microphone")

        try:
            # Start recording in a separate thread
            self.recording = True
            self.current_audio_file = None

            if self.audio_recorder.start_recording(mode):
                threading.Thread(
                    target=self._recording_thread,
                    args=(mode,),
                    daemon=True
                ).start()

                self.status_callback(f"Started {mode} recording")
                return True
            else:
                self.recording = False
                return False

        except Exception as e:
            self.recording = False
            self.error_handler.handle_error("audio", e, "starting recording")
            return False

    def stop_recording(self) -> bool:
        """Stop the current recording."""
        if not self.recording:
            self.status_callback("No recording in progress")
            return False

        try:
            self.recording = False
            if self.audio_recorder:
                self.audio_recorder.stop_recording()

            self.status_callback("Recording stopped. Processing...")
            return True

        except Exception as e:
            self.error_handler.handle_error("audio", e, "stopping recording")
            return False

    def _recording_thread(self, mode: str):
        """Thread function for handling recording operations."""
        try:
            # Execute recording based on mode
            if mode == "microphone":
                audio_file = self.audio_recorder.record_microphone_audio()
            elif mode == "system":
                audio_file = self.audio_recorder.record_system_audio()
            elif mode == "dual":
                audio_file = self.audio_recorder.record_dual_audio()
            else:
                logger.error(f"Unknown recording mode: {mode}")
                return

            if audio_file and os.path.exists(audio_file):
                self.current_audio_file = audio_file
                logger.info(f"Recording saved: {audio_file}")

                # Auto-process if configured
                if self.config.get("auto_process", True):
                    self._process_recording_thread()
            else:
                self.status_callback("Recording failed - no audio file created")

        except Exception as e:
            self.error_handler.handle_error("audio", e, "recording thread")
        finally:
            self.recording = False

    # =============================================================================
    # TRANSCRIPTION AND PROCESSING
    # =============================================================================

    def process_current_recording(self) -> bool:
        """Process the current audio recording."""
        if not self.current_audio_file:
            self.status_callback("No audio file to process")
            return False

        if self.processing:
            self.status_callback("Processing already in progress")
            return False

        # Start processing in separate thread
        threading.Thread(target=self._process_recording_thread, daemon=True).start()
        return True

    def start_processing(self) -> bool:
        """Start processing the current audio recording (alias for process_current_recording)."""
        return self.process_current_recording()

    def process_audio_file(self, audio_file_path: str) -> bool:
        """Process an external audio file."""
        if not os.path.exists(audio_file_path):
            self.status_callback("Audio file not found")
            return False

        self.current_audio_file = audio_file_path
        return self.process_current_recording()

    def _process_recording_thread(self):
        """Thread function for processing audio recordings with support for chunked and long recordings."""
        if self.processing:
            return

        self.processing = True

        try:
            # Check if AI service is available
            if not self.ai_service or not self.ai_service.is_available():
                self.status_callback("AI service not available. Please configure API key.")
                return

            # Handle chunked recordings from extended recording mode
            if self.config.get("use_extended_recording", False):
                self._process_chunked_recording()
            else:
                self._process_single_recording()

        except Exception as e:
            self.error_handler.handle_error("transcription", e, "processing recording")
        finally:
            self.processing = False

    def _process_single_recording(self):
        """Process a single audio recording."""
        self.status_callback("Starting transcription...")

        # Detect language if auto-detection is enabled
        language = self.config.get("language", "id")
        if self.config.get("auto_detect_language", False):
            self.status_callback("Detecting language...")
            quick_transcript = self.ai_service.transcribe_audio(
                self.current_audio_file, "auto", "whisper-large-v3"
            )
            if quick_transcript:
                detected_lang = self.ai_service.detect_language(quick_transcript[:500])
                language = detected_lang
                logger.info(f"Detected language: {language}")

        # Main transcription
        transcript = self.ai_service.transcribe_audio(
            self.current_audio_file,
            language,
            self._select_whisper_model()
        )

        if not transcript:
            self.status_callback("Transcription failed")
            return

        self.current_transcript = transcript
        self.status_callback("Transcription completed. Enhancing text...")

        # Enhance text if enabled
        if self.config.get("use_ai_enhancement", True):
            content_type = self.ai_service.analyze_content_type(transcript)
            enhanced_text = self.ai_service.enhance_text(
                transcript, content_type, "Indonesian" if language == "id" else "English"
            )

            if enhanced_text:
                self.current_enhanced_text = enhanced_text
                self.status_callback("Text enhancement completed")
            else:
                self.current_enhanced_text = transcript
                self.status_callback("Enhancement failed, using original transcript")
        else:
            self.current_enhanced_text = transcript

        # Create documents
        self._create_documents()

    def _process_chunked_recording(self):
        """Process chunked recordings from extended recording mode."""
        if not self.audio_recorder:
            return

        # Get all audio chunks from the recorder
        audio_chunks = self.audio_recorder.get_recorded_chunks()
        if not audio_chunks:
            self.status_callback("No audio chunks found")
            return

        self.status_callback(f"Processing {len(audio_chunks)} audio chunks...")

        all_transcripts = []
        language = self.config.get("language", "id")

        # Process each chunk
        for i, chunk_file in enumerate(audio_chunks):
            if not os.path.exists(chunk_file):
                continue

            self.status_callback(f"Transcribing chunk {i+1}/{len(audio_chunks)}...")

            transcript = self.ai_service.transcribe_audio(
                chunk_file,
                language,
                self._select_whisper_model()
            )

            if transcript:
                all_transcripts.append(transcript)

            # Add delay between API calls
            time.sleep(self.config.get("api_request_delay", 1.0))

        if not all_transcripts:
            self.status_callback("No successful transcriptions")
            return

        # Combine all transcripts
        combined_transcript = " ".join(all_transcripts)
        self.current_transcript = combined_transcript

        # Enhance combined text if enabled
        if self.config.get("use_ai_enhancement", True):
            self.status_callback("Enhancing combined text...")
            content_type = self.ai_service.analyze_content_type(combined_transcript)

            # For long texts, use document cohesion enhancement
            enhanced_text = self.ai_service.enhance_document_cohesion(
                combined_transcript, content_type, "Indonesian" if language == "id" else "English"
            )

            if enhanced_text:
                self.current_enhanced_text = enhanced_text
                self.status_callback("Text enhancement completed")
            else:
                self.current_enhanced_text = combined_transcript
                self.status_callback("Enhancement failed, using original transcript")
        else:
            self.current_enhanced_text = combined_transcript

        # Create documents
        self._create_documents()

    def _select_whisper_model(self) -> str:
        """Select the appropriate Whisper model based on configuration."""
        if self.config.get("use_economic_model", False):
            language = self.config.get("language", "id")
            if language.startswith("en"):
                return "distil-whisper-large-v3-en"
            else:
                return "whisper-large-v3"
        return "whisper-large-v3"

    def _create_documents(self):
        """Create output documents from processed text."""
        if not self.current_enhanced_text or not self.document_processor:
            return

        try:
            self.status_callback("Creating documents...")

            # Prepare metadata
            metadata = {
                "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                "audio_file": Path(self.current_audio_file).name if self.current_audio_file else "Unknown",
                "language": self.config.get("language", "id"),
                "content_type": self.ai_service.analyze_content_type(self.current_enhanced_text) if self.ai_service else "general"
            }

            # Create documents in configured formats
            output_formats = self.config.get("output_formats", ["markdown"])
            created_files = []

            for format_type in output_formats:
                if format_type == "markdown":
                    file_path = self.document_processor.create_markdown_document(
                        self.current_enhanced_text, "Transkrip Audio", metadata
                    )
                elif format_type == "word":
                    file_path = self.document_processor.create_word_document(
                        self.current_enhanced_text, "Transkrip Audio", metadata
                    )
                elif format_type == "text":
                    file_path = self.document_processor.create_text_document(
                        self.current_enhanced_text, "Transkrip Audio", metadata
                    )
                else:
                    continue

                if file_path:
                    created_files.append(file_path)

            # Create summary document if requested
            if self.config.get("create_summary", False) and self.current_transcript:
                summary_file = self.document_processor.create_summary_document(
                    self.current_transcript, self.current_enhanced_text, metadata
                )
                if summary_file:
                    created_files.append(summary_file)

            if created_files:
                self.status_callback(f"Documents created: {len(created_files)} files")
                logger.info(f"Created documents: {created_files}")
            else:
                self.status_callback("No documents created")

        except Exception as e:
            self.error_handler.handle_error("document", e, "creating documents")

    # =============================================================================
    # PUBLIC INTERFACE
    # =============================================================================

    def get_status(self) -> Dict[str, Any]:
        """Get current application status."""
        return {
            "recording": self.recording,
            "processing": self.processing,
            "has_audio_file": self.current_audio_file is not None,
            "has_transcript": self.current_transcript is not None,
            "has_enhanced_text": self.current_enhanced_text is not None,
            "ai_available": self.ai_service.is_available() if self.ai_service else False,
            "api_key_configured": self.config.has_valid_api_key()
        }

    def get_current_results(self) -> Dict[str, Optional[str]]:
        """Get current processing results."""
        return {
            "audio_file": self.current_audio_file,
            "transcript": self.current_transcript,
            "enhanced_text": self.current_enhanced_text
        }

    def clear_current_session(self):
        """Clear current session data."""
        self.current_audio_file = None
        self.current_transcript = None
        self.current_enhanced_text = None
        self.status_callback("Session cleared")

    def get_available_microphones(self) -> list:
        """Get list of available microphones with detailed information."""
        try:
            if self.audio_recorder:
                return self.audio_recorder.get_available_microphones()

            # Fallback to soundcard
            import soundcard as sc
            mics = sc.all_microphones()
            mic_list = []
            for i, mic in enumerate(mics):
                mic_info = {
                    'id': i,
                    'name': mic.name,
                    'channels': getattr(mic, 'channels', 'Unknown'),
                    'default': mic == sc.default_microphone()
                }
                mic_list.append(mic_info)
            return mic_list

        except Exception as e:
            logger.error(f"Error getting microphones: {e}")
            return []

    def get_available_speakers(self) -> list:
        """Get list of available speakers with loopback capability."""
        try:
            if self.audio_recorder:
                return self.audio_recorder.get_available_speakers()

            # Fallback to soundcard
            import soundcard as sc
            speakers = sc.all_speakers()
            speaker_list = []
            for i, speaker in enumerate(speakers):
                speaker_info = {
                    'id': i,
                    'name': speaker.name,
                    'channels': getattr(speaker, 'channels', 'Unknown'),
                    'default': speaker == sc.default_speaker(),
                    'loopback_available': True  # Most speakers support loopback
                }
                speaker_list.append(speaker_info)
            return speaker_list

        except Exception as e:
            logger.error(f"Error getting speakers: {e}")
            return []

    def set_audio_device(self, device_type: str, device_id: int) -> bool:
        """Set the audio device for recording."""
        try:
            if device_type == "microphone":
                self.config.set("selected_mic", f"{device_id}")
                if self.audio_recorder:
                    self.audio_recorder.set_microphone(device_id)
            elif device_type == "speaker":
                self.config.set("selected_speaker", f"{device_id}")
                if self.audio_recorder:
                    self.audio_recorder.set_speaker(device_id)

            self.status_callback(f"Audio device set: {device_type} {device_id}")
            return True

        except Exception as e:
            logger.error(f"Error setting audio device: {e}")
            return False

    def test_audio_device(self, device_type: str, device_id: int) -> bool:
        """Test if an audio device is working properly."""
        try:
            if self.audio_recorder:
                return self.audio_recorder.test_device(device_type, device_id)
            return False
        except Exception as e:
            logger.error(f"Error testing audio device: {e}")
            return False

    def cleanup(self):
        """Cleanup resources and temporary files."""
        try:
            if self.audio_recorder:
                self.audio_recorder.cleanup()
            logger.info("Application cleanup completed")
        except Exception as e:
            logger.error(f"Error during cleanup: {e}")

    def __del__(self):
        """Destructor to ensure cleanup."""
        self.cleanup()
