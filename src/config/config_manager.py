# -*- coding: utf-8 -*-
"""
Enhanced Configuration Management Module for EchoScribe AI
Fully integrated from monolithic system with complete API key management.
Handles all application settings securely and provides migration support.
"""

import json
import os
import logging
from pathlib import Path
from typing import Dict, Any, Optional, Callable

logger = logging.getLogger(__name__)

class ConfigManager:
    """
    Enhanced configuration management integrated from monolithic system.
    Provides secure API key handling and complete settings management.
    """

    def __init__(self):
        # Standard configuration paths
        self.app_data_dir = Path.home() / ".echoscribe"
        self.config_file = self.app_data_dir / "config.json"
        self.api_key_file = self.app_data_dir / "api_key.json"

        # Ensure directory exists
        self.app_data_dir.mkdir(exist_ok=True)

        # Load configuration
        self._config_data = self._load_config()

        # Migration from legacy config
        self._migrate_legacy_config()

    def _load_config(self) -> Dict[str, Any]:
        """Load configuration with defaults from monolithic system."""
        default_config = {
            # Recording settings from monolithic
            "use_system_audio": False,
            "use_dual_recording": False,
            "recording_mode": "microphone",
            "selected_mic": "0: Default Microphone",
            "use_extended_recording": True,
            "chunk_size": 600,

            # Audio settings from monolithic
            "sample_rate": 48000,
            "channels": 2,
            "blocksize": 1024,
            "sample_width": 2,
            "default_samplerate": 48000,
            "default_channels": 2,
            "default_blocksize": 1024,

            # Output settings from monolithic
            "output_folder": str(Path.home() / "Documents"),
            "filename_prefix": "catatan",

            # AI settings from monolithic
            "language": "id-ID",
            "engine": "Google",
            "use_ai_enhancement": True,
            "use_economic_model": False,
            "max_tokens": 4000,
            "api_request_delay": 10,

            # UI settings from monolithic
            "theme": "dark",
            "viz_enabled": True,
            "viz_mode": "waveform",
            "viz_sensitivity": 1.0,

            # Processing settings from monolithic
            "heading_spacing_before": 12,
            "heading_spacing_after": 6,
            "paragraph_spacing": 6,

            # Performance settings
            "rate_limit_delay": 1.0,
            "max_concurrent_requests": 3,

            # Error handling settings
            "retry_attempts": 3,
            "retry_delay": 2.0
        }

        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    # Merge with defaults to handle new settings
                    default_config.update(loaded_config)
                    logger.info("Configuration loaded successfully")
            else:
                logger.info("Creating default configuration")
                # Save default config
                self._save_config_data(default_config)
        except Exception as e:
            logger.error(f"Error loading config: {e}. Using defaults.")

        return default_config

    def _save_config_data(self, config_data: Dict[str, Any]) -> bool:
        """Save configuration data to file."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            logger.error(f"Error saving config: {e}")
            return False

    def _migrate_legacy_config(self):
        """Migrate configuration from legacy monolithic format."""
        try:
            # Check for legacy config patterns and migrate
            legacy_keys = {
                "audio_device": "selected_mic",
                "recording_duration": "chunk_size",
                "ai_model": "use_economic_model"
            }

            migrated = False
            for old_key, new_key in legacy_keys.items():
                if old_key in self._config_data and new_key not in self._config_data:
                    self._config_data[new_key] = self._config_data.pop(old_key)
                    migrated = True

            if migrated:
                self.save_config()
                logger.info("Legacy configuration migrated successfully")

        except Exception as e:
            logger.error(f"Error during config migration: {e}")

    def save_config(self) -> bool:
        """Save current configuration to file."""
        try:
            return self._save_config_data(self._config_data)
        except Exception as e:
            logger.error(f"Error saving config: {e}")
            return False

    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value."""
        return self._config_data.get(key, default)

    def set(self, key: str, value: Any) -> None:
        """Set configuration value and auto-save."""
        self._config_data[key] = value
        self.save_config()

    def update(self, config_dict: Dict[str, Any]) -> None:
        """Update multiple configuration values."""
        self._config_data.update(config_dict)
        self.save_config()

    def get_all_config(self) -> Dict[str, Any]:
        """Get all configuration data."""
        return self._config_data.copy()

    # API Key Management (Enhanced from monolithic)
    def save_user_api_key(self, api_key: str) -> bool:
        """
        Save user-provided API key securely.
        Enhanced from monolithic system with validation.
        """
        try:
            if not api_key or not isinstance(api_key, str):
                logger.warning("Invalid API key: empty or not string")
                return False

            if not api_key.startswith("gsk_"):
                logger.warning("Invalid Groq API key format - must start with 'gsk_'")
                return False

            if len(api_key) < 20:
                logger.warning("API key appears too short")
                return False

            # Save API key in separate file for security
            api_key_data = {
                "groq_api_key": api_key,
                "key_type": "user_provided",
                "saved_at": str(Path.home() / ".echoscribe")
            }

            with open(self.api_key_file, 'w', encoding='utf-8') as f:
                json.dump(api_key_data, f, indent=2)

            # Also save in main config for backward compatibility
            self.set("groq_api_key", api_key)

            logger.info("User API key saved successfully")
            return True

        except Exception as e:
            logger.error(f"Error saving API key: {e}")
            return False

    def get_user_api_key(self) -> Optional[str]:
        """
        Get user-provided API key.
        Enhanced from monolithic system with multiple sources.
        """
        try:
            # Try dedicated API key file first
            if self.api_key_file.exists():
                with open(self.api_key_file, 'r', encoding='utf-8') as f:
                    api_data = json.load(f)
                    api_key = api_data.get("groq_api_key")
                    if api_key and api_key.startswith("gsk_"):
                        return api_key

            # Fallback to main config
            api_key = self.get("groq_api_key")
            if api_key and api_key.startswith("gsk_"):
                return api_key

            return None

        except Exception as e:
            logger.error(f"Error loading API key: {e}")
            return None

    def remove_user_api_key(self) -> bool:
        """Remove user-provided API key."""
        try:
            # Remove from dedicated file
            if self.api_key_file.exists():
                self.api_key_file.unlink()

            # Remove from main config
            if "groq_api_key" in self._config_data:
                del self._config_data["groq_api_key"]
                self.save_config()

            logger.info("User API key removed successfully")
            return True

        except Exception as e:
            logger.error(f"Error removing API key: {e}")
            return False

    def has_user_api_key(self) -> bool:
        """Check if user has provided an API key."""
        return self.get_user_api_key() is not None

    def validate_api_key(self, api_key: str) -> bool:
        """Validate API key format."""
        if not api_key or not isinstance(api_key, str):
            return False

        if not api_key.startswith("gsk_"):
            return False

        if len(api_key) < 20:
            return False

        return True

    def has_valid_api_key(self) -> bool:
        """Check if a valid API key is configured."""
        api_key = self.get_user_api_key()
        return api_key is not None and self.validate_api_key(api_key)

    # Enhanced Configuration Getters (From monolithic system)
    def get_audio_config(self) -> Dict[str, Any]:
        """Get complete audio configuration for AudioRecorder."""
        return {
            "sample_rate": self.get("sample_rate", 48000),
            "channels": self.get("channels", 2),
            "blocksize": self.get("blocksize", 1024),
            "sample_width": self.get("sample_width", 2),
            "selected_mic": self.get("selected_mic", "0"),
            "selected_speaker": self.get("selected_speaker", "0"),
            "use_system_audio": self.get("use_system_audio", False),
            "use_dual_recording": self.get("use_dual_recording", False),
            "recording_mode": self.get("recording_mode", "microphone"),
            "use_extended_recording": self.get("use_extended_recording", True),
            "chunk_size": self.get("chunk_size", 600),
            "viz_enabled": self.get("viz_enabled", True),
            "viz_mode": self.get("viz_mode", "waveform"),
            "viz_sensitivity": self.get("viz_sensitivity", 1.0)
        }

    def get_doc_config(self) -> Dict[str, Any]:
        """Get complete document processing configuration."""
        return {
            "output_folder": self.get("output_folder", str(Path.home() / "Documents")),
            "filename_prefix": self.get("filename_prefix", "catatan"),
            "output_formats": self.get("output_formats", ["markdown"]),
            "heading_spacing_before": self.get("heading_spacing_before", 12),
            "heading_spacing_after": self.get("heading_spacing_after", 6),
            "paragraph_spacing": self.get("paragraph_spacing", 6),
            "create_summary": self.get("create_summary", False),
            "include_metadata": self.get("include_metadata", True),
            "document_theme": self.get("document_theme", "professional")
        }

    # Recording Configuration (from monolithic)
    def get_recording_config(self) -> Dict[str, Any]:
        """Get recording-specific configuration."""
        return {
            "recording_mode": self.get("recording_mode", "microphone"),
            "selected_mic": self.get("selected_mic", "0: Default Microphone"),
            "use_system_audio": self.get("use_system_audio", False),
            "use_dual_recording": self.get("use_dual_recording", False),
            "use_extended_recording": self.get("use_extended_recording", True),
            "chunk_size": self.get("chunk_size", 600),
            "sample_rate": self.get("sample_rate", 48000),
            "channels": self.get("channels", 2),
            "blocksize": self.get("blocksize", 1024),
            "sample_width": self.get("sample_width", 2)
        }

    def get_ai_config(self) -> Dict[str, Any]:
        """Get AI-specific configuration."""
        return {
            "language": self.get("language", "id-ID"),
            "engine": self.get("engine", "Google"),
            "use_ai_enhancement": self.get("use_ai_enhancement", True),
            "use_economic_model": self.get("use_economic_model", False),
            "max_tokens": self.get("max_tokens", 4000),
            "api_request_delay": self.get("api_request_delay", 10),
            "groq_api_key": self.get_user_api_key()
        }

    def get_ui_config(self) -> Dict[str, Any]:
        """Get UI-specific configuration."""
        return {
            "theme": self.get("theme", "dark"),
            "viz_enabled": self.get("viz_enabled", True),
            "viz_mode": self.get("viz_mode", "waveform"),
            "viz_sensitivity": self.get("viz_sensitivity", 1.0)
        }

    def get_output_config(self) -> Dict[str, Any]:
        """Get output-specific configuration."""
        return {
            "output_folder": self.get("output_folder", str(Path.home() / "Documents")),
            "filename_prefix": self.get("filename_prefix", "catatan"),
            "heading_spacing_before": self.get("heading_spacing_before", 12),
            "heading_spacing_after": self.get("heading_spacing_after", 6),
            "paragraph_spacing": self.get("paragraph_spacing", 6)
        }

    # Theme and UI Management
    def get_theme_colors(self) -> Dict[str, str]:
        """Get theme colors based on current theme setting."""
        if self.get("theme") == "dark":
            return {
                "bg_color": "#1E1E1E",
                "fg_color": "#E0E0E0",
                "accent_color": "#007ACC",
                "button_color": "#2A2A2A",
                "button_hover": "#3A3A3A",
                "border_color": "#3E3E3E"
            }
        else:  # light theme
            return {
                "bg_color": "#F0F0F0",
                "fg_color": "#1E1E1E",
                "accent_color": "#007ACC",
                "button_color": "#E0E0E0",
                "button_hover": "#D0D0D0",
                "border_color": "#C0C0C0"
            }

    # Import/Export Configuration
    def export_config(self, file_path: str) -> bool:
        """Export configuration to file."""
        try:
            export_data = {
                "config": self._config_data,
                "export_version": "1.0",
                "app_version": "EchoScribe AI v2.0"
            }

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, indent=2, ensure_ascii=False)

            logger.info(f"Configuration exported to {file_path}")
            return True

        except Exception as e:
            logger.error(f"Error exporting config: {e}")
            return False

    def import_config(self, file_path: str) -> bool:
        """Import configuration from file."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                import_data = json.load(f)

            if "config" in import_data:
                # Validate and merge imported config
                imported_config = import_data["config"]
                # Don't import API keys for security
                if "groq_api_key" in imported_config:
                    del imported_config["groq_api_key"]

                self._config_data.update(imported_config)
                self.save_config()

                logger.info(f"Configuration imported from {file_path}")
                return True
            else:
                logger.error("Invalid config file format")
                return False

        except Exception as e:
            logger.error(f"Error importing config: {e}")
            return False

    # Configuration Validation
    def validate_config(self) -> Dict[str, Any]:
        """Validate current configuration and return status."""
        validation_result = {
            "valid": True,
            "warnings": [],
            "errors": []
        }

        try:
            # Validate output folder
            output_folder = self.get("output_folder")
            if not os.path.exists(output_folder):
                validation_result["warnings"].append(f"Output folder does not exist: {output_folder}")

            # Validate API key
            api_key = self.get_user_api_key()
            if api_key and not self.validate_api_key(api_key):
                validation_result["errors"].append("Invalid API key format")
                validation_result["valid"] = False

            # Validate numeric settings
            numeric_settings = ["chunk_size", "sample_rate", "channels", "blocksize", "max_tokens"]
            for setting in numeric_settings:
                value = self.get(setting)
                if not isinstance(value, (int, float)) or value <= 0:
                    validation_result["errors"].append(f"Invalid {setting}: {value}")
                    validation_result["valid"] = False

        except Exception as e:
            logger.error(f"Error during config validation: {e}")
            validation_result["errors"].append(f"Validation error: {e}")
            validation_result["valid"] = False

        return validation_result

    def reset_to_defaults(self) -> bool:
        """Reset configuration to default values."""
        try:
            # Backup current API key
            api_key = self.get_user_api_key()

            # Load defaults
            self._config_data = self._load_config()

            # Restore API key if it existed
            if api_key:
                self.save_user_api_key(api_key)

            self.save_config()
            logger.info("Configuration reset to defaults")
            return True

        except Exception as e:
            logger.error(f"Error resetting config: {e}")
            return False
        """Get user-provided API key."""
        return self.get("groq_api_key")

    def remove_user_api_key(self) -> bool:
        """Remove user API key from configuration."""
        try:
            if "groq_api_key" in self._config_data:
                del self._config_data["groq_api_key"]
                self.save_config()
            logger.info("User API key removed")
            return True
        except Exception as e:
            logger.error(f"Error removing API key: {e}")
            return False

    def has_valid_api_key(self) -> bool:
        """Check if a valid API key is configured."""
        api_key = self.get_user_api_key()
        return api_key is not None and api_key.startswith("gsk_")

    # Audio Configuration Helpers
    def get_audio_config(self) -> Dict[str, Any]:
        """Get all audio-related configuration."""
        return {
            "sample_rate": self.get("sample_rate", 48000),
            "channels": self.get("channels", 2),
            "blocksize": self.get("blocksize", 1024),
            "sample_width": self.get("sample_width", 2),
            "chunk_size": self.get("chunk_size", 600),
            "use_extended_recording": self.get("use_extended_recording", True)
        }

    def get_recording_config(self) -> Dict[str, Any]:
        """Get recording mode configuration."""
        return {
            "recording_mode": self.get("recording_mode", "microphone"),
            "use_system_audio": self.get("use_system_audio", False),
            "use_dual_recording": self.get("use_dual_recording", False)
        }

    def get_ai_config(self) -> Dict[str, Any]:
        """Get AI-related configuration."""
        return {
            "language": self.get("language", "id-ID"),
            "engine": self.get("engine", "Google"),
            "use_ai_enhancement": self.get("use_ai_enhancement", True),
            "use_economic_model": self.get("use_economic_model", False)
        }
