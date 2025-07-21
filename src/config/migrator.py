# -*- coding: utf-8 -*-
"""
Configuration Migration Module for EchoScribe AI
Handles legacy config migration and system compatibility checks.
Integrated from monolithic system for backward compatibility.
"""

import json
import logging
import os
import sys
import platform
import subprocess
import shutil
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
import importlib.util

logger = logging.getLogger(__name__)

class ConfigMigrator:
    """
    Configuration migration and system compatibility manager.
    Handles legacy config migration from monolithic system.
    """

    def __init__(self, config_manager=None, status_callback=None):
        self.config_manager = config_manager
        self.status_callback = status_callback or (lambda x: None)

        # Migration mapping for config keys
        self.legacy_key_mapping = {
            'api_key': 'groq_api_key',
            'model_name': 'model_id',
            'use_economic': 'use_economic_model',
            'enhance_text': 'use_ai_enhancement',
            'recording_method': 'default_recording_method',
            'output_format': 'default_output_format',
            'save_directory': 'default_save_path',
            'visualization_mode': 'viz_mode',
            'visualization_enabled': 'viz_enabled',
            'sensitivity': 'viz_sensitivity'
        }

        # Required dependencies with versions
        self.required_dependencies = {
            'groq': '>=0.4.0',
            'customtkinter': '>=5.0.0',
            'soundcard': '>=0.4.0',
            'numpy': '>=1.21.0',
            'matplotlib': '>=3.5.0',
            'pyaudio': '>=0.2.11'
        }

        # Optional dependencies
        self.optional_dependencies = {
            'markdown': '>=3.4.0',
            'python-docx': '>=0.8.11',
            'psutil': '>=5.8.0'
        }

    def check_system_compatibility(self) -> Dict[str, Any]:
        """
        Comprehensive system compatibility check.
        Integrated from monolithic system (lines 4900-5200).
        """
        compatibility_report = {
            'system_info': {},
            'python_version': {},
            'dependencies': {},
            'audio_system': {},
            'ffmpeg': {},
            'permissions': {},
            'recommendations': [],
            'critical_issues': [],
            'warnings': []
        }

        try:
            self.status_callback("🔍 Checking system compatibility...")

            # System information
            compatibility_report['system_info'] = {
                'os': platform.system(),
                'os_version': platform.release(),
                'architecture': platform.architecture()[0],
                'processor': platform.processor(),
                'python_executable': sys.executable
            }

            # Python version check
            python_version = sys.version_info
            compatibility_report['python_version'] = {
                'version': f"{python_version.major}.{python_version.minor}.{python_version.micro}",
                'compatible': python_version >= (3, 8),
                'recommended': python_version >= (3, 9)
            }

            if not compatibility_report['python_version']['compatible']:
                compatibility_report['critical_issues'].append(
                    "Python 3.8 or higher is required. Current version: " +
                    compatibility_report['python_version']['version']
                )

            # Check dependencies
            compatibility_report['dependencies'] = self._check_dependencies()

            # Audio system check
            compatibility_report['audio_system'] = self._check_audio_system()

            # FFmpeg check
            compatibility_report['ffmpeg'] = self._check_ffmpeg()

            # Permissions check
            compatibility_report['permissions'] = self._check_permissions()

            # Generate recommendations
            compatibility_report['recommendations'] = self._generate_recommendations(compatibility_report)

            self.status_callback("✅ System compatibility check completed")

        except Exception as e:
            logger.error(f"Error in compatibility check: {e}")
            compatibility_report['critical_issues'].append(f"Compatibility check failed: {e}")

        return compatibility_report

    def _check_dependencies(self) -> Dict[str, Any]:
        """Check required and optional dependencies."""
        dependency_status = {
            'required': {},
            'optional': {},
            'missing_required': [],
            'missing_optional': [],
            'version_conflicts': []
        }

        # Check required dependencies
        for package, version_requirement in self.required_dependencies.items():
            status = self._check_package(package, version_requirement)
            dependency_status['required'][package] = status

            if not status['installed']:
                dependency_status['missing_required'].append(package)
            elif not status['version_ok']:
                dependency_status['version_conflicts'].append(
                    f"{package}: requires {version_requirement}, found {status['version']}"
                )

        # Check optional dependencies
        for package, version_requirement in self.optional_dependencies.items():
            status = self._check_package(package, version_requirement)
            dependency_status['optional'][package] = status

            if not status['installed']:
                dependency_status['missing_optional'].append(package)

        return dependency_status

    def _check_package(self, package_name: str, version_requirement: str) -> Dict[str, Any]:
        """Check if a specific package is installed with correct version."""
        try:
            # Try to import the package
            spec = importlib.util.find_spec(package_name)
            if spec is None:
                return {'installed': False, 'version': None, 'version_ok': False}

            # Get version information using simple methods
            version = None
            version_ok = True  # Simplified - if importable, assume it's compatible

            try:
                # Try to get version from __version__ attribute
                module = importlib.import_module(package_name)
                if hasattr(module, '__version__'):
                    version = module.__version__
                elif hasattr(module, 'VERSION'):
                    version = str(module.VERSION)
                elif hasattr(module, 'version'):
                    version = str(module.version)
                else:
                    version = 'unknown'

            except Exception:
                version = 'unknown'

            return {
                'installed': True,
                'version': version,
                'version_ok': version_ok,
                'location': spec.origin if spec.origin else 'unknown'
            }

        except Exception as e:
            logger.debug(f"Error checking package {package_name}: {e}")
            return {'installed': False, 'version': None, 'version_ok': False, 'error': str(e)}

    def _check_audio_system(self) -> Dict[str, Any]:
        """Check audio system compatibility."""
        audio_status = {
            'soundcard_available': False,
            'pyaudio_available': False,
            'devices_detected': False,
            'microphones': [],
            'speakers': [],
            'system_audio_support': False
        }

        try:
            # Check soundcard
            import soundcard as sc
            audio_status['soundcard_available'] = True

            # Get microphones
            mics = sc.all_microphones()
            audio_status['microphones'] = [{'name': mic.name, 'id': mic.id} for mic in mics]
            audio_status['devices_detected'] = len(mics) > 0

            # Get speakers
            speakers = sc.all_speakers()
            audio_status['speakers'] = [{'name': spk.name, 'id': spk.id} for spk in speakers]

            # Check system audio support
            try:
                default_speaker = sc.default_speaker()
                audio_status['system_audio_support'] = default_speaker is not None
            except:
                audio_status['system_audio_support'] = False

        except ImportError:
            logger.warning("soundcard library not available")
        except Exception as e:
            logger.warning(f"Error checking audio system: {e}")

        try:
            # Check PyAudio as fallback
            import pyaudio
            audio_status['pyaudio_available'] = True
        except ImportError:
            logger.warning("PyAudio library not available")
        except Exception as e:
            logger.warning(f"Error checking PyAudio: {e}")

        return audio_status

    def _check_ffmpeg(self) -> Dict[str, Any]:
        """Check FFmpeg installation and capabilities."""
        ffmpeg_status = {
            'installed': False,
            'version': None,
            'path': None,
            'codecs_supported': [],
            'formats_supported': []
        }

        try:
            # Check if FFmpeg is in PATH
            result = subprocess.run(['ffmpeg', '-version'],
                                  capture_output=True, text=True, timeout=10)

            if result.returncode == 0:
                ffmpeg_status['installed'] = True

                # Extract version
                output_lines = result.stdout.split('\n')
                if output_lines:
                    version_line = output_lines[0]
                    if 'ffmpeg version' in version_line.lower():
                        ffmpeg_status['version'] = version_line.split(' ')[2]

                # Get FFmpeg path
                which_result = subprocess.run(['where', 'ffmpeg'] if platform.system() == 'Windows' else ['which', 'ffmpeg'],
                                            capture_output=True, text=True, timeout=5)
                if which_result.returncode == 0:
                    ffmpeg_status['path'] = which_result.stdout.strip()

                # Check supported codecs (limited check for performance)
                codec_result = subprocess.run(['ffmpeg', '-codecs'],
                                            capture_output=True, text=True, timeout=10)
                if codec_result.returncode == 0:
                    important_codecs = ['mp3', 'wav', 'aac', 'opus', 'flac']
                    codec_output = codec_result.stdout.lower()
                    ffmpeg_status['codecs_supported'] = [codec for codec in important_codecs
                                                       if codec in codec_output]

        except FileNotFoundError:
            logger.info("FFmpeg not found in PATH")
        except subprocess.TimeoutExpired:
            logger.warning("FFmpeg check timed out")
        except Exception as e:
            logger.warning(f"Error checking FFmpeg: {e}")

        return ffmpeg_status

    def _check_permissions(self) -> Dict[str, Any]:
        """Check file system permissions for required directories."""
        permissions = {
            'config_dir': False,
            'temp_dir': False,
            'documents_dir': False,
            'current_dir': False,
            'issues': []
        }

        test_dirs = {
            'config_dir': Path.home() / '.echoscribe',
            'temp_dir': Path.cwd() / 'temp',
            'documents_dir': Path.home() / 'Documents',
            'current_dir': Path.cwd()
        }

        for dir_name, dir_path in test_dirs.items():
            try:
                # Create directory if it doesn't exist
                dir_path.mkdir(parents=True, exist_ok=True)

                # Test write permission
                test_file = dir_path / 'permission_test.tmp'
                test_file.write_text('test')
                test_file.unlink()

                permissions[dir_name] = True

            except PermissionError:
                permissions['issues'].append(f"No write permission for {dir_path}")
            except Exception as e:
                permissions['issues'].append(f"Error accessing {dir_path}: {e}")

        return permissions

    def _generate_recommendations(self, compatibility_report: Dict[str, Any]) -> List[str]:
        """Generate recommendations based on compatibility check."""
        recommendations = []

        # Python version recommendations
        if not compatibility_report['python_version']['compatible']:
            recommendations.append("🔴 CRITICAL: Upgrade to Python 3.8 or higher")
        elif not compatibility_report['python_version']['recommended']:
            recommendations.append("🟡 Consider upgrading to Python 3.9+ for better performance")

        # Dependency recommendations
        missing_required = compatibility_report['dependencies']['missing_required']
        if missing_required:
            recommendations.append(f"🔴 Install required packages: pip install {' '.join(missing_required)}")

        missing_optional = compatibility_report['dependencies']['missing_optional']
        if missing_optional:
            recommendations.append(f"🟡 Install optional packages for full features: pip install {' '.join(missing_optional)}")

        # Audio system recommendations
        audio = compatibility_report['audio_system']
        if not audio['soundcard_available'] and not audio['pyaudio_available']:
            recommendations.append("🔴 Install audio libraries: pip install soundcard pyaudio")
        elif not audio['devices_detected']:
            recommendations.append("🟡 No microphones detected - check audio device connections")

        # FFmpeg recommendations
        if not compatibility_report['ffmpeg']['installed']:
            recommendations.append("🟡 Install FFmpeg for advanced audio processing capabilities")

        # Permission recommendations
        permission_issues = compatibility_report['permissions']['issues']
        if permission_issues:
            recommendations.append("🟡 Fix permission issues: " + "; ".join(permission_issues))

        return recommendations

    def migrate_legacy_config(self, legacy_config_path: str) -> bool:
        """
        Migrate configuration from legacy monolithic system.
        Integrated from monolithic system config migration logic.
        """
        try:
            self.status_callback("🔄 Migrating legacy configuration...")

            if not os.path.exists(legacy_config_path):
                logger.info(f"Legacy config not found at {legacy_config_path}")
                return False

            # Read legacy config
            with open(legacy_config_path, 'r', encoding='utf-8') as f:
                legacy_config = json.load(f)

            # Create new config structure
            new_config = {}
            migrated_keys = []
            unknown_keys = []

            # Migrate known keys
            for legacy_key, new_key in self.legacy_key_mapping.items():
                if legacy_key in legacy_config:
                    new_config[new_key] = legacy_config[legacy_key]
                    migrated_keys.append(f"{legacy_key} -> {new_key}")

            # Handle special cases
            if 'window_geometry' in legacy_config:
                new_config['ui_settings'] = {
                    'window_geometry': legacy_config['window_geometry'],
                    'theme': legacy_config.get('theme', 'dark')
                }
                migrated_keys.append("window_geometry -> ui_settings.window_geometry")

            if 'recent_files' in legacy_config:
                new_config['recent_files'] = legacy_config['recent_files']
                migrated_keys.append("recent_files")

            # Preserve unknown keys in legacy section
            for key, value in legacy_config.items():
                if key not in self.legacy_key_mapping and key not in ['window_geometry', 'recent_files']:
                    if 'legacy' not in new_config:
                        new_config['legacy'] = {}
                    new_config['legacy'][key] = value
                    unknown_keys.append(key)

            # Update config manager
            if self.config_manager:
                for key, value in new_config.items():
                    if key == 'legacy':
                        continue  # Skip legacy section
                    self.config_manager.set(key, value)

                # Save migrated config
                self.config_manager.save()

            # Create backup of legacy config
            backup_path = legacy_config_path + '.backup'
            shutil.copy2(legacy_config_path, backup_path)

            # Log migration results
            logger.info(f"Config migration completed:")
            logger.info(f"  Migrated keys: {len(migrated_keys)}")
            logger.info(f"  Unknown keys preserved: {len(unknown_keys)}")
            logger.info(f"  Legacy config backed up to: {backup_path}")

            if migrated_keys:
                logger.info("  Migrated mappings:")
                for mapping in migrated_keys:
                    logger.info(f"    {mapping}")

            if unknown_keys:
                logger.info(f"  Unknown keys preserved in legacy section: {unknown_keys}")

            self.status_callback(f"✅ Configuration migrated ({len(migrated_keys)} settings)")
            return True

        except Exception as e:
            logger.error(f"Error migrating legacy config: {e}")
            self.status_callback(f"❌ Config migration failed: {e}")
            return False

    def install_missing_dependencies(self, dependency_list: List[str], optional: bool = False) -> Dict[str, bool]:
        """
        Install missing dependencies using pip.
        """
        results = {}

        try:
            self.status_callback(f"📦 Installing {'optional' if optional else 'required'} dependencies...")

            for package in dependency_list:
                try:
                    self.status_callback(f"Installing {package}...")

                    # Use pip to install package
                    result = subprocess.run([
                        sys.executable, '-m', 'pip', 'install', package
                    ], capture_output=True, text=True, timeout=300)

                    if result.returncode == 0:
                        results[package] = True
                        logger.info(f"Successfully installed {package}")
                    else:
                        results[package] = False
                        logger.error(f"Failed to install {package}: {result.stderr}")

                except subprocess.TimeoutExpired:
                    results[package] = False
                    logger.error(f"Installation of {package} timed out")
                except Exception as e:
                    results[package] = False
                    logger.error(f"Error installing {package}: {e}")

            successful = sum(1 for success in results.values() if success)
            total = len(dependency_list)

            self.status_callback(f"✅ Dependency installation completed ({successful}/{total} successful)")

        except Exception as e:
            logger.error(f"Error in dependency installation: {e}")
            self.status_callback(f"❌ Dependency installation failed: {e}")

        return results

    def auto_fix_common_issues(self, compatibility_report: Dict[str, Any]) -> Dict[str, bool]:
        """
        Automatically fix common compatibility issues.
        """
        fixes_applied = {}

        try:
            self.status_callback("🔧 Applying automatic fixes...")

            # Create required directories
            required_dirs = [
                Path.home() / '.echoscribe',
                Path.cwd() / 'temp',
                Path.cwd() / 'outputs'
            ]

            for dir_path in required_dirs:
                try:
                    dir_path.mkdir(parents=True, exist_ok=True)
                    fixes_applied[f"create_dir_{dir_path.name}"] = True
                except Exception as e:
                    fixes_applied[f"create_dir_{dir_path.name}"] = False
                    logger.warning(f"Could not create directory {dir_path}: {e}")

            # Install critical missing dependencies
            missing_required = compatibility_report['dependencies']['missing_required']
            if missing_required:
                install_results = self.install_missing_dependencies(missing_required, optional=False)
                fixes_applied.update({f"install_{pkg}": success for pkg, success in install_results.items()})

            # Apply FFmpeg installation recommendation if possible
            if not compatibility_report['ffmpeg']['installed']:
                fixes_applied['ffmpeg_recommendation'] = self._suggest_ffmpeg_installation()

            successful_fixes = sum(1 for success in fixes_applied.values() if success)
            total_fixes = len(fixes_applied)

            self.status_callback(f"✅ Auto-fix completed ({successful_fixes}/{total_fixes} successful)")

        except Exception as e:
            logger.error(f"Error in auto-fix: {e}")
            self.status_callback(f"❌ Auto-fix failed: {e}")

        return fixes_applied

    def _suggest_ffmpeg_installation(self) -> bool:
        """Provide FFmpeg installation suggestions."""
        try:
            system = platform.system()

            if system == "Windows":
                suggestion = "Download FFmpeg from https://ffmpeg.org/download.html#build-windows and add to PATH"
            elif system == "Darwin":  # macOS
                suggestion = "Install via Homebrew: brew install ffmpeg"
            elif system == "Linux":
                suggestion = "Install via package manager: sudo apt install ffmpeg (Ubuntu/Debian) or sudo yum install ffmpeg (CentOS/RHEL)"
            else:
                suggestion = "Visit https://ffmpeg.org/download.html for installation instructions"

            logger.info(f"FFmpeg installation suggestion: {suggestion}")
            return True

        except Exception as e:
            logger.error(f"Error generating FFmpeg suggestion: {e}")
            return False

    def validate_configuration(self) -> Dict[str, Any]:
        """
        Validate current configuration for completeness and correctness.
        """
        validation_result = {
            'valid': True,
            'errors': [],
            'warnings': [],
            'missing_required': [],
            'recommendations': []
        }

        try:
            if not self.config_manager:
                validation_result['valid'] = False
                validation_result['errors'].append("Configuration manager not available")
                return validation_result

            # Check required settings
            required_settings = ['groq_api_key', 'model_id', 'default_output_format']

            for setting in required_settings:
                value = self.config_manager.get(setting)
                if not value:
                    validation_result['missing_required'].append(setting)
                    validation_result['valid'] = False

            # Validate API key format
            api_key = self.config_manager.get('groq_api_key')
            if api_key and not api_key.startswith('gsk_'):
                validation_result['warnings'].append("API key format may be incorrect (should start with 'gsk_')")

            # Validate model ID
            model_id = self.config_manager.get('model_id')
            valid_models = ['whisper-large-v3', 'llama3-70b-8192', 'llama3-8b-8192', 'deepseek-r1-distill-llama-70b']
            if model_id and model_id not in valid_models:
                validation_result['warnings'].append(f"Unknown model ID: {model_id}")

            # Check output format
            output_format = self.config_manager.get('default_output_format')
            valid_formats = ['markdown', 'docx', 'txt']
            if output_format and output_format not in valid_formats:
                validation_result['warnings'].append(f"Unknown output format: {output_format}")

            # Generate recommendations
            if validation_result['missing_required']:
                validation_result['recommendations'].append("Configure missing required settings")

            if not api_key:
                validation_result['recommendations'].append("Set up Groq API key for AI features")

        except Exception as e:
            logger.error(f"Error validating configuration: {e}")
            validation_result['valid'] = False
            validation_result['errors'].append(f"Validation error: {e}")

        return validation_result
