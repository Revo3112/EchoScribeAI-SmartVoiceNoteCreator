# -*- coding: utf-8 -*-
"""
Device Management Module for EchoScribe AI
Integrated from monolithic system for complete audio device handling.
"""

import logging
import soundcard as sc
import pyaudio
from typing import List, Dict, Any, Optional

logger = logging.getLogger(__name__)

class DeviceManager:
    """
    Enhanced device management integrated from monolithic system.
    Handles both PyAudio and SoundCard device discovery and management.
    """

    def __init__(self):
        self.pyaudio_instance = None
        self.soundcard_devices = []
        self.pyaudio_devices = []
        self._initialize_devices()

    def _initialize_devices(self):
        """Initialize device lists from both PyAudio and SoundCard."""
        try:
            self._refresh_pyaudio_devices()
            self._refresh_soundcard_devices()
            logger.info("Device manager initialized successfully")
        except Exception as e:
            logger.error(f"Error initializing device manager: {e}")

    def _refresh_pyaudio_devices(self):
        """Refresh PyAudio device list."""
        try:
            if self.pyaudio_instance is None:
                self.pyaudio_instance = pyaudio.PyAudio()

            self.pyaudio_devices = []
            device_count = self.pyaudio_instance.get_device_count()

            for i in range(device_count):
                try:
                    device_info = self.pyaudio_instance.get_device_info_by_index(i)

                    # Enhanced device info from monolithic
                    device_data = {
                        'index': i,
                        'name': device_info.get('name', f'Device {i}'),
                        'max_input_channels': device_info.get('maxInputChannels', 0),
                        'max_output_channels': device_info.get('maxOutputChannels', 0),
                        'default_sample_rate': device_info.get('defaultSampleRate', 44100),
                        'host_api': device_info.get('hostApi', 0),
                        'is_input': device_info.get('maxInputChannels', 0) > 0,
                        'is_output': device_info.get('maxOutputChannels', 0) > 0,
                        'type': 'pyaudio'
                    }

                    # Add host API info
                    try:
                        host_info = self.pyaudio_instance.get_host_api_info_by_index(device_info.get('hostApi', 0))
                        device_data['host_api_name'] = host_info.get('name', 'Unknown')
                    except:
                        device_data['host_api_name'] = 'Unknown'

                    self.pyaudio_devices.append(device_data)

                except Exception as e:
                    logger.warning(f"Error getting PyAudio device {i}: {e}")

            logger.info(f"Found {len(self.pyaudio_devices)} PyAudio devices")

        except Exception as e:
            logger.error(f"Error refreshing PyAudio devices: {e}")

    def _refresh_soundcard_devices(self):
        """Refresh SoundCard device list."""
        try:
            self.soundcard_devices = []

            # Get all microphones
            try:
                mics = sc.all_microphones(include_loopback=False)
                for i, mic in enumerate(mics):
                    device_data = {
                        'index': i,
                        'name': mic.name,
                        'id': mic.id,
                        'channels': mic.channels,
                        'is_input': True,
                        'is_output': False,
                        'is_loopback': False,
                        'type': 'soundcard_mic',
                        'device_object': mic
                    }
                    self.soundcard_devices.append(device_data)
            except Exception as e:
                logger.warning(f"Error getting SoundCard microphones: {e}")

            # Get all speakers (for loopback)
            try:
                speakers = sc.all_speakers()
                for i, speaker in enumerate(speakers):
                    device_data = {
                        'index': len(self.soundcard_devices) + i,
                        'name': f"{speaker.name} (Loopback)",
                        'id': speaker.id,
                        'channels': speaker.channels,
                        'is_input': True,  # Can record from it via loopback
                        'is_output': True,
                        'is_loopback': True,
                        'type': 'soundcard_speaker',
                        'device_object': speaker
                    }
                    self.soundcard_devices.append(device_data)
            except Exception as e:
                logger.warning(f"Error getting SoundCard speakers: {e}")

            logger.info(f"Found {len(self.soundcard_devices)} SoundCard devices")

        except Exception as e:
            logger.error(f"Error refreshing SoundCard devices: {e}")

    def get_input_devices(self, include_loopback=True) -> List[Dict[str, Any]]:
        """Get all available input devices."""
        input_devices = []

        # Add PyAudio input devices
        for device in self.pyaudio_devices:
            if device['is_input']:
                input_devices.append({
                    'display_name': f"[PyAudio] {device['name']}",
                    'device_info': device,
                    'source': 'pyaudio'
                })

        # Add SoundCard devices
        for device in self.soundcard_devices:
            if device['is_input']:
                if include_loopback or not device['is_loopback']:
                    input_devices.append({
                        'display_name': f"[SoundCard] {device['name']}",
                        'device_info': device,
                        'source': 'soundcard'
                    })

        return input_devices

    def get_output_devices(self) -> List[Dict[str, Any]]:
        """Get all available output devices."""
        output_devices = []

        # Add PyAudio output devices
        for device in self.pyaudio_devices:
            if device['is_output']:
                output_devices.append({
                    'display_name': f"[PyAudio] {device['name']}",
                    'device_info': device,
                    'source': 'pyaudio'
                })

        # Add SoundCard speakers
        for device in self.soundcard_devices:
            if device['is_output'] and not device['is_loopback']:
                output_devices.append({
                    'display_name': f"[SoundCard] {device['name']}",
                    'device_info': device,
                    'source': 'soundcard'
                })

        return output_devices

    def get_system_audio_devices(self) -> List[Dict[str, Any]]:
        """Get devices that support system audio recording (loopback)."""
        system_devices = []

        for device in self.soundcard_devices:
            if device['is_loopback']:
                system_devices.append({
                    'display_name': device['name'],
                    'device_info': device,
                    'source': 'soundcard'
                })

        return system_devices

    def get_device_by_name(self, name: str, source: str = None) -> Optional[Dict[str, Any]]:
        """Get device by name and optionally source."""
        all_devices = self.pyaudio_devices + self.soundcard_devices

        for device in all_devices:
            if device['name'] == name:
                if source is None or device['type'].startswith(source):
                    return device

        return None

    def get_default_input_device(self) -> Optional[Dict[str, Any]]:
        """Get default input device."""
        try:
            if self.pyaudio_instance:
                default_info = self.pyaudio_instance.get_default_input_device_info()
                return self.get_device_by_name(default_info['name'], 'pyaudio')
        except:
            pass

        # Fallback to first available input device
        input_devices = self.get_input_devices(include_loopback=False)
        if input_devices:
            return input_devices[0]['device_info']

        return None

    def get_default_output_device(self) -> Optional[Dict[str, Any]]:
        """Get default output device."""
        try:
            if self.pyaudio_instance:
                default_info = self.pyaudio_instance.get_default_output_device_info()
                return self.get_device_by_name(default_info['name'], 'pyaudio')
        except:
            pass

        # Fallback to first available output device
        output_devices = self.get_output_devices()
        if output_devices:
            return output_devices[0]['device_info']

        return None

    def test_device(self, device_info: Dict[str, Any], duration: float = 1.0) -> bool:
        """Test if a device can be used for recording."""
        try:
            if device_info['type'].startswith('pyaudio'):
                return self._test_pyaudio_device(device_info, duration)
            elif device_info['type'].startswith('soundcard'):
                return self._test_soundcard_device(device_info, duration)
        except Exception as e:
            logger.error(f"Error testing device {device_info['name']}: {e}")
            return False

        return False

    def _test_pyaudio_device(self, device_info: Dict[str, Any], duration: float) -> bool:
        """Test PyAudio device."""
        try:
            if not self.pyaudio_instance:
                return False

            # Try to open a stream
            stream = self.pyaudio_instance.open(
                format=pyaudio.paInt16,
                channels=min(device_info['max_input_channels'], 2),
                rate=int(device_info['default_sample_rate']),
                input=True,
                input_device_index=device_info['index'],
                frames_per_buffer=1024
            )

            # Try to read some data
            stream.read(int(device_info['default_sample_rate'] * duration * 0.1))
            stream.stop_stream()
            stream.close()

            return True

        except Exception as e:
            logger.debug(f"PyAudio device test failed: {e}")
            return False

    def _test_soundcard_device(self, device_info: Dict[str, Any], duration: float) -> bool:
        """Test SoundCard device."""
        try:
            device_obj = device_info['device_object']

            # Try to record a small sample
            with device_obj.recorder(samplerate=48000, channels=device_info['channels']) as recorder:
                data = recorder.record(numframes=int(48000 * duration * 0.1))
                return len(data) > 0

        except Exception as e:
            logger.debug(f"SoundCard device test failed: {e}")
            return False

    def refresh_devices(self):
        """Refresh all device lists."""
        logger.info("Refreshing device lists...")
        self._refresh_pyaudio_devices()
        self._refresh_soundcard_devices()

    def get_device_info_string(self, device_info: Dict[str, Any]) -> str:
        """Get detailed device information as string."""
        info_parts = [
            f"Name: {device_info['name']}",
            f"Type: {device_info['type']}",
            f"Index: {device_info['index']}"
        ]

        if device_info['type'].startswith('pyaudio'):
            info_parts.extend([
                f"Input Channels: {device_info['max_input_channels']}",
                f"Output Channels: {device_info['max_output_channels']}",
                f"Sample Rate: {device_info['default_sample_rate']}",
                f"Host API: {device_info.get('host_api_name', 'Unknown')}"
            ])
        elif device_info['type'].startswith('soundcard'):
            info_parts.extend([
                f"Channels: {device_info['channels']}",
                f"ID: {device_info['id']}",
                f"Loopback: {'Yes' if device_info['is_loopback'] else 'No'}"
            ])

        return "\n".join(info_parts)

    def cleanup(self):
        """Cleanup resources."""
        try:
            if self.pyaudio_instance:
                self.pyaudio_instance.terminate()
                self.pyaudio_instance = None
            logger.info("Device manager cleaned up successfully")
        except Exception as e:
            logger.error(f"Error during device manager cleanup: {e}")

    def __del__(self):
        """Destructor to ensure cleanup."""
        self.cleanup()
