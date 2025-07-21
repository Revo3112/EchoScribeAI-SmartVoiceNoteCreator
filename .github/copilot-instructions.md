# EchoScribe AI - Smart Voice Note Creator

## Project Overview

EchoScribe AI is a sophisticated voice-to-text application that transforms audio recordings into professionally formatted documents using AI enhancement. The project features a **modular architecture** migrated from a monolithic system (~14,416 lines) with enhanced functionality including real-time audio visualization, intelligent content analysis, and enterprise-grade error handling.

## Production Status
- **Architecture**: Modular design with service orchestration
- **Entry Point**: `python main_enhanced.py` (production-ready UI)
- **Legacy Reference**: `EchoScribe AI - Smart Voice Note Creator.py` (monolithic fallback)
- **Dependencies**: CustomTkinter UI, Groq AI, soundcard audio, python-docx documents

## Architecture & Core Components

### Entry Points & Execution
- **Primary**: `python main_enhanced.py` → `EchoScribeUI` (CustomTkinter interface)
- **Legacy**: `EchoScribe AI - Smart Voice Note Creator.py` (monolithic reference)
- **Dependencies**: Install via `pip install -r requirements_new.txt`

### Modular Structure
```
src/
├── app_controller.py          # Service orchestrator (EchoScribeApp class)
├── config/
│   ├── config_manager.py      # ConfigManager: ~/.echoscribe/config.json
│   └── migrator.py           # System compatibility & legacy migration
├── audio/
│   ├── recorder.py           # AudioRecorder: mic/system/dual recording
│   └── visualizer.py         # AudioVisualizer: real-time 4-mode display
├── ai/
│   └── ai_service.py         # AIService: Groq API + content analysis
├── document/
│   └── processor.py          # DocumentProcessor: Word/Markdown export
├── ui/
│   ├── main_window.py        # EchoScribeUI: main interface
│   ├── device_manager.py     # Audio device management
│   ├── components/           # ValueTrackingSlider, ApiKeyDialog
│   ├── tabs/                 # recording_tab.py, settings_tab.py
│   └── services/folder_service.py
└── utils/
    └── error_handler.py      # ErrorHandler: 6 error categories + fallbacks
```

## Development Patterns

### Service Orchestration (app_controller.py)
```python
class EchoScribeApp:
    def __init__(self, status_callback: Optional[Callable[[str], None]] = None):
        self.config = ConfigManager()
        self.error_handler = ErrorHandler(status_callback)
        self.audio_recorder = AudioRecorder(config.get_audio_config(), status_callback)
        self.ai_service = AIService(config.get_user_api_key())
        self.document_processor = DocumentProcessor(config.get_doc_config())
```

### Status Callback Pattern (Universal)
Every module constructor accepts: `status_callback: Optional[Callable[[str], None]]`
```python
self.status_callback = status_callback or (lambda x: None)
# Usage: self.status_callback("Processing audio...")
```

### API Key Management (Production Security)
```python
# config_manager.py - User-provided keys only
def save_user_api_key(self, api_key: str) -> bool:
    if not api_key.startswith("gsk_"):
        return False
    # Secure storage: ~/.echoscribe/config.json (NOT in repo)
```

### Audio Recording Patterns (From Working Test Files)
```python
# audio/recorder.py - Based on test_for_device_and_microphone_v2.py
def record_dual_audio(self) -> Optional[str]:
    # Parallel threads: loopback_thread + mic_thread
    # Audio mixing: channel detection → stereo conversion → normalization
    # Uses soundcard.get_microphone(include_loopback=True)

def record_system_audio(self) -> Optional[str]:
    # Loopback recording with COM initialization and fallback
    # Reference: test_for_device_only.py proven pattern
```

## Critical Dependencies & Implementation References

### Core Dependencies
- **Audio**: `soundcard` (primary), `pyaudio`, `numpy` - Working patterns in test files
- **AI**: `groq` client for Whisper transcription and LLM enhancement
- **Document**: `python-docx`, `markdown` for multiple output formats
- **UI**: `customtkinter` with dark theme and modern components
- **Visualization**: `matplotlib` for real-time audio visualization

### Monolithic Source Code Reference
**File**: `EchoScribe AI - Smart Voice Note Creator.py` (~14,416 lines)

**CRITICAL**: All modular implementations must reference the monolithic patterns:
- Audio Recording (lines 2439-3700): Complete loopback/mic/dual recording
- AI Service (lines 4320-5100): Groq integration with error handling
- Document Processing (lines 6800-13400): 30+ formatting patterns
- UI Components (lines 299-1300): Visualization and enhanced controls

### Working Audio Test Files
**Essential for audio module development:**
1. `test_for_device_only.py`: Basic loopback using `soundcard.get_microphone(include_loopback=True)`
2. `test_for_device_and_microphone_v1.py`: Dual recording foundation with threading
3. `test_for_device_and_microphone_v2.py`: Complete dual recording with audio mixing
## Critical Workflows

### Entry Points
- **Primary**: `python main_enhanced.py` (modular architecture with CustomTkinter UI)
- **Legacy**: `EchoScribe AI - Smart Voice Note Creator.py` (monolithic backup, ~14,400 lines)

### Audio Recording Modes
1. **Microphone**: Standard soundcard recording
2. **System**: Loopback recording using `soundcard.get_microphone(include_loopback=True)`
3. **Dual**: Parallel threads with audio mixing and normalization

### Configuration & API Security
- **First Run**: UI prompts for Groq API key (gsk_* validation required)
- **Storage**: `~/.echoscribe/config.json` (secure local storage)
- **Critical**: All API keys are user-provided and validated - NO hardcoded keys

### Processing Pipeline
```python
# app_controller.py orchestrates this workflow:
1. Audio Recording → AudioRecorder.record_*_audio()
2. Transcription → AIService.transcribe_audio()
3. Enhancement → AIService.enhance_text()
4. Document Creation → DocumentProcessor.create_*_document()
```

## Development Guidelines

### Project-Specific Patterns

#### Module Constructor Pattern
```python
# Every module follows this pattern:
class ServiceModule:
    def __init__(self, config: ConfigManager, status_callback: Optional[Callable] = None):
        self.config = config
        self.status_callback = status_callback or (lambda x: None)
        # Initialize specific service components
```

#### Error Handling Integration
```python
# All modules use ErrorHandler for consistent error management:
from src.utils.error_handler import ErrorHandler

try:
    # Service operation
    pass
except Exception as e:
    self.error_handler.handle_error("error_category", e)
```

#### UI Component Development
```python
# UI components integrate with main_window.py:
# - Use CustomTkinter (ctk) for modern styling
# - Implement ValueTrackingSlider for controls with tooltips
# - Support real-time status updates via status_callback
```

### Audio Implementation Requirements
**CRITICAL**: All audio modules must reference working test file patterns:
- Use `soundcard` library as primary audio interface
- Reference `test_for_device_and_microphone_v2.py` for dual recording
- Implement thread-safe audio mixing for system + microphone
- Follow proven device detection patterns from test files

### Configuration Management
- **Secure**: User API keys only, stored in `~/.echoscribe/config.json`
- **Migration**: Auto-migrate legacy configs from monolithic version
- **Validation**: All API keys must start with "gsk_" for Groq validation

### Testing & Debugging
- **Audio Debug**: Use existing test files as reference implementations
- **Module Testing**: Individual module validation through imports
- **Common Issues**: Check soundcard installation, API key validation, module imports
