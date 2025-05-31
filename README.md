# EchoScribe AI - Smart Voice Note Creator

![EchoScribe AI Logo](./icon.png)

**EchoScribe AI** is an advanced desktop application designed to automatically transform your voice recordings into structured and professional text notes. Powered by cutting-edge AI technology, this application not only transcribes audio but also enhances, formats, and presents it in ready-to-use Word (.docx) documents.

[![Python Version](https://img.shields.io/badge/python-3.7%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## Table of Contents

1.  [Quick Demo](#quick-demo)
2.  [Main Features](#main-features)
3.  [Prerequisites](#prerequisites)
4.  [Installation](#installation)
5.  [Configuration](#configuration)
    *   [FFmpeg](#ffmpeg)
    *   [Groq API Key](#groq-api-key)
6.  [How to Use](#how-to-use)
    *   [Recording Tab](#recording-tab)
    *   [Settings Tab](#settings-tab)
    *   [Output Tab](#output-tab)
    *   [Status Bar](#status-bar)
7.  [How It Works](#how-it-works)
    *   [Main Workflow](#main-workflow)
    *   [Key Components](#key-components)
8.  [Troubleshooting](#troubleshooting)
9.  [Key Dependencies](#key-dependencies)
10. [Contributing](#contributing)
11. [License](#license)

---

## Quick Demo

Here's a glimpse of the EchoScribe AI interface:

![EchoScribe AI Screenshot](https://via.placeholder.com/600x400.png?text=EchoScribe+AI+Application+Screenshot)
*(Replace with a link to your application's screenshot)*

---

## Main Features

*   🎙️ **Flexible Audio Recording**: Record audio from your microphone, system audio (Windows-specific with PyAudioWPatch), or both simultaneously (dual recording).
*   📊 **Real-time Audio Visualization**: Visually monitor your audio input with various modes (waveform, bars, spectrum, fill) and adjustable sensitivity.
*   🔊 **Accurate Transcription**: Choose between Google Speech Recognition or the Groq API with advanced Whisper models for high-quality transcription.
*   🤖 **Smart AI Enhancement**: Leverage the power of Groq AI to automatically:
    *   Analyze audio context (meetings, lectures, dictations, etc.).
    *   Select the optimal AI model for transcription and enhancement.
    *   Reorganize raw transcripts into structured, coherent, and professional notes.
    *   Add relevant titles, subtitles, bullet points, and other formatting.
*   📄 **Export to Word Document (.docx)**: The final output is saved in DOCX format with professional styling, including:
    *   Adaptive document themes based on content type.
    *   Automatic headers and footers.
    *   Heading styling with contextual icons.
    *   Formatting for lists (bullet & numbered), tables, quotes, and code blocks.
    *   Enhanced callout/admonition blocks (Note, Warning, Important, etc.).
*   ⚙️ **Comprehensive Settings**:
    *   Groq API Key Management (use your own or the default API key).
    *   Transcription language selection (Indonesian, English, Japanese, Mandarin).
    *   Economical mode for more cost-effective AI usage.
    *   Output configuration (folder, filename prefix).
    *   Advanced recording settings (long recording support, chunk size, API delay).
*   💾 **Configuration Management**: Application settings are automatically saved and loaded on startup.
*   🛠️ **Error Handling & Logging**: Robust error handling system and detailed logging for troubleshooting.
*   🎨 **Modern & Intuitive Interface**: Built with CustomTkinter for an attractive and user-friendly experience.
*   🛡️ **System Compatibility Checks**: Initial compatibility checks to ensure dependencies are met.

---

## Prerequisites

Before installing and running EchoScribe AI, ensure your system meets the following requirements:

1.  **Python**: Version 3.7 or newer.
2.  **FFmpeg**:
    *   Required by `pydub` for processing certain audio formats.
    *   Ensure FFmpeg is installed and `ffmpeg.exe` (or `ffmpeg` on Linux/macOS) is accessible via your system's PATH, or placed in one of the application's standard search locations (see [FFmpeg Configuration](#ffmpeg)).
    *   Download from [ffmpeg.org](https://ffmpeg.org/download.html).
3.  **Groq API Key (Optional, Highly Recommended)**:
    *   For full Whisper transcription and AI enhancement functionality, using your own Groq API key is highly recommended.
    *   The application provides a default API key with usage limitations.
    *   Obtain an API key from the [Groq Console](https://console.groq.com/).
4.  **Operating System**:
    *   **Windows**: Recommended for full functionality, especially system audio recording (using `PyAudioWPatch`).
    *   **Linux/macOS**: Core features like microphone recording and AI processing will work. System audio recording may require additional configuration or may not function optimally.
5.  **Internet Connection**: Required for cloud transcription services (Google, Groq) and AI enhancements.
6.  **Microphone**: Required for recording audio from a microphone.

---

## Installation

1.  **Clone Repository (If from source code)**:
    ```bash
    git clone https://your_repository_link.git
    cd your_repository_folder_name
    ```

2.  **Create and Activate a Virtual Environment (Recommended)**:
    ```bash
    python -m venv venv
    # Windows
    venv\Scripts\activate
    # macOS/Linux
    source venv/bin/activate
    ```

3.  **Install Dependencies**:
    Ensure you have a `requirements.txt` file or install manually:
    ```bash
    pip install -r requirements.txt
    ```
    If `requirements.txt` is not available, you'll need to install the main libraries:
    ```bash
    pip install speechrecognition pydub customtkinter python-docx groq matplotlib numpy pyaudio sounddevice
    # For Windows, consider installing pyaudiowpatch for better system audio recording:
    pip install pyaudiowpatch
    ```
    See the [Key Dependencies](#key-dependencies) section for a complete list.

4.  **Ensure FFmpeg is Installed**:
    Follow the instructions in the [Prerequisites](#prerequisites) and [FFmpeg Configuration](#ffmpeg) sections.

5.  **Run the Application**:
    ```bash
    python your_main_script_name.py
    ```
    (Replace `your_main_script_name.py` with your main Python script file, e.g., `main.py` or `app.py` as per your code).

---

## Configuration

### FFmpeg

This application uses `pydub`, which relies on FFmpeg to process some audio formats.
The application will attempt to find FFmpeg automatically in the following locations:
*   The same directory as the application's executable file.
*   `C:\FFmpeg\bin\ffmpeg.exe` (Windows)
*   `%USERPROFILE%\FFmpeg\bin\ffmpeg.exe` (Windows)
*   `%LOCALAPPDATA%\FFmpeg\bin\ffmpeg.exe` (Windows)
*   `%PROGRAMFILES%\FFmpeg\bin\ffmpeg.exe` (Windows)
*   `%PROGRAMFILES(X86)%\FFmpeg\bin\ffmpeg.exe` (Windows)
*   Via the system PATH.

**If FFmpeg is not found:**
1.  Download FFmpeg from [https://ffmpeg.org/download.html](https://ffmpeg.org/download.html).
2.  Extract the archive.
3.  Add the `bin` folder (containing `ffmpeg.exe`) to your system's PATH environment variable, OR place `ffmpeg.exe` in one of the locations listed above.

### Groq API Key

To use Whisper transcription and AI enhancement features without limitations, you can configure your own Groq API key.

1.  **Obtain an API Key**:
    *   Visit [https://console.groq.com/](https://console.groq.com/).
    *   Sign up or log in to your account.
    *   Create a new API Key in your dashboard.
    *   Copy the API key.

2.  **Configure in the Application**:
    *   Open the **Settings** tab in the EchoScribe AI application.
    *   Click the **"Manage API Key"** button.
    *   A dialog will appear. Paste your API key into the input field.
    *   Click **"Save"**. The application will attempt to verify your API key.
    *   If you prefer not to use your own API key, you can select **"Use Default"**.

Your custom API key will be stored locally on your computer in a `config.json` file within the `.echoscribe` folder (in your user's home directory).

---

## How to Use

The EchoScribe AI interface is divided into three main tabs: **Recording**, **Settings**, and **Output**.

### Recording Tab

This is the main area for controlling the recording process.

![Recording Tab Screenshot](https://via.placeholder.com/500x350.png?text=Recording+Tab)
*(Replace with a screenshot of your Recording Tab)*

*   **Timer**: Displays the current recording duration (`00:00:00`).
*   **Audio Visualization**: Displays a real-time visual representation of the audio input. You can select the visualization mode (Waveform, Bars, Spectrum, Fill) and adjust its sensitivity using the controls below the visualization area. Visualization can be enabled/disabled.
*   **Start/Stop Recording Button**:
    *   **"Start Recording"**: Click to start a new recording session.
    *   **"Stop Recording"**: Click to end the recording session and begin the transcription and AI enhancement process.
*   **Progress Bar**: Shows the progress of the transcription and enhancement process after recording is stopped.
*   **Quick Settings**:
    *   **Microphone**: Select the desired microphone device from the dropdown list. The `⟳` button next to it refreshes the microphone list.
    *   **Audio Source**:
        *   **Microphone only**: Records only from the selected microphone.
        *   **System audio only**: Records all sound played by your computer system (e.g., from YouTube videos, music, online calls). This feature works optimally on Windows with PyAudioWPatch.
        *   **Microphone + System audio**: Records from both the microphone and system audio simultaneously, merging them into a single recording.

### Settings Tab

Here you can customize various aspects of the application's behavior.

![Settings Tab Screenshot](https://via.placeholder.com/500x350.png?text=Settings+Tab)
*(Replace with a screenshot of your Settings Tab)*

*   **API Configuration**:
    *   **API Key Status**: Displays the status of the currently used API key (Default or Custom).
    *   **Manage API Key**: Opens a dialog to enter, save, or remove your custom Groq API key.
*   **Speech Recognition Settings**:
    *   **Language**: Select the primary language of the audio to be recorded (e.g., `id-ID` for Indonesian, `en-US` for English).
    *   **Recognition Engine**: Choose the transcription engine (`Google` or `Whisper` via Groq).
    *   **Economical Mode**: If checked and using Groq, the application will attempt to use a more economical AI model (e.g., `distil-whisper` for English).
*   **AI Enhancement Settings**:
    *   **Use AI**: Enable or disable the AI note enhancement feature after transcription. If disabled, you will only get the raw transcript.
*   **Output Settings**:
    *   **Output Folder**: Specify the folder where the resulting Word (.docx) documents will be saved. Click **"Browse"** to select a folder.
    *   **Filename Prefix**: Define the prefix to be used for output filenames (e.g., `meeting_notes_`).
*   **Advanced Settings**:
    *   **Long Recording**: Enable support for very long recordings. The application will break the audio into chunks and process them incrementally to avoid memory and API limits.
    *   **Chunk Size**: If "Long Recording" is active, set the duration of each audio chunk (in seconds or minutes) to be processed. Smaller values mean more frequent processing, larger values mean less frequent processing.
    *   **API Delay (seconds)**: Sets the delay between requests to the Groq API to avoid rate limiting.

### Output Tab

This area displays the transcribed and AI-enhanced text results.

![Output Tab Screenshot](https://via.placeholder.com/500x350.png?text=Output+Tab)
*(Replace with a screenshot of your Output Tab)*

*   **Result Display Area**: Displays a preview of the processed text. For long recordings, the preview may only show the initial part of the content. The full content will be available in the `.docx` file.
*   **Export Buttons**:
    *   **Copy to Clipboard**: Copies all text from the display area to the clipboard.
    *   **Export to Word**: Saves the current text from the display area directly to a new Word file (without further AI enhancement from the Recording tab).
    *   **Open Output Folder**: Opens the folder you specified in the Output Settings using your system's file explorer.

### Status Bar

Located at the bottom of the application window, the status bar provides real-time information about:
*   The current application status (Ready, Recording, Processing, Error, etc.).
*   The current system time.

---

## How It Works

EchoScribe AI integrates several technologies and workflows to transform audio into structured notes.

### Main Workflow

1.  **Audio Input**:
    *   The user selects the audio source (microphone, system, or dual).
    *   When recording starts, audio is captured in PCM format.
    *   For "Long Recording", audio is saved in temporary WAV chunks. Otherwise, it's saved as a single temporary WAV file.
    *   During recording, audio data is sent to the visualization module for real-time feedback.

2.  **Recording Stop & Pre-processing**:
    *   When recording stops, audio chunks (if any) or the single audio file are prepared for transcription.
    *   The application detects the audio context (`detect_audio_context`) such as duration, volume level, and silence ratio, to help select the optimal transcription model.

3.  **Transcription (Speech-to-Text)**:
    *   Each audio chunk or single audio file is sent to the selected speech recognition engine:
        *   **Google Speech Recognition**: Uses `recognizer.recognize_google()`.
        *   **Groq API (Whisper)**: Uses `groq_client.audio.transcriptions.create()` with a selected Whisper model (`select_optimal_transcription_model` based on audio context and user choice).
    *   The result is the raw text from the audio.

4.  **AI Enhancement**:
    *   If enabled, the raw text from each chunk (or the entire text if not a long recording) is sent to the Groq API (`groq_client.chat.completions.create()`) for enhancement.
    *   `_analyze_content_characteristics`: The application analyzes content characteristics (type, language, complexity, elements like tables/lists) using a combination of rule-based methods and AI (DeepSeek via Groq) for more accurate classification.
    *   `_select_optimal_model`: Based on content analysis, the most suitable Groq AI model is selected (e.g., a model better suited for technical content, meetings, or narratives).
    *   `_create_content_adaptive_prompts`: Highly specific and adaptive prompts are created for the Groq LLM, instructing the AI on how to structure, format, and enhance the text based on the detected content type.
    *   The AI processes the text to:
        *   Correct grammar and spelling.
        *   Add structure (titles, subtitles, bullet points).
        *   Eliminate redundancy while preserving important details.
        *   Format technical terms, data, etc.
    *   `enhance_document_cohesion`: For long recordings consisting of many chunks, after each chunk is enhanced, the entire combined text can be processed again to improve coherence and flow between sections.

5.  **Word Document (.docx) Generation**:
    *   The enhanced text is then formatted into a Word document using the `python-docx` library.
    *   `_process_markdown_content`: This function acts like an advanced Markdown parser. It parses the enhanced text (which is expected to have Markdown-like syntax from the AI) and translates it into Word elements:
        *   Headings (`#`, `##`, etc.) with contextual icons and adaptive styling.
        *   Bullet and numbered lists (with nested levels).
        *   Task lists (`[ ]`, `[x]`).
        *   Quotes (`> `).
        *   Code blocks (``` ```) with styling.
        *   Tables (Markdown format).
        *   Callout/admonition blocks (`:::note`, `:::warning`, etc.) with visual styling.
        *   Various inline formats (bold, italic, underline, strikethrough, code, highlight, etc.).
    *   `_setup_document_styles`, `_configure_page_layout`, `_add_document_header`, `_add_document_footer`: These functions prepare professional and consistent themes, styles, page layouts, headers, and footers for the document.
    *   `finalize_document_formatting_enhanced`: Performs final touches on document formatting.

6.  **Output & Storage**:
    *   A text preview is displayed in the "Output" tab.
    *   The `.docx` document is saved to the user-specified folder.

### Key Components

*   **`VoiceToMarkdownApp`**: The main application class, managing the UI, state, and workflow.
*   **`APIKeyDialog`**: Dialog for Groq API key management.
*   **`ErrorHandler`**: Centralized class for handling errors and displaying messages to the user.
*   **Configuration Management (`setup_config_management`, `load_config`, `save_config`)**: Saves and loads user preferences.
*   **Audio Recording**:
    *   `record_microphone_audio()`: Records from the microphone.
    *   `record_system_audio()`: Records system audio (relies on `pyaudiowpatch`).
    *   `record_dual_audio()`: Records both and merges them.
    *   `save_audio_chunk()`: Saves audio segments for long recordings.
*   **Audio Visualization**:
    *   Uses `matplotlib` and `numpy` to display real-time waveforms, bars, etc.
    *   `update_visualization_loop()` runs in a separate thread.
*   **Audio Processing**:
    *   `process_audio_thread()`: Main thread for transcription and enhancement.
    *   `process_standard_recording_enhanced()` and `process_extended_recording_optimized()`: Logic for processing short and long recordings.
    *   `transcribe_with_groq_whisper()`: Interface to the Groq transcription API.
*   **AI Enhancement**:
    *   `enhance_with_ai()`: Main function for sending text to the Groq LLM for enhancement.
    *   `_analyze_content_characteristics()`: Analyzes text to determine content type and structure.
    *   `_select_optimal_model()`: Selects the most suitable Groq model.
    *   `_create_content_adaptive_prompts()`: Creates dynamic prompts for the LLM.
    *   `enhance_document_cohesion()`: Improves the overall document flow.
*   **Word Document Generation**:
    *   `save_as_word_document()`: Main function for creating `.docx` files.
    *   `_process_markdown_content()`: Parses the enhanced text and maps it to Word elements.
    *   Various `_add_...` and `_style_...` functions for formatting specific elements (headings, lists, tables, callouts, etc.).

---

## Troubleshooting

*   **FFmpeg not found**:
    *   Ensure FFmpeg is correctly installed and its PATH is set. See the [FFmpeg Configuration](#ffmpeg) section.
    *   The application might still run, but some audio processing features may fail.
*   **System Audio Recording Not Working**:
    *   This feature heavily relies on `PyAudioWPatch` and generally only works on Windows.
    *   Ensure `PyAudioWPatch` is installed.
    *   Check your Windows sound settings. Ensure the default output device is correct and "Stereo Mix" (or similar) is enabled if available.
    *   If it fails, try using "Dual Recording" mode as an alternative, or record system audio using other software and import the file (import feature not yet present, this is general advice).
    *   The application has a troubleshooting dialog (`_show_enhanced_system_audio_troubleshooting`) that will appear if issues arise.
*   **Groq API Key Error**:
    *   Ensure your API key is valid and has sufficient quota.
    *   Check your internet connection.
    *   The API key dialog in Settings allows you to re-enter your key or use the default.
*   **Microphone Not Detected**:
    *   Ensure the microphone is properly connected and permitted by your operating system.
    *   Use the `⟳` (Refresh) button in the Recording tab to update the microphone list.
*   **Low Transcription Quality**:
    *   Ensure the recording environment has minimal noise.
    *   Use a good quality microphone close to the sound source.
    *   Select the appropriate language in Settings.
*   **Application Slow or Freezes When Processing Long Recordings**:
    *   Very long recordings require significant resources. Ensure "Long Recording" is enabled in Settings.
    *   Try reducing the "Chunk Size" for more frequent processing of smaller portions, or increase it to reduce API call frequency (but this increases the risk of timeout if chunks are too large).
*   **Log Files**:
    *   The application saves detailed logs in the `.echoscribe` folder in your user's home directory (e.g., `C:\Users\YourName\.echoscribe\echoscribe.log`). This log file is very useful for diagnosing problems.

---

## Key Dependencies

This application uses several key Python libraries:

*   `customtkinter`: For the modern graphical user interface.
*   `speech_recognition`: For interacting with speech recognition APIs.
*   `pydub` & `wave` & `audioop`: For audio file manipulation and processing.
*   `groq`: For interacting with the Groq API (Whisper and LLMs).
*   `python-docx`: For creating and manipulating Word (.docx) files.
*   `matplotlib` & `numpy`: For real-time audio visualization.
*   `pyaudio` & `sounddevice`: For audio recording and playback.
*   `pyaudiowpatch` (Windows): For improved system audio recording on Windows.
*   `threading`: For background operations to keep the UI responsive.

It is recommended to install all dependencies using `pip install -r requirements.txt` (if provided) or manually.

---

## Contributing

Contributions to EchoScribe AI are highly welcome! If you'd like to contribute, please:
1.  Fork this repository.
2.  Create a new feature branch (`git checkout -b feature/FeatureName`).
3.  Commit your changes (`git commit -am 'Add feature X'`).
4.  Push to the branch (`git push origin feature/FeatureName`).
5.  Create a new Pull Request.

Please ensure your code adheres to quality standards and includes relevant documentation.

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
*(You'll need to create a LICENSE file if it doesn't exist, e.g., with the MIT License content)*


Important Notes for You:

Replace Placeholders:

https://via.placeholder.com/300x150.png?text=EchoScribe+AI+Logo -> Your actual logo URL.

https://via.placeholder.com/600x400.png?text=EchoScribe+AI+Application+Screenshot -> Your actual application screenshot URL.

Similar for the Tab screenshots.

https://example.com/audio_demo.mp3 -> Your actual audio demo URL.

https://your_repository_link.git -> Your Git repository URL.

your_repository_folder_name -> The folder name after cloning.

your_main_script_name.py -> The actual name of your main Python script.

Create a LICENSE file (e.g., with MIT License text).

requirements.txt:
It's highly recommended to generate a requirements.txt file by running:

pip freeze > requirements.txt
IGNORE_WHEN_COPYING_START
content_copy
download
Use code with caution.
Bash
IGNORE_WHEN_COPYING_END

This lists all dependencies and their versions, making it easier for others to install.

Further Customization:

Adjust the Troubleshooting section based on common issues users might face with your specific implementation.

If there are OS-specific configuration steps (beyond FFmpeg), add them.

Ensure all links are valid.
