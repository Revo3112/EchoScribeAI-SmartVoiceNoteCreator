# -*- coding: utf-8 -*-
"""
Settings Tab Module for EchoScribe AI
Integrated from monolithic system with complete settings interface.
"""

import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox, colorchooser
import logging
from pathlib import Path
from typing import Optional, Callable

from ..device_manager import DeviceManager

logger = logging.getLogger(__name__)

class SettingsTab(ctk.CTkFrame):
    """
    Enhanced settings tab integrated from monolithic system.
    Provides comprehensive application settings management.
    """

    def __init__(self, master, config_manager, status_callback: Optional[Callable] = None):
        super().__init__(master)

        self.config = config_manager
        self.status_callback = status_callback

        # Settings variables
        self.setup_variables()
        self.setup_ui()
        self.load_settings()

    def setup_variables(self):
        """Setup all settings variables."""
        # Audio settings
        self.sample_rate = tk.IntVar(value=48000)
        self.channels = tk.IntVar(value=2)
        self.blocksize = tk.IntVar(value=1024)

        # AI settings
        self.language = tk.StringVar(value="id-ID - Bahasa Indonesia")
        self.engine = tk.StringVar(value="Google")
        self.use_ai_enhancement = tk.BooleanVar(value=True)
        self.use_economic_model = tk.BooleanVar(value=False)
        self.max_tokens = tk.IntVar(value=4000)
        self.api_request_delay = tk.IntVar(value=10)

        # Output settings
        self.output_folder = tk.StringVar(value=str(Path.home() / "Documents"))
        self.filename_prefix = tk.StringVar(value="catatan")

        # UI settings
        self.theme = tk.StringVar(value="dark")
        self.viz_enabled = tk.BooleanVar(value=True)
        self.viz_mode = tk.StringVar(value="waveform")
        self.viz_sensitivity = tk.DoubleVar(value=1.0)

        # Document formatting
        self.heading_spacing_before = tk.IntVar(value=12)
        self.heading_spacing_after = tk.IntVar(value=6)
        self.paragraph_spacing = tk.IntVar(value=6)

    def setup_ui(self):
        """Setup the settings interface."""

        # Main scrollable frame
        main_container = ctk.CTkScrollableFrame(self)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Title
        title_label = ctk.CTkLabel(
            main_container,
            text="⚙️ Pengaturan Aplikasi",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 30))

        # Settings sections
        self.setup_audio_settings(main_container)
        self.setup_ai_settings(main_container)
        self.setup_output_settings(main_container)
        self.setup_ui_settings(main_container)
        self.setup_document_settings(main_container)
        self.setup_advanced_settings(main_container)

        # Control buttons
        self.setup_control_buttons(main_container)

    def setup_audio_settings(self, parent):
        """Setup audio settings section."""
        audio_frame = ctk.CTkFrame(parent)
        audio_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            audio_frame,
            text="🎵 Pengaturan Audio",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 15))

        settings_container = ctk.CTkFrame(audio_frame, fg_color="transparent")
        settings_container.pack(fill=tk.X, padx=20, pady=(0, 20))

        # Sample rate
        sample_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        sample_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(sample_frame, text="Sample Rate (Hz):", width=150).pack(side=tk.LEFT)

        sample_rates = ["8000", "16000", "22050", "44100", "48000", "96000"]
        sample_dropdown = ctk.CTkComboBox(
            sample_frame,
            values=sample_rates,
            variable=self.sample_rate,
            width=120,
            state="readonly"
        )
        sample_dropdown.pack(side=tk.LEFT, padx=(10, 0))

        # Channels
        channels_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        channels_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(channels_frame, text="Channels:", width=150).pack(side=tk.LEFT)

        channels_dropdown = ctk.CTkComboBox(
            channels_frame,
            values=["1", "2"],
            variable=self.channels,
            width=120,
            state="readonly"
        )
        channels_dropdown.pack(side=tk.LEFT, padx=(10, 0))

        # Block size
        block_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        block_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(block_frame, text="Block Size:", width=150).pack(side=tk.LEFT)

        block_sizes = ["256", "512", "1024", "2048", "4096"]
        block_dropdown = ctk.CTkComboBox(
            block_frame,
            values=block_sizes,
            variable=self.blocksize,
            width=120,
            state="readonly"
        )
        block_dropdown.pack(side=tk.LEFT, padx=(10, 0))

    def setup_ai_settings(self, parent):
        """Setup AI settings section."""
        ai_frame = ctk.CTkFrame(parent)
        ai_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            ai_frame,
            text="🤖 Pengaturan AI",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 15))

        settings_container = ctk.CTkFrame(ai_frame, fg_color="transparent")
        settings_container.pack(fill=tk.X, padx=20, pady=(0, 20))

        # API Key management
        api_frame = ctk.CTkFrame(settings_container)
        api_frame.pack(fill=tk.X, pady=(0, 15))

        ctk.CTkLabel(
            api_frame,
            text="API Key Groq:",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 5))

        api_status = self.get_api_key_status()
        status_label = ctk.CTkLabel(
            api_frame,
            text=api_status,
            text_color="green" if "Tersimpan" in api_status else "red"
        )
        status_label.pack(anchor="w", padx=15, pady=(0, 10))

        api_btn_frame = ctk.CTkFrame(api_frame, fg_color="transparent")
        api_btn_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

        manage_api_btn = ctk.CTkButton(
            api_btn_frame,
            text="🔑 Kelola API Key",
            command=self.manage_api_key,
            width=150
        )
        manage_api_btn.pack(side=tk.LEFT)

        # Language selection
        lang_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        lang_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(lang_frame, text="Bahasa:", width=150).pack(side=tk.LEFT)

        languages = [
            ("id-ID", "Bahasa Indonesia"),
            ("en-US", "English (US)")
        ]

        lang_values = [f"{code} - {name}" for code, name in languages]
        lang_dropdown = ctk.CTkComboBox(
            lang_frame,
            values=lang_values,
            variable=self.language,
            width=200,
            state="readonly",
            command=self.on_language_change
        )
        lang_dropdown.pack(side=tk.LEFT, padx=(10, 0))

        # Engine selection
        engine_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        engine_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(engine_frame, text="Engine:", width=150).pack(side=tk.LEFT)

        engine_dropdown = ctk.CTkComboBox(
            engine_frame,
            values=["Google", "Whisper", "Azure"],
            variable=self.engine,
            width=120,
            state="readonly"
        )
        engine_dropdown.pack(side=tk.LEFT, padx=(10, 0))

        # AI Enhancement checkbox
        enhancement_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        enhancement_frame.pack(fill=tk.X, pady=(0, 10))

        enhancement_checkbox = ctk.CTkCheckBox(
            enhancement_frame,
            text="Gunakan Peningkatan AI",
            variable=self.use_ai_enhancement,
            font=ctk.CTkFont(size=14)
        )
        enhancement_checkbox.pack(anchor="w")

        # Economic model checkbox
        economic_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        economic_frame.pack(fill=tk.X, pady=(0, 10))

        economic_checkbox = ctk.CTkCheckBox(
            economic_frame,
            text="Gunakan Model Ekonomis",
            variable=self.use_economic_model,
            font=ctk.CTkFont(size=14)
        )
        economic_checkbox.pack(anchor="w")

        # Max tokens slider
        tokens_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        tokens_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(tokens_frame, text="Max Tokens:", width=150).pack(side=tk.LEFT)

        tokens_slider = ctk.CTkSlider(
            tokens_frame,
            from_=1000,
            to=8000,
            variable=self.max_tokens,
            width=200
        )
        tokens_slider.pack(side=tk.LEFT, padx=(10, 0))

        self.tokens_label = ctk.CTkLabel(tokens_frame, text=f"{self.max_tokens.get()}")
        self.tokens_label.pack(side=tk.LEFT, padx=(10, 0))

        self.max_tokens.trace_add("write", self.update_tokens_label)

    def setup_output_settings(self, parent):
        """Setup output settings section."""
        output_frame = ctk.CTkFrame(parent)
        output_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            output_frame,
            text="📁 Pengaturan Output",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 15))

        settings_container = ctk.CTkFrame(output_frame, fg_color="transparent")
        settings_container.pack(fill=tk.X, padx=20, pady=(0, 20))

        # Output folder
        folder_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        folder_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(folder_frame, text="Folder Output:", width=120).pack(side=tk.LEFT)

        self.folder_entry = ctk.CTkEntry(
            folder_frame,
            textvariable=self.output_folder,
            width=300
        )
        self.folder_entry.pack(side=tk.LEFT, padx=(10, 0))

        browse_btn = ctk.CTkButton(
            folder_frame,
            text="📁 Browse",
            command=self.browse_output_folder,
            width=80
        )
        browse_btn.pack(side=tk.LEFT, padx=(10, 0))

        # Filename prefix
        prefix_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        prefix_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(prefix_frame, text="Prefix Filename:", width=120).pack(side=tk.LEFT)

        prefix_entry = ctk.CTkEntry(
            prefix_frame,
            textvariable=self.filename_prefix,
            width=200
        )
        prefix_entry.pack(side=tk.LEFT, padx=(10, 0))

    def setup_ui_settings(self, parent):
        """Setup UI settings section."""
        ui_frame = ctk.CTkFrame(parent)
        ui_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            ui_frame,
            text="🎨 Pengaturan Interface",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 15))

        settings_container = ctk.CTkFrame(ui_frame, fg_color="transparent")
        settings_container.pack(fill=tk.X, padx=20, pady=(0, 20))

        # Theme selection
        theme_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        theme_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(theme_frame, text="Tema:", width=150).pack(side=tk.LEFT)

        theme_dropdown = ctk.CTkComboBox(
            theme_frame,
            values=["dark", "light"],
            variable=self.theme,
            width=120,
            state="readonly",
            command=self.on_theme_change
        )
        theme_dropdown.pack(side=tk.LEFT, padx=(10, 0))

        # Visualization settings
        viz_frame = ctk.CTkFrame(settings_container)
        viz_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(
            viz_frame,
            text="Visualisasi Audio:",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        viz_enable_frame = ctk.CTkFrame(viz_frame, fg_color="transparent")
        viz_enable_frame.pack(fill=tk.X, padx=15, pady=(0, 10))

        viz_checkbox = ctk.CTkCheckBox(
            viz_enable_frame,
            text="Aktifkan Visualisasi",
            variable=self.viz_enabled,
            command=self.on_viz_enable_change
        )
        viz_checkbox.pack(anchor="w")

        viz_mode_frame = ctk.CTkFrame(viz_frame, fg_color="transparent")
        viz_mode_frame.pack(fill=tk.X, padx=15, pady=(0, 10))

        ctk.CTkLabel(viz_mode_frame, text="Mode:", width=100).pack(side=tk.LEFT)

        self.viz_mode_dropdown = ctk.CTkComboBox(
            viz_mode_frame,
            values=["waveform", "bars", "spectrum", "fill"],
            variable=self.viz_mode,
            width=120,
            state="readonly"
        )
        self.viz_mode_dropdown.pack(side=tk.LEFT, padx=(10, 0))

        viz_sens_frame = ctk.CTkFrame(viz_frame, fg_color="transparent")
        viz_sens_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

        ctk.CTkLabel(viz_sens_frame, text="Sensitivitas:", width=100).pack(side=tk.LEFT)

        self.viz_sens_slider = ctk.CTkSlider(
            viz_sens_frame,
            from_=0.1,
            to=3.0,
            variable=self.viz_sensitivity,
            width=150
        )
        self.viz_sens_slider.pack(side=tk.LEFT, padx=(10, 0))

        self.viz_sens_label = ctk.CTkLabel(viz_sens_frame, text=f"{self.viz_sensitivity.get():.1f}")
        self.viz_sens_label.pack(side=tk.LEFT, padx=(10, 0))

        self.viz_sensitivity.trace_add("write", self.update_viz_sens_label)

    def setup_document_settings(self, parent):
        """Setup document formatting settings."""
        doc_frame = ctk.CTkFrame(parent)
        doc_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            doc_frame,
            text="📄 Pengaturan Dokumen",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 15))

        settings_container = ctk.CTkFrame(doc_frame, fg_color="transparent")
        settings_container.pack(fill=tk.X, padx=20, pady=(0, 20))

        # Heading spacing before
        heading_before_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        heading_before_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(heading_before_frame, text="Spasi Sebelum Heading:", width=180).pack(side=tk.LEFT)

        heading_before_slider = ctk.CTkSlider(
            heading_before_frame,
            from_=0,
            to=24,
            variable=self.heading_spacing_before,
            width=150
        )
        heading_before_slider.pack(side=tk.LEFT, padx=(10, 0))

        self.heading_before_label = ctk.CTkLabel(heading_before_frame, text=f"{self.heading_spacing_before.get()}")
        self.heading_before_label.pack(side=tk.LEFT, padx=(10, 0))

        # Heading spacing after
        heading_after_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        heading_after_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(heading_after_frame, text="Spasi Setelah Heading:", width=180).pack(side=tk.LEFT)

        heading_after_slider = ctk.CTkSlider(
            heading_after_frame,
            from_=0,
            to=24,
            variable=self.heading_spacing_after,
            width=150
        )
        heading_after_slider.pack(side=tk.LEFT, padx=(10, 0))

        self.heading_after_label = ctk.CTkLabel(heading_after_frame, text=f"{self.heading_spacing_after.get()}")
        self.heading_after_label.pack(side=tk.LEFT, padx=(10, 0))

        # Paragraph spacing
        para_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        para_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(para_frame, text="Spasi Paragraf:", width=180).pack(side=tk.LEFT)

        para_slider = ctk.CTkSlider(
            para_frame,
            from_=0,
            to=18,
            variable=self.paragraph_spacing,
            width=150
        )
        para_slider.pack(side=tk.LEFT, padx=(10, 0))

        self.para_label = ctk.CTkLabel(para_frame, text=f"{self.paragraph_spacing.get()}")
        self.para_label.pack(side=tk.LEFT, padx=(10, 0))

        # Add trace callbacks for document spacing
        self.heading_spacing_before.trace_add("write", self.update_heading_before_label)
        self.heading_spacing_after.trace_add("write", self.update_heading_after_label)
        self.paragraph_spacing.trace_add("write", self.update_para_label)

    def setup_advanced_settings(self, parent):
        """Setup advanced settings section."""
        advanced_frame = ctk.CTkFrame(parent)
        advanced_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            advanced_frame,
            text="🔧 Pengaturan Lanjutan",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 15))

        settings_container = ctk.CTkFrame(advanced_frame, fg_color="transparent")
        settings_container.pack(fill=tk.X, padx=20, pady=(0, 20))

        # API request delay
        delay_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        delay_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(delay_frame, text="Delay Request API (detik):", width=180).pack(side=tk.LEFT)

        delay_slider = ctk.CTkSlider(
            delay_frame,
            from_=1,
            to=30,
            variable=self.api_request_delay,
            width=150
        )
        delay_slider.pack(side=tk.LEFT, padx=(10, 0))

        self.delay_label = ctk.CTkLabel(delay_frame, text=f"{self.api_request_delay.get()}")
        self.delay_label.pack(side=tk.LEFT, padx=(10, 0))

        self.api_request_delay.trace_add("write", self.update_delay_label)

    def setup_control_buttons(self, parent):
        """Setup control buttons."""
        button_frame = ctk.CTkFrame(parent, fg_color="transparent")
        button_frame.pack(fill=tk.X, pady=30)

        # Reset button
        reset_btn = ctk.CTkButton(
            button_frame,
            text="🔄 Reset ke Default",
            command=self.reset_to_defaults,
            fg_color="orange",
            hover_color="darkorange",
            width=150
        )
        reset_btn.pack(side=tk.LEFT, padx=(0, 20))

        # Export/Import buttons
        export_btn = ctk.CTkButton(
            button_frame,
            text="📤 Export Settings",
            command=self.export_settings,
            width=150
        )
        export_btn.pack(side=tk.LEFT, padx=(0, 10))

        import_btn = ctk.CTkButton(
            button_frame,
            text="📥 Import Settings",
            command=self.import_settings,
            width=150
        )
        import_btn.pack(side=tk.LEFT, padx=(0, 20))

        # Save button
        save_btn = ctk.CTkButton(
            button_frame,
            text="💾 Simpan Pengaturan",
            command=self.save_settings,
            width=150
        )
        save_btn.pack(side=tk.RIGHT)

    # Event handlers and utility methods
    def get_api_key_status(self):
        """Get API key status message."""
        if self.config.has_user_api_key():
            api_key = self.config.get_user_api_key()
            return f"✅ API Key Tersimpan: {api_key[:15]}...{api_key[-5:]}"
        else:
            return "❌ Tidak Ada API Key"

    def manage_api_key(self):
        """Open API key management dialog."""
        from ..main_window import APIKeyDialog

        current_key = self.config.get_user_api_key() or ""
        dialog = APIKeyDialog(self, self.config, current_key)

        # Wait for dialog to close and refresh status
        self.wait_window(dialog)
        self.refresh_api_status()

    def refresh_api_status(self):
        """Refresh API key status display."""
        # Find and update the API status label
        # This is a simplified approach - in a real implementation,
        # you'd store a reference to the status label
        self.update_status("API key status updated")

    def on_language_change(self, value):
        """Handle language selection change."""
        # Value is already in correct format "code - name", keep it as is
        # No need to modify self.language since ComboBox already set it
        logger.info(f"Language changed to: {value}")

        # Save the change to config
        self.save_language_setting(value)

    def save_language_setting(self, value):
        """Save language setting to config."""
        try:
            # Extract language code for config storage
            lang_code = value.split(" - ")[0] if " - " in value else value
            self.config.set("language", lang_code)
            logger.info(f"Language saved to config: {lang_code}")
        except Exception as e:
            logger.error(f"Error saving language setting: {e}")

    def on_theme_change(self, value):
        """Handle theme change."""
        ctk.set_appearance_mode(value)
        logger.info(f"Theme changed to: {value}")

    def on_viz_enable_change(self):
        """Handle visualization enable/disable."""
        enabled = self.viz_enabled.get()
        state = "normal" if enabled else "disabled"

        self.viz_mode_dropdown.configure(state=state)
        self.viz_sens_slider.configure(state=state)

    def browse_output_folder(self):
        """Browse for output folder - simplified monolithic implementation."""
        folder = filedialog.askdirectory(initialdir=self.output_folder.get())
        if folder:
            self.output_folder.set(folder)

    # Label update methods
    def update_tokens_label(self, *args):
        """Update max tokens label."""
        self.tokens_label.configure(text=f"{self.max_tokens.get()}")

    def update_viz_sens_label(self, *args):
        """Update visualization sensitivity label."""
        self.viz_sens_label.configure(text=f"{self.viz_sensitivity.get():.1f}")

    def update_heading_before_label(self, *args):
        """Update heading before spacing label."""
        self.heading_before_label.configure(text=f"{self.heading_spacing_before.get()}")

    def update_heading_after_label(self, *args):
        """Update heading after spacing label."""
        self.heading_after_label.configure(text=f"{self.heading_spacing_after.get()}")

    def update_para_label(self, *args):
        """Update paragraph spacing label."""
        self.para_label.configure(text=f"{self.paragraph_spacing.get()}")

    def update_delay_label(self, *args):
        """Update API delay label."""
        self.delay_label.configure(text=f"{self.api_request_delay.get()}")

    def load_settings(self):
        """Load settings from configuration."""
        try:
            # Load audio settings
            audio_config = self.config.get_recording_config()
            self.sample_rate.set(audio_config.get('sample_rate', 48000))
            self.channels.set(audio_config.get('channels', 2))
            self.blocksize.set(audio_config.get('blocksize', 1024))

            # Load AI settings
            ai_config = self.config.get_ai_config()
            lang_code = ai_config.get('language', 'id-ID')

            # Convert language code to display format
            if lang_code == 'id-ID':
                self.language.set('id-ID - Bahasa Indonesia')
            elif lang_code == 'en-US':
                self.language.set('en-US - English (US)')
            else:
                self.language.set('id-ID - Bahasa Indonesia')  # Default fallback

            self.engine.set(ai_config.get('engine', 'Google'))
            self.use_ai_enhancement.set(ai_config.get('use_ai_enhancement', True))
            self.use_economic_model.set(ai_config.get('use_economic_model', False))
            self.max_tokens.set(ai_config.get('max_tokens', 4000))
            self.api_request_delay.set(ai_config.get('api_request_delay', 10))

            # Load output settings
            output_config = self.config.get_output_config()
            self.output_folder.set(output_config.get('output_folder', str(Path.home() / "Documents")))
            self.filename_prefix.set(output_config.get('filename_prefix', 'catatan'))
            self.heading_spacing_before.set(output_config.get('heading_spacing_before', 12))
            self.heading_spacing_after.set(output_config.get('heading_spacing_after', 6))
            self.paragraph_spacing.set(output_config.get('paragraph_spacing', 6))

            # Load UI settings
            ui_config = self.config.get_ui_config()
            self.theme.set(ui_config.get('theme', 'dark'))
            self.viz_enabled.set(ui_config.get('viz_enabled', True))
            self.viz_mode.set(ui_config.get('viz_mode', 'waveform'))
            self.viz_sensitivity.set(ui_config.get('viz_sensitivity', 1.0))

            # Update UI state
            self.on_viz_enable_change()

            logger.info("Settings loaded successfully")

        except Exception as e:
            logger.error(f"Error loading settings: {e}")
            self.update_status("Gagal memuat pengaturan")

    def save_settings(self):
        """Save current settings to configuration."""
        try:
            settings = {
                # Audio settings
                'sample_rate': self.sample_rate.get(),
                'channels': self.channels.get(),
                'blocksize': self.blocksize.get(),

                # AI settings
                'language': self.language.get().split(" - ")[0] if " - " in self.language.get() else self.language.get(),
                'engine': self.engine.get(),
                'use_ai_enhancement': self.use_ai_enhancement.get(),
                'use_economic_model': self.use_economic_model.get(),
                'max_tokens': self.max_tokens.get(),
                'api_request_delay': self.api_request_delay.get(),

                # Output settings
                'output_folder': self.output_folder.get(),
                'filename_prefix': self.filename_prefix.get(),
                'heading_spacing_before': self.heading_spacing_before.get(),
                'heading_spacing_after': self.heading_spacing_after.get(),
                'paragraph_spacing': self.paragraph_spacing.get(),

                # UI settings
                'theme': self.theme.get(),
                'viz_enabled': self.viz_enabled.get(),
                'viz_mode': self.viz_mode.get(),
                'viz_sensitivity': self.viz_sensitivity.get()
            }

            self.config.update(settings)

            messagebox.showinfo("Sukses", "Pengaturan berhasil disimpan!")
            self.update_status("Pengaturan disimpan")
            logger.info("Settings saved successfully")

        except Exception as e:
            logger.error(f"Error saving settings: {e}")
            messagebox.showerror("Error", f"Gagal menyimpan pengaturan:\n{str(e)}")
            self.update_status("Gagal menyimpan pengaturan")

    def reset_to_defaults(self):
        """Reset all settings to defaults."""
        if messagebox.askyesno("Konfirmasi", "Apakah Anda yakin ingin mereset semua pengaturan ke default?"):
            try:
                self.config.reset_to_defaults()
                self.load_settings()

                messagebox.showinfo("Sukses", "Pengaturan berhasil direset ke default!")
                self.update_status("Pengaturan direset")
                logger.info("Settings reset to defaults")

            except Exception as e:
                logger.error(f"Error resetting settings: {e}")
                messagebox.showerror("Error", f"Gagal mereset pengaturan:\n{str(e)}")

    def export_settings(self):
        """Export settings to file."""
        try:
            file_path = filedialog.asksaveasfilename(
                title="Export Pengaturan",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )

            if file_path:
                if self.config.export_config(file_path):
                    messagebox.showinfo("Sukses", f"Pengaturan berhasil diekspor ke:\n{file_path}")
                    self.update_status("Pengaturan diekspor")
                else:
                    messagebox.showerror("Error", "Gagal mengekspor pengaturan")

        except Exception as e:
            logger.error(f"Error exporting settings: {e}")
            messagebox.showerror("Error", f"Gagal mengekspor pengaturan:\n{str(e)}")

    def import_settings(self):
        """Import settings from file."""
        try:
            file_path = filedialog.askopenfilename(
                title="Import Pengaturan",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )

            if file_path:
                if self.config.import_config(file_path):
                    self.load_settings()
                    messagebox.showinfo("Sukses", f"Pengaturan berhasil diimpor dari:\n{file_path}")
                    self.update_status("Pengaturan diimpor")
                else:
                    messagebox.showerror("Error", "Gagal mengimpor pengaturan")

        except Exception as e:
            logger.error(f"Error importing settings: {e}")
            messagebox.showerror("Error", f"Gagal mengimpor pengaturan:\n{str(e)}")

    def update_status(self, message: str):
        """Update status message."""
        if self.status_callback:
            self.status_callback(message)
        logger.info(f"Settings: {message}")
