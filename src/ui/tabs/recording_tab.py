# -*- coding: utf-8 -*-
"""
Recording Tab Module for EchoScribe AI
Integrated from monolithic system with complete recording interface.
"""

import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import time
import logging
from typing import Optional, Callable

from ..device_manager import DeviceManager

logger = logging.getLogger(__name__)

class RecordingTab(ctk.CTkFrame):
    """
    Enhanced recording tab integrated from monolithic system.
    Provides complete recording interface with device selection and controls.
    """

    def __init__(self, master, app_controller, config_manager, status_callback: Optional[Callable] = None):
        super().__init__(master)

        self.app = app_controller
        self.config = config_manager
        self.status_callback = status_callback
        self.device_manager = DeviceManager()

        # Recording state
        self.recording = False
        self.elapsed_time = 0
        self.timer_running = False

        # UI Variables
        self.timer_var = tk.StringVar(value="00:00:00")
        self.progress_var = tk.DoubleVar()

        # Recording mode variables
        self.recording_mode = tk.StringVar(value="microphone")
        self.selected_mic = tk.StringVar(value="0: Default Microphone")
        self.use_extended_recording = tk.BooleanVar(value=True)
        self.chunk_size = tk.IntVar(value=600)

        self.setup_ui()
        self.load_settings()

    def setup_ui(self):
        """Setup the recording tab interface."""

        # Main container with padding
        main_container = ctk.CTkFrame(self)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Title
        title_label = ctk.CTkLabel(
            main_container,
            text="🎤 Perekaman Audio",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 30))

        # Recording mode selection
        self.setup_recording_mode_frame(main_container)

        # Device selection
        self.setup_device_selection_frame(main_container)

        # Recording settings
        self.setup_recording_settings_frame(main_container)

        # Recording controls
        self.setup_recording_controls_frame(main_container)

        # Status and progress
        self.setup_status_frame(main_container)

    def setup_recording_mode_frame(self, parent):
        """Setup recording mode selection frame."""
        mode_frame = ctk.CTkFrame(parent)
        mode_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            mode_frame,
            text="Mode Perekaman:",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 10))

        # Radio buttons for recording modes
        modes_container = ctk.CTkFrame(mode_frame, fg_color="transparent")
        modes_container.pack(fill=tk.X, padx=20, pady=(0, 20))

        modes = [
            ("microphone", "🎤 Mikrofon", "Rekam dari mikrofon"),
            ("system", "🔊 Audio Sistem", "Rekam audio yang sedang diputar komputer"),
            ("dual", "🎵 Gabungan", "Rekam mikrofon dan audio sistem bersamaan")
        ]

        for value, text, tooltip in modes:
            radio = ctk.CTkRadioButton(
                modes_container,
                text=text,
                variable=self.recording_mode,
                value=value,
                command=self.on_mode_change,
                font=ctk.CTkFont(size=14)
            )
            radio.pack(anchor="w", pady=5)

            # Add tooltip (simplified)
            self.add_tooltip(radio, tooltip)

    def setup_device_selection_frame(self, parent):
        """Setup device selection frame."""
        self.device_frame = ctk.CTkFrame(parent)
        self.device_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            self.device_frame,
            text="Pemilihan Perangkat:",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 10))

        # Device selection container
        device_container = ctk.CTkFrame(self.device_frame, fg_color="transparent")
        device_container.pack(fill=tk.X, padx=20, pady=(0, 20))

        # Microphone selection
        mic_frame = ctk.CTkFrame(device_container, fg_color="transparent")
        mic_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(mic_frame, text="Mikrofon:", font=ctk.CTkFont(size=14)).pack(side=tk.LEFT)

        self.mic_dropdown = ctk.CTkComboBox(
            mic_frame,
            values=self.get_microphone_options(),
            variable=self.selected_mic,
            width=300,
            state="readonly"
        )
        self.mic_dropdown.pack(side=tk.LEFT, padx=(10, 0))

        refresh_btn = ctk.CTkButton(
            mic_frame,
            text="🔄 Refresh",
            command=self.refresh_devices,
            width=80,
            height=32
        )
        refresh_btn.pack(side=tk.LEFT, padx=(10, 0))

        # Test device button
        test_btn = ctk.CTkButton(
            mic_frame,
            text="🔍 Test",
            command=self.test_selected_device,
            width=80,
            height=32
        )
        test_btn.pack(side=tk.LEFT, padx=(10, 0))

    def setup_recording_settings_frame(self, parent):
        """Setup recording settings frame."""
        settings_frame = ctk.CTkFrame(parent)
        settings_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            settings_frame,
            text="Pengaturan Perekaman:",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 10))

        settings_container = ctk.CTkFrame(settings_frame, fg_color="transparent")
        settings_container.pack(fill=tk.X, padx=20, pady=(0, 20))

        # Extended recording checkbox
        extended_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        extended_frame.pack(fill=tk.X, pady=(0, 15))

        self.extended_checkbox = ctk.CTkCheckBox(
            extended_frame,
            text="Gunakan Perekaman Diperpanjang",
            variable=self.use_extended_recording,
            command=self.on_extended_change,
            font=ctk.CTkFont(size=14)
        )
        self.extended_checkbox.pack(side=tk.LEFT)

        # Chunk size slider
        chunk_frame = ctk.CTkFrame(settings_container, fg_color="transparent")
        chunk_frame.pack(fill=tk.X, pady=(0, 10))

        self.chunk_slider_frame = ctk.CTkFrame(chunk_frame, fg_color="transparent")
        self.chunk_slider_frame.pack(fill=tk.X)

        ctk.CTkLabel(
            self.chunk_slider_frame,
            text="Durasi Chunk (detik):",
            font=ctk.CTkFont(size=14)
        ).pack(anchor="w", pady=(0, 5))

        self.chunk_slider = ctk.CTkSlider(
            self.chunk_slider_frame,
            from_=30,
            to=1200,
            variable=self.chunk_size,
            number_of_steps=39,
            width=300
        )
        self.chunk_slider.pack(anchor="w", pady=(0, 5))

        self.chunk_value_label = ctk.CTkLabel(
            self.chunk_slider_frame,
            text=f"{self.chunk_size.get()} detik",
            font=ctk.CTkFont(size=12)
        )
        self.chunk_value_label.pack(anchor="w")

        # Update chunk label when slider changes
        self.chunk_size.trace_add("write", self.update_chunk_label)

    def setup_recording_controls_frame(self, parent):
        """Setup recording controls frame."""
        controls_frame = ctk.CTkFrame(parent)
        controls_frame.pack(fill=tk.X, pady=(0, 20))

        # Timer and controls container
        timer_controls_container = ctk.CTkFrame(controls_frame, fg_color="transparent")
        timer_controls_container.pack(expand=True, pady=30)

        # Timer display
        self.timer_label = ctk.CTkLabel(
            timer_controls_container,
            textvariable=self.timer_var,
            font=ctk.CTkFont(size=48, weight="bold"),
            text_color=("gray10", "gray90")
        )
        self.timer_label.pack(pady=(0, 30))

        # Control buttons
        button_frame = ctk.CTkFrame(timer_controls_container, fg_color="transparent")
        button_frame.pack()

        self.record_btn = ctk.CTkButton(
            button_frame,
            text="🔴 Mulai Rekam",
            command=self.toggle_recording,
            font=ctk.CTkFont(size=16, weight="bold"),
            width=200,
            height=50,
            fg_color="red",
            hover_color="darkred"
        )
        self.record_btn.pack(side=tk.LEFT, padx=(0, 20))

        self.stop_btn = ctk.CTkButton(
            button_frame,
            text="⏹️ Stop",
            command=self.stop_recording,
            font=ctk.CTkFont(size=16, weight="bold"),
            width=120,
            height=50,
            state="disabled"
        )
        self.stop_btn.pack(side=tk.LEFT)

    def setup_status_frame(self, parent):
        """Setup status and progress frame."""
        status_frame = ctk.CTkFrame(parent)
        status_frame.pack(fill=tk.X)

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(status_frame, width=400)
        self.progress_bar.pack(pady=(20, 10))
        self.progress_bar.set(0)

        # Status label
        self.status_label = ctk.CTkLabel(
            status_frame,
            text="Siap untuk merekam",
            font=ctk.CTkFont(size=14)
        )
        self.status_label.pack(pady=(0, 20))

    def get_microphone_options(self):
        """Get list of available microphones."""
        try:
            devices = self.device_manager.get_input_devices(include_loopback=False)
            options = []

            for i, device in enumerate(devices):
                options.append(f"{i}: {device['display_name']}")

            if not options:
                options = ["0: Default Microphone"]

            return options
        except Exception as e:
            logger.error(f"Error getting microphone options: {e}")
            return ["0: Default Microphone"]

    def refresh_devices(self):
        """Refresh device list."""
        try:
            self.device_manager.refresh_devices()
            options = self.get_microphone_options()
            self.mic_dropdown.configure(values=options)

            if options and self.selected_mic.get() not in options:
                self.selected_mic.set(options[0])

            self.update_status("Perangkat berhasil diperbarui")
        except Exception as e:
            logger.error(f"Error refreshing devices: {e}")
            self.update_status("Gagal memperbarui perangkat")

    def test_selected_device(self):
        """Test the selected audio device."""
        try:
            # Get selected device info
            mic_selection = self.selected_mic.get()
            if not mic_selection:
                messagebox.showwarning("Peringatan", "Silakan pilih perangkat terlebih dahulu")
                return

            # Extract device index
            device_index = int(mic_selection.split(":")[0])
            devices = self.device_manager.get_input_devices(include_loopback=False)

            if device_index < len(devices):
                device_info = devices[device_index]['device_info']

                self.update_status("Menguji perangkat...")

                # Test device in separate thread
                def test_device():
                    try:
                        success = self.device_manager.test_device(device_info, duration=2.0)

                        if success:
                            self.root.after(0, lambda: messagebox.showinfo(
                                "Test Berhasil",
                                f"Perangkat '{device_info['name']}' dapat digunakan untuk merekam"
                            ))
                            self.root.after(0, lambda: self.update_status("Test perangkat berhasil"))
                        else:
                            self.root.after(0, lambda: messagebox.showerror(
                                "Test Gagal",
                                f"Perangkat '{device_info['name']}' tidak dapat digunakan"
                            ))
                            self.root.after(0, lambda: self.update_status("Test perangkat gagal"))

                    except Exception as e:
                        logger.error(f"Device test error: {e}")
                        self.root.after(0, lambda: messagebox.showerror(
                            "Error",
                            f"Terjadi kesalahan saat menguji perangkat:\n{str(e)}"
                        ))
                        self.root.after(0, lambda: self.update_status("Error saat test perangkat"))

                threading.Thread(target=test_device, daemon=True).start()

        except Exception as e:
            logger.error(f"Error testing device: {e}")
            messagebox.showerror("Error", f"Gagal menguji perangkat:\n{str(e)}")

    def on_mode_change(self):
        """Handle recording mode change."""
        mode = self.recording_mode.get()
        logger.info(f"Recording mode changed to: {mode}")

        # Update UI based on mode
        if mode == "system":
            # Show system audio devices
            self.mic_dropdown.configure(values=self.get_system_audio_options())
        else:
            # Show regular input devices
            self.mic_dropdown.configure(values=self.get_microphone_options())

        self.save_settings()

    def get_system_audio_options(self):
        """Get system audio device options."""
        try:
            devices = self.device_manager.get_system_audio_devices()
            options = []

            for i, device in enumerate(devices):
                options.append(f"{i}: {device['display_name']}")

            if not options:
                options = ["0: Default System Audio"]

            return options
        except Exception as e:
            logger.error(f"Error getting system audio options: {e}")
            return ["0: Default System Audio"]

    def on_extended_change(self):
        """Handle extended recording setting change."""
        extended = self.use_extended_recording.get()
        logger.info(f"Extended recording: {extended}")

        # Enable/disable chunk size slider based on extended recording
        if extended:
            self.chunk_slider.configure(state="normal")
            self.chunk_slider_frame.configure(fg_color=("gray92", "gray14"))
        else:
            self.chunk_slider.configure(state="disabled")
            self.chunk_slider_frame.configure(fg_color=("gray85", "gray25"))

        self.save_settings()

    def update_chunk_label(self, *args):
        """Update chunk size label."""
        value = self.chunk_size.get()
        self.chunk_value_label.configure(text=f"{value} detik")

    def toggle_recording(self):
        """Toggle recording state."""
        if not self.recording:
            self.start_recording()
        else:
            self.stop_recording()

    def start_recording(self):
        """Start recording."""
        try:
            # Get recording configuration
            config = {
                'mode': self.recording_mode.get(),
                'device': self.selected_mic.get(),
                'extended': self.use_extended_recording.get(),
                'chunk_size': self.chunk_size.get()
            }

            # Start recording via app controller
            if self.app.start_recording(config):
                self.recording = True
                self.start_timer()

                # Update UI
                self.record_btn.configure(
                    text="⏸️ Pause",
                    fg_color="orange",
                    hover_color="darkorange"
                )
                self.stop_btn.configure(state="normal")

                self.update_status("Merekam...")
                logger.info("Recording started")
            else:
                self.update_status("Gagal memulai perekaman")

        except Exception as e:
            logger.error(f"Error starting recording: {e}")
            self.update_status(f"Error: {str(e)}")

    def stop_recording(self):
        """Stop recording."""
        try:
            if self.app.stop_recording():
                self.recording = False
                self.stop_timer()

                # Update UI
                self.record_btn.configure(
                    text="🔴 Mulai Rekam",
                    fg_color="red",
                    hover_color="darkred"
                )
                self.stop_btn.configure(state="disabled")

                self.update_status("Perekaman dihentikan")
                logger.info("Recording stopped")
            else:
                self.update_status("Gagal menghentikan perekaman")

        except Exception as e:
            logger.error(f"Error stopping recording: {e}")
            self.update_status(f"Error: {str(e)}")

    def start_timer(self):
        """Start recording timer."""
        self.timer_running = True
        self.elapsed_time = 0
        self.update_timer()

    def stop_timer(self):
        """Stop recording timer."""
        self.timer_running = False

    def update_timer(self):
        """Update timer display."""
        if self.timer_running:
            hours = self.elapsed_time // 3600
            minutes = (self.elapsed_time % 3600) // 60
            seconds = self.elapsed_time % 60

            time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            self.timer_var.set(time_str)

            # Update progress bar based on chunk size
            if self.use_extended_recording.get():
                chunk_progress = (self.elapsed_time % self.chunk_size.get()) / self.chunk_size.get()
                self.progress_bar.set(chunk_progress)

            self.elapsed_time += 1
            self.after(1000, self.update_timer)

    def update_status(self, message: str):
        """Update status message."""
        self.status_label.configure(text=message)
        if self.status_callback:
            self.status_callback(message)

    def load_settings(self):
        """Load recording settings from config."""
        try:
            config = self.config.get_recording_config()

            self.recording_mode.set(config.get('recording_mode', 'microphone'))
            self.selected_mic.set(config.get('selected_mic', '0: Default Microphone'))
            self.use_extended_recording.set(config.get('use_extended_recording', True))
            self.chunk_size.set(config.get('chunk_size', 600))

            # Update UI based on loaded settings
            self.on_mode_change()
            self.on_extended_change()

        except Exception as e:
            logger.error(f"Error loading recording settings: {e}")

    def save_settings(self):
        """Save current recording settings."""
        try:
            settings = {
                'recording_mode': self.recording_mode.get(),
                'selected_mic': self.selected_mic.get(),
                'use_extended_recording': self.use_extended_recording.get(),
                'chunk_size': self.chunk_size.get()
            }

            self.config.update(settings)

        except Exception as e:
            logger.error(f"Error saving recording settings: {e}")

    def add_tooltip(self, widget, text):
        """Add simple tooltip to widget."""
        def show_tooltip(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")

            label = ctk.CTkLabel(
                tooltip,
                text=text,
                fg_color=("gray75", "gray25"),
                corner_radius=6,
                padx=8,
                pady=4
            )
            label.pack()

            def hide_tooltip():
                tooltip.destroy()

            tooltip.after(3000, hide_tooltip)  # Auto-hide after 3 seconds

        widget.bind("<Enter>", show_tooltip)
