# -*- coding: utf-8 -*-
"""
Enhanced User Interface Module for EchoScribe AI
Fully integrated from monolithic system with complete functionality.
Modern, responsive UI using customtkinter with tabbed interface.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import customtkinter as ctk
import threading
import time
from pathlib import Path
from typing import Optional, Dict, Any, Callable
import sys
import os
import logging
import json

# Import our app controller
from src.app_controller import EchoScribeApp

logger = logging.getLogger(__name__)

class ValueTrackingSlider(ctk.CTkFrame):
    """Enhanced slider with value tracking from monolithic system."""

    def __init__(self, master, title, from_, to, variable, format_string="{:.0f}", width=300, **kwargs):
        super().__init__(master, **kwargs)

        self.format_string = format_string
        self.variable = variable
        self.title = title
        self.tooltip_window = None

        self.title_label = ctk.CTkLabel(self, text=title)
        self.title_label.grid(row=0, column=0, sticky="w", padx=(0,10))

        self.value_text = ctk.StringVar()
        self.update_value_text()
        self.value_label = ctk.CTkLabel(self, textvariable=self.value_text, width=60)
        self.value_label.grid(row=0, column=1, sticky="e")

        self.slider = ctk.CTkSlider(self, from_=from_, to=to, variable=variable, width=width,
                                   command=self.on_slider_change)
        self.slider.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(5, 0))

        self.tooltip_text = f"{title}: {self.format_string.format(variable.get())}"

        # Tooltip events
        self.slider.bind("<Enter>", self.show_tooltip)
        self.slider.bind("<Leave>", self.hide_tooltip)
        self.slider.bind("<Motion>", self.update_tooltip_position)

        variable.trace_add("write", self.update_value_text)

    def update_value_text(self, *args):
        """Update displayed value text."""
        value = self.variable.get()
        self.value_text.set(self.format_string.format(value))
        self.tooltip_text = f"{self.title}: {self.format_string.format(value)}"
        if hasattr(self, 'tooltip_window') and self.tooltip_window:
            self.tooltip_window.winfo_children()[0].configure(text=self.tooltip_text)

    def on_slider_change(self, value):
        """Handle slider value change."""
        if hasattr(self, 'variable'):
            self.variable.set(value)
        self.update_value_text()
        logger.debug(f"Slider {self.title} changed to: {value}")

    def hide_tooltip(self, event):
        """Hide tooltip when mouse leaves."""
        if hasattr(self, 'tooltip_window') and self.tooltip_window:
            if event.widget == self.slider:
                self.tooltip_window.destroy()
                self.tooltip_window = None

    def show_tooltip(self, event):
        """Show tooltip on mouse enter."""
        x, y = event.x_root, event.y_root
        self.tooltip_window = tk.Toplevel(self)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x+10}+{y+10}")

        label = ctk.CTkLabel(
            self.tooltip_window,
            text=self.tooltip_text,
            fg_color=("gray75", "gray25"),
            corner_radius=6,
            padx=8,
            pady=4
        )
        label.pack()

    def update_tooltip_position(self, event):
        """Update tooltip position on mouse move."""
        if hasattr(self, 'tooltip_window') and self.tooltip_window:
            x, y = event.x_root, event.y_root
            self.tooltip_window.wm_geometry(f"+{x+10}+{y+10}")

class APIKeyDialog(ctk.CTkToplevel):
    """Enhanced API Key Dialog integrated from monolithic system."""

    def __init__(self, parent, config_manager, current_key=""):
        super().__init__(parent)

        self.title("Pengaturan API Key Groq")
        self.geometry("500x450")
        self.transient(parent)
        self.grab_set()

        self.config_manager = config_manager
        self.api_key = None
        self.result = None

        # Center dialog
        self.geometry(f"+{parent.winfo_rootx() + 50}+{parent.winfo_rooty() + 50}")
        self.setup_ui(current_key)

    def setup_ui(self, current_key):
        """Setup enhanced UI for API key dialog."""
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Title
        title_label = ctk.CTkLabel(
            main_frame,
            text="🔑 Konfigurasi API Key Groq",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(pady=(0, 20))

        # Information
        info_text = """EchoScribe AI menggunakan API Groq untuk transkripsi dan peningkatan teks.
Anda perlu menyediakan API key Groq sendiri untuk menggunakan aplikasi ini.
API key akan disimpan secara aman di komputer lokal Anda."""

        info_label = ctk.CTkLabel(
            main_frame,
            text=info_text,
            wraplength=450,
            justify="left",
            font=ctk.CTkFont(size=12)
        )
        info_label.pack(pady=(0, 20))

        # Current key status
        status_frame = ctk.CTkFrame(main_frame)
        status_frame.pack(fill=tk.X, pady=(0, 20))

        if current_key:
            status_text = f"✅ API Key Tersimpan: {current_key[:15]}...{current_key[-5:]}"
            status_color = "green"
        else:
            status_text = "❌ Tidak Ada API Key Tersimpan"
            status_color = "red"

        status_label = ctk.CTkLabel(
            status_frame,
            text=status_text,
            text_color=status_color,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        status_label.pack(pady=10)

        # Input section
        input_frame = ctk.CTkFrame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            input_frame,
            text="Masukkan API Key Groq Anda:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(anchor="w", padx=10, pady=(15, 5))

        self.api_key_entry = ctk.CTkEntry(
            input_frame,
            placeholder_text="gsk_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
            width=450,
            height=35,
            show="*"
        )
        self.api_key_entry.pack(padx=10, pady=(0, 10))

        if current_key:
            self.api_key_entry.insert(0, current_key)

        # Show/Hide checkbox
        show_frame = ctk.CTkFrame(input_frame, fg_color="transparent")
        show_frame.pack(padx=10, pady=(0, 15))

        self.show_key = tk.BooleanVar()
        show_checkbox = ctk.CTkCheckBox(
            show_frame,
            text="Tampilkan API Key",
            variable=self.show_key,
            command=self.toggle_key_visibility
        )
        show_checkbox.pack(side=tk.LEFT)

        # Validation button
        validate_btn = ctk.CTkButton(
            show_frame,
            text="🔍 Validasi",
            command=self.validate_key_only,
            width=100,
            height=28
        )
        validate_btn.pack(side=tk.RIGHT)

        # Instructions
        instructions_frame = ctk.CTkFrame(main_frame)
        instructions_frame.pack(fill=tk.X, pady=(0, 20))

        instructions_text = """📋 Cara mendapatkan API Key Groq:

1. Buka https://console.groq.com/ di browser
2. Daftar akun baru atau login ke akun yang sudah ada
3. Navigasi ke bagian "API Keys" di dashboard
4. Klik "Create API Key" untuk membuat key baru
5. Salin API key yang dihasilkan dan tempel di sini
6. Pastikan API key dimulai dengan "gsk_"

💡 Tips: Simpan API key di tempat aman untuk penggunaan masa depan"""

        instructions_label = ctk.CTkLabel(
            instructions_frame,
            text=instructions_text,
            justify="left",
            font=ctk.CTkFont(size=11),
            wraplength=450
        )
        instructions_label.pack(padx=15, pady=15)

        # Button frame
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill=tk.X, pady=(0, 10))

        # Remove button (if key exists)
        if current_key:
            remove_btn = ctk.CTkButton(
                button_frame,
                text="🗑️ Hapus Key",
                command=self.remove_key,
                fg_color="red",
                hover_color="darkred",
                width=120,
                height=35
            )
            remove_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Cancel button
        cancel_btn = ctk.CTkButton(
            button_frame,
            text="❌ Batal",
            command=self.cancel,
            fg_color="gray",
            hover_color="darkgray",
            width=100,
            height=35
        )
        cancel_btn.pack(side=tk.RIGHT)

        # Save button
        save_btn = ctk.CTkButton(
            button_frame,
            text="💾 Simpan",
            command=self.save_key,
            width=120,
            height=35
        )
        save_btn.pack(side=tk.RIGHT, padx=(10, 0))

    def toggle_key_visibility(self):
        """Toggle API key visibility."""
        if self.show_key.get():
            self.api_key_entry.configure(show="")
        else:
            self.api_key_entry.configure(show="*")

    def validate_key_only(self):
        """Validate API key without saving."""
        key = self.api_key_entry.get().strip()

        if not key:
            messagebox.showwarning("Peringatan", "Silakan masukkan API key terlebih dahulu")
            return

        if not self.config_manager.validate_api_key(key):
            messagebox.showwarning("Peringatan", "Format API key tidak valid")
            return

        # Test key with Groq API
        try:
            import groq
            test_client = groq.Groq(api_key=key)

            # Simple test request
            response = test_client.chat.completions.create(
                model="deepseek-r1-distill-llama-70b",
                messages=[{"role": "user", "content": "test"}],
                max_tokens=5
            )

            messagebox.showinfo("✅ Validasi Berhasil", "API key valid dan dapat digunakan!")

        except Exception as e:
            error_msg = str(e)
            if "401" in error_msg or "authentication" in error_msg.lower():
                messagebox.showerror("❌ Validasi Gagal", "API key tidak valid atau tidak memiliki akses")
            elif "429" in error_msg or "rate limit" in error_msg.lower():
                messagebox.showinfo("⚠️ Rate Limit", "API key valid, tetapi rate limit tercapai")
            else:
                messagebox.showerror("❌ Error", f"Gagal memvalidasi API key:\n{error_msg[:200]}")

    def save_key(self):
        """Save API key after validation."""
        key = self.api_key_entry.get().strip()

        if not key:
            messagebox.showwarning("Peringatan", "Silakan masukkan API key")
            return

        if not self.config_manager.validate_api_key(key):
            messagebox.showwarning("Peringatan", "Format API key tidak valid")
            return

        # Save the key
        if self.config_manager.save_user_api_key(key):
            self.api_key = key
            self.result = "saved"
            messagebox.showinfo("✅ Sukses", "API key berhasil disimpan!")
            self.destroy()
        else:
            messagebox.showerror("❌ Error", "Gagal menyimpan API key")

    def remove_key(self):
        """Remove stored API key."""
        if messagebox.askyesno("Konfirmasi", "Apakah Anda yakin ingin menghapus API key yang tersimpan?"):
            if self.config_manager.remove_user_api_key():
                self.result = "removed"
                messagebox.showinfo("✅ Sukses", "API key berhasil dihapus")
                self.destroy()
            else:
                messagebox.showerror("❌ Error", "Gagal menghapus API key")

    def cancel(self):
        """Cancel dialog."""
        self.result = "cancel"
        self.destroy()

class EchoScribeUI:
    """
    Enhanced UI for EchoScribe AI with complete integration from monolithic system.
    Features tabbed interface, real-time visualization, and modern design.
    """

    def __init__(self):
        # Set CustomTkinter appearance before creating widgets
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
        self.root.title("EchoScribe AI - Smart Voice Note Creator")
        self.root.geometry("1200x900")
        self.root.minsize(1000, 700)

        # Set icon
        self._set_window_icon()

        # Initialize the app controller
        self.app = EchoScribeApp(self._update_status)

        # UI Variables from monolithic integration
        self.status_var = tk.StringVar(value="Siap untuk merekam")
        self.timer_var = tk.StringVar(value="00:00:00")
        self.progress_var = tk.DoubleVar()

        # Recording state
        self.elapsed_time = 0
        self.timer_running = False
        self.recording = False

        # Visualization variables from monolithic
        self.viz_data = []
        self.viz_after_id = None

        # UI Components
        self.setup_ui()
        self.setup_keybindings()

        logger.info("EchoScribe UI initialized successfully")

    def _set_window_icon(self):
        """Set application window icon."""
        try:
            # Try to set icon from various possible locations
            icon_paths = [
                Path(__file__).parent.parent.parent / "icon.ico",
                Path("icon.ico"),
                Path("assets/icon.ico")
            ]

            for icon_path in icon_paths:
                if icon_path.exists():
                    self.root.iconbitmap(str(icon_path))
                    logger.info(f"Icon set from: {icon_path}")
                    return

            logger.warning("No icon file found")

        except Exception as e:
            logger.warning(f"Could not set window icon: {e}")

    def setup_ui(self):
        """Setup the main user interface with tabbed layout."""

        # Configure root window
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Main container
        main_frame = ctk.CTkFrame(self.root)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # Create tabbed interface
        self.create_tabbed_interface(main_frame)

        # Status bar at bottom
        self.create_status_bar(main_frame)

    def create_tabbed_interface(self, parent):
        """Create the main tabbed interface."""

        # Tab view
        self.tab_view = ctk.CTkTabview(parent, width=1150, height=700)
        self.tab_view.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 5))

        # Recording tab
        self.tab_view.add("🎤 Perekaman")
        recording_frame = self.tab_view.tab("🎤 Perekaman")

        from .tabs.recording_tab import RecordingTab
        self.recording_tab = RecordingTab(
            recording_frame,
            self.app,
            self.app.config,
            self._update_status
        )
        self.recording_tab.pack(fill=tk.BOTH, expand=True)

        # Processing tab
        self.tab_view.add("⚙️ Pemrosesan")
        processing_frame = self.tab_view.tab("⚙️ Pemrosesan")
        self.create_processing_tab(processing_frame)

        # Results tab
        self.tab_view.add("📄 Hasil")
        results_frame = self.tab_view.tab("📄 Hasil")
        self.create_results_tab(results_frame)

        # Visualization tab
        self.tab_view.add("📊 Visualisasi")
        viz_frame = self.tab_view.tab("📊 Visualisasi")
        self.create_visualization_tab(viz_frame)

        # Settings tab
        self.tab_view.add("⚙️ Pengaturan")
        settings_frame = self.tab_view.tab("⚙️ Pengaturan")

        from .tabs.settings_tab import SettingsTab
        self.settings_tab = SettingsTab(
            settings_frame,
            self.app.config,
            self._update_status
        )
        self.settings_tab.pack(fill=tk.BOTH, expand=True)

        # Help tab
        self.tab_view.add("❓ Bantuan")
        help_frame = self.tab_view.tab("❓ Bantuan")
        self.create_help_tab(help_frame)

    def create_processing_tab(self, parent):
        """Create processing tab interface."""

        # Main container
        container = ctk.CTkFrame(parent)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Title
        title_label = ctk.CTkLabel(
            container,
            text="⚙️ Pemrosesan Audio",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 30))

        # Processing status
        status_frame = ctk.CTkFrame(container)
        status_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            status_frame,
            text="Status Pemrosesan:",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 10))

        self.processing_status_label = ctk.CTkLabel(
            status_frame,
            text="Siap untuk memproses",
            font=ctk.CTkFont(size=14)
        )
        self.processing_status_label.pack(anchor="w", padx=20, pady=(0, 20))

        # Progress bar
        self.processing_progress = ctk.CTkProgressBar(status_frame, width=500)
        self.processing_progress.pack(padx=20, pady=(0, 20))
        self.processing_progress.set(0)

        # Processing controls
        controls_frame = ctk.CTkFrame(container)
        controls_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            controls_frame,
            text="Kontrol Pemrosesan:",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 15))

        button_frame = ctk.CTkFrame(controls_frame, fg_color="transparent")
        button_frame.pack(padx=20, pady=(0, 20))

        self.process_btn = ctk.CTkButton(
            button_frame,
            text="🚀 Mulai Pemrosesan",
            command=self.start_processing,
            font=ctk.CTkFont(size=14, weight="bold"),
            width=200,
            height=40
        )
        self.process_btn.pack(side=tk.LEFT, padx=(0, 20))

        self.stop_process_btn = ctk.CTkButton(
            button_frame,
            text="⏹️ Stop Pemrosesan",
            command=self.stop_processing,
            font=ctk.CTkFont(size=14, weight="bold"),
            width=200,
            height=40,
            state="disabled"
        )
        self.stop_process_btn.pack(side=tk.LEFT)

        # Processing log
        log_frame = ctk.CTkFrame(container)
        log_frame.pack(fill=tk.BOTH, expand=True)

        ctk.CTkLabel(
            log_frame,
            text="Log Pemrosesan:",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 10))

        self.processing_log = ctk.CTkTextbox(log_frame, height=200)
        self.processing_log.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

    def create_results_tab(self, parent):
        """Create results tab interface."""

        # Main container
        container = ctk.CTkFrame(parent)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Title
        title_label = ctk.CTkLabel(
            container,
            text="📄 Hasil Transkripsi",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 30))

        # Results controls
        controls_frame = ctk.CTkFrame(container)
        controls_frame.pack(fill=tk.X, pady=(0, 20))

        button_frame = ctk.CTkFrame(controls_frame, fg_color="transparent")
        button_frame.pack(pady=15)

        open_btn = ctk.CTkButton(
            button_frame,
            text="📂 Buka File",
            command=self.open_result_file,
            width=120
        )
        open_btn.pack(side=tk.LEFT, padx=(0, 10))

        save_btn = ctk.CTkButton(
            button_frame,
            text="💾 Simpan",
            command=self.save_result,
            width=120
        )
        save_btn.pack(side=tk.LEFT, padx=(0, 10))

        copy_btn = ctk.CTkButton(
            button_frame,
            text="📋 Salin",
            command=self.copy_result,
            width=120
        )
        copy_btn.pack(side=tk.LEFT, padx=(0, 10))

        clear_btn = ctk.CTkButton(
            button_frame,
            text="🗑️ Bersihkan",
            command=self.clear_result,
            width=120
        )
        clear_btn.pack(side=tk.LEFT)

        # Results display
        results_frame = ctk.CTkFrame(container)
        results_frame.pack(fill=tk.BOTH, expand=True)

        ctk.CTkLabel(
            results_frame,
            text="Hasil Transkripsi:",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 10))

        self.results_text = ctk.CTkTextbox(results_frame, height=400)
        self.results_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

    def create_visualization_tab(self, parent):
        """Create visualization tab interface."""

        # Main container
        container = ctk.CTkFrame(parent)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Title
        title_label = ctk.CTkLabel(
            container,
            text="📊 Visualisasi Audio",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 30))

        # Visualization controls
        controls_frame = ctk.CTkFrame(container)
        controls_frame.pack(fill=tk.X, pady=(0, 20))

        ctk.CTkLabel(
            controls_frame,
            text="Kontrol Visualisasi:",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=20, pady=(20, 15))

        viz_controls = ctk.CTkFrame(controls_frame, fg_color="transparent")
        viz_controls.pack(fill=tk.X, padx=20, pady=(0, 20))

        # Mode selection
        mode_frame = ctk.CTkFrame(viz_controls, fg_color="transparent")
        mode_frame.pack(anchor="w", pady=(0, 10))

        ctk.CTkLabel(mode_frame, text="Mode:", width=100).pack(side=tk.LEFT)

        self.viz_mode_var = tk.StringVar(value="waveform")
        viz_mode_combo = ctk.CTkComboBox(
            mode_frame,
            values=["waveform", "bars", "spectrum", "fill"],
            variable=self.viz_mode_var,
            width=150,
            state="readonly"
        )
        viz_mode_combo.pack(side=tk.LEFT, padx=(10, 0))

        # Enable/disable toggle
        enable_frame = ctk.CTkFrame(viz_controls, fg_color="transparent")
        enable_frame.pack(anchor="w")

        self.viz_enabled_var = tk.BooleanVar(value=True)
        viz_enable_check = ctk.CTkCheckBox(
            enable_frame,
            text="Aktifkan Visualisasi Real-time",
            variable=self.viz_enabled_var,
            command=self.toggle_visualization
        )
        viz_enable_check.pack(side=tk.LEFT)

        # Visualization display area
        self.viz_frame = ctk.CTkFrame(container)
        self.viz_frame.pack(fill=tk.BOTH, expand=True)

        # Placeholder for visualization
        self.viz_placeholder = ctk.CTkLabel(
            self.viz_frame,
            text="📊 Visualisasi audio akan ditampilkan di sini saat perekaman aktif",
            font=ctk.CTkFont(size=16),
            text_color="gray"
        )
        self.viz_placeholder.pack(expand=True)

    def create_help_tab(self, parent):
        """Create help tab interface."""

        # Main scrollable container
        container = ctk.CTkScrollableFrame(parent)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Title
        title_label = ctk.CTkLabel(
            container,
            text="❓ Bantuan & Dokumentasi",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 30))

        # Quick start section
        self.create_help_section(
            container,
            "🚀 Memulai Cepat",
            """1. Pastikan mikrofon terhubung dan berfungsi
2. Pilih mode perekaman di tab Perekaman
3. Klik tombol 'Mulai Rekam' untuk memulai
4. Berbicara dengan jelas ke mikrofon
5. Klik 'Stop' untuk menghentikan perekaman
6. Lihat hasil transkripsi di tab Hasil"""
        )

        # Features section
        self.create_help_section(
            container,
            "✨ Fitur Utama",
            """• Perekaman audio dari mikrofon atau sistem
• Transkripsi otomatis menggunakan AI Groq
• Peningkatan teks dengan AI
• Visualisasi audio real-time
• Export ke berbagai format dokumen
• Pengaturan audio yang dapat disesuaikan
• Antarmuka modern dan responsif"""
        )

        # Recording modes section
        self.create_help_section(
            container,
            "🎤 Mode Perekaman",
            """• Mikrofon: Rekam dari mikrofon yang dipilih
• Audio Sistem: Rekam audio yang sedang diputar komputer
• Gabungan: Rekam mikrofon dan audio sistem bersamaan

Gunakan mode yang sesuai dengan kebutuhan Anda."""
        )

        # Tips section
        self.create_help_section(
            container,
            "💡 Tips Penggunaan",
            """• Gunakan mikrofon berkualitas baik untuk hasil optimal
• Hindari kebisingan latar belakang
• Berbicara dengan kecepatan normal dan jelas
• Gunakan perekaman diperpanjang untuk sesi yang lama
• Sesuaikan pengaturan audio sesuai lingkungan
• Pastikan koneksi internet stabil untuk AI processing"""
        )

        # Troubleshooting section
        self.create_help_section(
            container,
            "🔧 Troubleshooting",
            """• Jika mikrofon tidak terdeteksi, coba refresh device list
• Jika hasil transkripsi tidak akurat, periksa kualitas audio
• Jika AI processing lambat, periksa koneksi internet
• Jika aplikasi crash, restart dan coba lagi
• Untuk error API, periksa API key Groq Anda"""
        )

        # Contact section
        self.create_help_section(
            container,
            "📞 Dukungan",
            """Jika Anda mengalami masalah atau memiliki pertanyaan:

• Periksa log error di console
• Restart aplikasi
• Periksa pengaturan sistem audio
• Pastikan API key Groq valid dan aktif

EchoScribe AI - Smart Voice Note Creator
Versi 2.0 Enhanced"""
        )

    def create_help_section(self, parent, title, content):
        """Create a help section with title and content."""
        section_frame = ctk.CTkFrame(parent)
        section_frame.pack(fill=tk.X, pady=(0, 20))

        title_label = ctk.CTkLabel(
            section_frame,
            text=title,
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(anchor="w", padx=20, pady=(20, 10))

        content_label = ctk.CTkLabel(
            section_frame,
            text=content,
            justify="left",
            wraplength=700,
            font=ctk.CTkFont(size=12)
        )
        content_label.pack(anchor="w", padx=20, pady=(0, 20))

    def create_status_bar(self, parent):
        """Create status bar at bottom."""
        status_frame = ctk.CTkFrame(parent)
        status_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(5, 10))

        # Status label
        self.status_label = ctk.CTkLabel(
            status_frame,
            textvariable=self.status_var,
            font=ctk.CTkFont(size=12)
        )
        self.status_label.pack(side=tk.LEFT, padx=20, pady=10)

        # Timer in status bar
        self.timer_status_label = ctk.CTkLabel(
            status_frame,
            textvariable=self.timer_var,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.timer_status_label.pack(side=tk.RIGHT, padx=20, pady=10)

    def setup_keybindings(self):
        """Setup keyboard shortcuts."""
        # Recording shortcuts
        self.root.bind('<F1>', lambda e: self.toggle_recording())
        self.root.bind('<F2>', lambda e: self.stop_recording())
        self.root.bind('<Escape>', lambda e: self.stop_recording())

        # Tab shortcuts
        self.root.bind('<Control-1>', lambda e: self.tab_view.set("🎤 Perekaman"))
        self.root.bind('<Control-2>', lambda e: self.tab_view.set("⚙️ Pemrosesan"))
        self.root.bind('<Control-3>', lambda e: self.tab_view.set("📄 Hasil"))
        self.root.bind('<Control-4>', lambda e: self.tab_view.set("📊 Visualisasi"))
        self.root.bind('<Control-5>', lambda e: self.tab_view.set("⚙️ Pengaturan"))

        # General shortcuts
        self.root.bind('<Control-s>', lambda e: self.save_result())
        self.root.bind('<Control-o>', lambda e: self.open_result_file())
        self.root.bind('<Control-q>', lambda e: self.on_closing())

    # Event handlers and utility methods
    def toggle_recording(self):
        """Toggle recording state."""
        if hasattr(self, 'recording_tab'):
            self.recording_tab.toggle_recording()

    def stop_recording(self):
        """Stop recording."""
        if hasattr(self, 'recording_tab'):
            self.recording_tab.stop_recording()

    def start_processing(self):
        """Start audio processing."""
        try:
            self.process_btn.configure(state="disabled")
            self.stop_process_btn.configure(state="normal")

            # Start processing in app controller
            if self.app.start_processing():
                self.processing_status_label.configure(text="Memproses audio...")
                self.add_processing_log("Pemrosesan dimulai...")
            else:
                self.processing_status_label.configure(text="Gagal memulai pemrosesan")
                self.process_btn.configure(state="normal")
                self.stop_process_btn.configure(state="disabled")

        except Exception as e:
            logger.error(f"Error starting processing: {e}")
            self.processing_status_label.configure(text=f"Error: {str(e)}")

    def stop_processing(self):
        """Stop audio processing."""
        try:
            if self.app.stop_processing():
                self.processing_status_label.configure(text="Pemrosesan dihentikan")
                self.add_processing_log("Pemrosesan dihentikan oleh pengguna")

            self.process_btn.configure(state="normal")
            self.stop_process_btn.configure(state="disabled")

        except Exception as e:
            logger.error(f"Error stopping processing: {e}")

    def add_processing_log(self, message):
        """Add message to processing log."""
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.processing_log.insert("end", log_entry)
        self.processing_log.see("end")

    def open_result_file(self):
        """Open result file."""
        try:
            file_path = filedialog.askopenfilename(
                title="Buka File Hasil",
                filetypes=[
                    ("Text files", "*.txt"),
                    ("Word documents", "*.docx"),
                    ("All files", "*.*")
                ]
            )

            if file_path:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    self.results_text.delete("1.0", "end")
                    self.results_text.insert("1.0", content)

                self._update_status(f"File dibuka: {Path(file_path).name}")

        except Exception as e:
            logger.error(f"Error opening file: {e}")
            messagebox.showerror("Error", f"Gagal membuka file:\n{str(e)}")

    def save_result(self):
        """Save current result."""
        try:
            content = self.results_text.get("1.0", "end-1c")
            if not content.strip():
                messagebox.showwarning("Peringatan", "Tidak ada konten untuk disimpan")
                return

            file_path = filedialog.asksaveasfilename(
                title="Simpan Hasil",
                defaultextension=".txt",
                filetypes=[
                    ("Text files", "*.txt"),
                    ("Word documents", "*.docx"),
                    ("All files", "*.*")
                ]
            )

            if file_path:
                if file_path.endswith('.docx'):
                    # Save as Word document
                    self.app.save_as_docx(content, file_path)
                else:
                    # Save as text file
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(content)

                self._update_status(f"File disimpan: {Path(file_path).name}")
                messagebox.showinfo("Sukses", f"File berhasil disimpan:\n{file_path}")

        except Exception as e:
            logger.error(f"Error saving file: {e}")
            messagebox.showerror("Error", f"Gagal menyimpan file:\n{str(e)}")

    def copy_result(self):
        """Copy result to clipboard."""
        try:
            content = self.results_text.get("1.0", "end-1c")
            if content.strip():
                self.root.clipboard_clear()
                self.root.clipboard_append(content)
                self._update_status("Hasil disalin ke clipboard")
                messagebox.showinfo("Sukses", "Hasil berhasil disalin ke clipboard")
            else:
                messagebox.showwarning("Peringatan", "Tidak ada konten untuk disalin")

        except Exception as e:
            logger.error(f"Error copying to clipboard: {e}")
            messagebox.showerror("Error", f"Gagal menyalin ke clipboard:\n{str(e)}")

    def clear_result(self):
        """Clear result text."""
        if messagebox.askyesno("Konfirmasi", "Apakah Anda yakin ingin menghapus semua hasil?"):
            self.results_text.delete("1.0", "end")
            self._update_status("Hasil dihapus")

    def toggle_visualization(self):
        """Toggle visualization on/off."""
        enabled = self.viz_enabled_var.get()
        if enabled:
            self.start_visualization()
        else:
            self.stop_visualization()

    def start_visualization(self):
        """Start audio visualization."""
        try:
            if hasattr(self.app, 'visualizer'):
                self.app.visualizer.start_visualization(self.viz_frame)
                self.viz_placeholder.pack_forget()
                self._update_status("Visualisasi diaktifkan")
        except Exception as e:
            logger.error(f"Error starting visualization: {e}")

    def stop_visualization(self):
        """Stop audio visualization."""
        try:
            if hasattr(self.app, 'visualizer'):
                self.app.visualizer.stop_visualization()
                self.viz_placeholder.pack(expand=True)
                self._update_status("Visualisasi dinonaktifkan")
        except Exception as e:
            logger.error(f"Error stopping visualization: {e}")

    def _update_status(self, message: str):
        """Update status message."""
        self.status_var.set(message)
        logger.info(f"Status: {message}")

    def update_results(self, content: str):
        """Update results display with new content."""
        self.results_text.delete("1.0", "end")
        self.results_text.insert("1.0", content)

        # Switch to results tab
        self.tab_view.set("📄 Hasil")

    def update_processing_progress(self, progress: float, message: str = ""):
        """Update processing progress."""
        self.processing_progress.set(progress)
        if message:
            self.processing_status_label.configure(text=message)
            self.add_processing_log(message)

    def on_closing(self):
        """Handle window closing event."""
        try:
            # Stop any ongoing recording or processing
            if self.recording:
                self.stop_recording()

            # Save current settings
            if hasattr(self, 'settings_tab'):
                self.settings_tab.save_settings()

            # Cleanup resources
            if hasattr(self.app, 'cleanup'):
                self.app.cleanup()

            self.root.quit()
            self.root.destroy()

        except Exception as e:
            logger.error(f"Error during shutdown: {e}")
            self.root.quit()

    def run(self):
        """Run the application."""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Check API key on startup
        if not self.app.config.has_user_api_key():
            self.root.after(1000, self.show_api_key_setup)

        # Start the main loop
        self.root.mainloop()

    def show_api_key_setup(self):
        """Show API key setup dialog on first run."""
        messagebox.showinfo(
            "Selamat Datang!",
            "Selamat datang di EchoScribe AI!\n\n" +
            "Untuk menggunakan aplikasi ini, Anda perlu mengatur API key Groq terlebih dahulu.\n\n" +
            "Silakan buka tab Pengaturan untuk mengatur API key Anda."
        )

        # Switch to settings tab
        self.tab_view.set("⚙️ Pengaturan")

        # Load initial configuration
        self._load_ui_config()

        # Bind cleanup
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _set_window_icon(self):
        """Set window icon if available."""
        try:
            base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(base_path, "..", "..", "icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Could not set icon: {e}")

    def setup_ui(self):
        """Setup the main UI layout."""
        # Main container
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Create sections
        self._create_header_section()
        self._create_control_section()
        self._create_settings_section()
        self._create_output_section()
        self._create_status_section()

    def _create_header_section(self):
        """Create the header section with title and API key setup."""
        header_frame = ctk.CTkFrame(self.main_frame)
        header_frame.pack(fill="x", padx=10, pady=(10, 5))

        # Title
        title_label = ctk.CTkLabel(
            header_frame,
            text="🎙️ EchoScribe AI - Smart Voice Note Creator",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=15)

        # API Key section
        api_frame = ctk.CTkFrame(header_frame)
        api_frame.pack(fill="x", padx=20, pady=(0, 15))

        # API Key status and button
        self.api_status_label = ctk.CTkLabel(api_frame, text="API Key: Not configured")
        self.api_status_label.pack(side="left", padx=10, pady=10)

        self.api_button = ctk.CTkButton(
            api_frame,
            text="Configure API Key",
            command=self._show_api_dialog,
            width=150
        )
        self.api_button.pack(side="right", padx=10, pady=10)

        self._update_api_status()

    def _create_control_section(self):
        """Create recording control section."""
        control_frame = ctk.CTkFrame(self.main_frame)
        control_frame.pack(fill="x", padx=10, pady=5)

        # Recording controls
        recording_frame = ctk.CTkFrame(control_frame)
        recording_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(recording_frame, text="Recording Controls", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=5)

        # Recording mode selection
        mode_frame = ctk.CTkFrame(recording_frame)
        mode_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(mode_frame, text="Recording Mode:").pack(side="left", padx=5)

        self.recording_mode = ctk.CTkOptionMenu(
            mode_frame,
            values=["microphone", "system", "dual"],
            command=self._on_mode_change
        )
        self.recording_mode.pack(side="left", padx=5)

        # Record button and timer
        button_frame = ctk.CTkFrame(recording_frame)
        button_frame.pack(fill="x", padx=10, pady=10)

        self.record_button = ctk.CTkButton(
            button_frame,
            text="🎙️ Start Recording",
            command=self._toggle_recording,
            width=200,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.record_button.pack(side="left", padx=5)

        self.timer_label = ctk.CTkLabel(
            button_frame,
            textvariable=self.timer_var,
            font=ctk.CTkFont(size=16, family="monospace")
        )
        self.timer_label.pack(side="left", padx=20)

        # Process buttons
        process_frame = ctk.CTkFrame(recording_frame)
        process_frame.pack(fill="x", padx=10, pady=5)

        self.process_button = ctk.CTkButton(
            process_frame,
            text="Process Current Recording",
            command=self._process_current,
            state="disabled"
        )
        self.process_button.pack(side="left", padx=5)

        self.load_file_button = ctk.CTkButton(
            process_frame,
            text="Load Audio File",
            command=self._load_audio_file
        )
        self.load_file_button.pack(side="left", padx=5)

        self.clear_button = ctk.CTkButton(
            process_frame,
            text="Clear Session",
            command=self._clear_session
        )
        self.clear_button.pack(side="left", padx=5)

    def _create_settings_section(self):
        """Create settings section."""
        settings_frame = ctk.CTkFrame(self.main_frame)
        settings_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(settings_frame, text="Settings", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=5)

        # Settings grid
        settings_grid = ctk.CTkFrame(settings_frame)
        settings_grid.pack(fill="x", padx=10, pady=5)

        # Left column
        left_settings = ctk.CTkFrame(settings_grid)
        left_settings.pack(side="left", fill="both", expand=True, padx=5)

        # Language
        lang_frame = ctk.CTkFrame(left_settings)
        lang_frame.pack(fill="x", padx=5, pady=2)
        ctk.CTkLabel(lang_frame, text="Language:").pack(side="left", padx=5)
        self.language_var = ctk.CTkOptionMenu(
            lang_frame,
            values=["id-ID", "en-US", "auto"],
            command=self._on_language_change
        )
        self.language_var.pack(side="right", padx=5)

        # Output folder
        folder_frame = ctk.CTkFrame(left_settings)
        folder_frame.pack(fill="x", padx=5, pady=2)
        ctk.CTkLabel(folder_frame, text="Output Folder:").pack(side="left", padx=5)

        # Button container for folder operations
        folder_button_frame = ctk.CTkFrame(folder_frame, fg_color="transparent")
        folder_button_frame.pack(side="right", padx=5)

        self.folder_button = ctk.CTkButton(
            folder_button_frame,
            text="Select Folder",
            command=self._select_output_folder,
            width=90
        )
        self.folder_button.pack(side="left", padx=2)

        # Open folder button (berdasarkan monolithic pattern)
        self.open_folder_button = ctk.CTkButton(
            folder_button_frame,
            text="📁 Open",
            command=self._open_output_folder,
            width=80
        )
        self.open_folder_button.pack(side="left", padx=2)

        # Right column
        right_settings = ctk.CTkFrame(settings_grid)
        right_settings.pack(side="right", fill="both", expand=True, padx=5)

        # AI Enhancement
        self.ai_enhancement_var = ctk.CTkCheckBox(
            right_settings,
            text="Enable AI Enhancement",
            command=self._on_ai_toggle
        )
        self.ai_enhancement_var.pack(anchor="w", padx=5, pady=2)

        # Auto process
        self.auto_process_var = ctk.CTkCheckBox(
            right_settings,
            text="Auto-process recordings",
            command=self._on_auto_process_toggle
        )
        self.auto_process_var.pack(anchor="w", padx=5, pady=2)

        # Create summary
        self.create_summary_var = ctk.CTkCheckBox(
            right_settings,
            text="Create summary document",
            command=self._on_summary_toggle
        )
        self.create_summary_var.pack(anchor="w", padx=5, pady=2)

    def _create_output_section(self):
        """Create output display section."""
        output_frame = ctk.CTkFrame(self.main_frame)
        output_frame.pack(fill="both", expand=True, padx=10, pady=5)

        ctk.CTkLabel(output_frame, text="Output", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=5)

        # Create tabview for different outputs
        self.output_tabs = ctk.CTkTabview(output_frame)
        self.output_tabs.pack(fill="both", expand=True, padx=10, pady=5)

        # Original transcript tab
        self.output_tabs.add("Original Transcript")
        self.original_text = scrolledtext.ScrolledText(
            self.output_tabs.tab("Original Transcript"),
            wrap=tk.WORD,
            width=80,
            height=15,
            font=("Consolas", 10)
        )
        self.original_text.pack(fill="both", expand=True, padx=5, pady=5)

        # Enhanced text tab
        self.output_tabs.add("Enhanced Text")
        self.enhanced_text = scrolledtext.ScrolledText(
            self.output_tabs.tab("Enhanced Text"),
            wrap=tk.WORD,
            width=80,
            height=15,
            font=("Calibri", 11)
        )
        self.enhanced_text.pack(fill="both", expand=True, padx=5, pady=5)

        # Export buttons
        export_frame = ctk.CTkFrame(output_frame)
        export_frame.pack(fill="x", padx=10, pady=5)

        self.export_md_button = ctk.CTkButton(
            export_frame,
            text="Export as Markdown",
            command=lambda: self._export_text("markdown"),
            state="disabled"
        )
        self.export_md_button.pack(side="left", padx=5)

        self.export_word_button = ctk.CTkButton(
            export_frame,
            text="Export as Word",
            command=lambda: self._export_text("word"),
            state="disabled"
        )
        self.export_word_button.pack(side="left", padx=5)

        self.export_txt_button = ctk.CTkButton(
            export_frame,
            text="Export as Text",
            command=lambda: self._export_text("text"),
            state="disabled"
        )
        self.export_txt_button.pack(side="left", padx=5)

    def _create_status_section(self):
        """Create status and progress section."""
        status_frame = ctk.CTkFrame(self.main_frame)
        status_frame.pack(fill="x", padx=10, pady=5)

        # Status label
        self.status_label = ctk.CTkLabel(
            status_frame,
            textvariable=self.status_var,
            font=ctk.CTkFont(size=12)
        )
        self.status_label.pack(side="left", padx=10, pady=5)

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(status_frame)
        self.progress_bar.pack(side="right", padx=10, pady=5, fill="x", expand=True)
        self.progress_bar.set(0)

    # =============================================================================
    # EVENT HANDLERS
    # =============================================================================

    def _toggle_recording(self):
        """Toggle recording state."""
        if not self.app.recording:
            # Start recording
            mode = self.recording_mode.get()
            if self.app.start_recording(mode):
                self.record_button.configure(text="⏹️ Stop Recording", fg_color="#CC3030")
                self._start_timer()
                self._disable_controls(True)
        else:
            # Stop recording
            if self.app.stop_recording():
                self.record_button.configure(text="🎙️ Start Recording", fg_color="#007ACC")
                self._stop_timer()
                self._disable_controls(False)
                self.process_button.configure(state="normal")

    def _process_current(self):
        """Process current recording."""
        if self.app.process_current_recording():
            self._disable_controls(True)
            self.progress_bar.set(0.3)
            # Progress will be updated by status callbacks

    def _load_audio_file(self):
        """Load external audio file."""
        file_path = filedialog.askopenfilename(
            title="Select Audio File",
            filetypes=[
                ("Audio Files", "*.wav *.mp3 *.flac *.m4a *.ogg"),
                ("All Files", "*.*")
            ]
        )

        if file_path:
            if self.app.process_audio_file(file_path):
                self._disable_controls(True)
                self.progress_bar.set(0.3)

    def _clear_session(self):
        """Clear current session."""
        self.app.clear_current_session()
        self.original_text.delete(1.0, tk.END)
        self.enhanced_text.delete(1.0, tk.END)
        self.process_button.configure(state="disabled")
        self._enable_export_buttons(False)
        self.progress_bar.set(0)

    def _show_api_dialog(self):
        """Show API key configuration dialog."""
        dialog = APIKeyDialog(self.root, self.app.config.get_user_api_key() or "")
        api_key = dialog.get_api_key()

        if api_key:
            if self.app.update_api_key(api_key):
                self._update_api_status()
                messagebox.showinfo("Success", "API key updated successfully!")
            else:
                messagebox.showerror("Error", "Invalid API key format. Must start with 'gsk_'")

    def _select_output_folder(self):
        """Select output folder."""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.app.set_config_value("output_folder", folder)
            self.folder_button.configure(text=f"...{folder[-20:]}")

    def _open_output_folder(self):
        """
        Open the output folder in file explorer.
        Implementasi berdasarkan monolithic system dengan threading untuk mencegah lag.
        """
        def open_folder_thread():
            try:
                folder = self.app.get_config_value("output_folder", str(Path.home() / "EchoScribe_Output"))

                if not os.path.exists(folder):
                    self.root.after(0, lambda: messagebox.showwarning("Warning", f"Folder not found:\n{folder}"))
                    return

                # Open folder using appropriate command for the OS (from monolithic pattern)
                if os.name == 'nt':  # Windows
                    os.startfile(folder)
                elif os.name == 'posix':  # macOS and Linux
                    import subprocess
                    import sys
                    if sys.platform == 'darwin':  # macOS
                        subprocess.run(['open', folder])
                    else:  # Linux
                        subprocess.run(['xdg-open', folder])

                # Update status in main thread
                self.root.after(0, lambda: self._update_status(f"Opened folder: {os.path.basename(folder)}"))
                self.root.after(3000, lambda: self._update_status("Ready"))

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to open folder: {e}"))
                self.root.after(0, lambda: self._update_status("Error opening folder"))

        # Run in background thread to prevent UI lag
        threading.Thread(target=open_folder_thread, daemon=True).start()

    def _open_file(self, filepath):
        """
        Open a file with its default application.
        Implementasi berdasarkan monolithic system dengan threading untuk mencegah lag.
        """
        def open_file_thread():
            try:
                if not os.path.exists(filepath):
                    self.root.after(0, lambda: messagebox.showwarning("Warning", f"File not found:\n{filepath}"))
                    return

                # Open file using appropriate command for the OS (from monolithic pattern)
                if os.name == 'nt':  # Windows
                    os.startfile(filepath)
                elif os.name == 'posix':  # macOS and Linux
                    import subprocess
                    import sys
                    if sys.platform == 'darwin':  # macOS
                        subprocess.run(['open', filepath])
                    else:  # Linux
                        subprocess.run(['xdg-open', filepath])

                # Update status in main thread
                self.root.after(0, lambda: self._update_status(f"Opened file: {os.path.basename(filepath)}"))
                self.root.after(3000, lambda: self._update_status("Ready"))

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to open file: {e}"))
                self.root.after(0, lambda: self._update_status("Error opening file"))

        # Run in background thread to prevent UI lag
        threading.Thread(target=open_file_thread, daemon=True).start()

    def _export_text(self, format_type: str):
        """Export current text in specified format."""
        results = self.app.get_current_results()
        if not results["enhanced_text"]:
            messagebox.showwarning("Warning", "No text to export")
            return

        # This would trigger the document processor
        # For now, just show a message
        messagebox.showinfo("Export", f"Would export as {format_type}")

    # Configuration change handlers
    def _on_mode_change(self, mode: str):
        self.app.set_config_value("recording_mode", mode)

    def _on_language_change(self, language: str):
        self.app.set_config_value("language", language)

    def _on_ai_toggle(self):
        self.app.set_config_value("use_ai_enhancement", self.ai_enhancement_var.get())

    def _on_auto_process_toggle(self):
        self.app.set_config_value("auto_process", self.auto_process_var.get())

    def _on_summary_toggle(self):
        self.app.set_config_value("create_summary", self.create_summary_var.get())

    # =============================================================================
    # UI HELPER METHODS
    # =============================================================================

    def _update_status(self, message: str):
        """Update status display (called by app controller)."""
        if self.root:
            self.root.after(0, lambda: self.status_var.set(message))

            # Update progress based on status
            if "transcription" in message.lower():
                self.root.after(0, lambda: self.progress_bar.set(0.5))
            elif "enhancement" in message.lower():
                self.root.after(0, lambda: self.progress_bar.set(0.7))
            elif "document" in message.lower():
                self.root.after(0, lambda: self.progress_bar.set(0.9))
            elif "completed" in message.lower():
                self.root.after(0, lambda: self.progress_bar.set(1.0))
                self.root.after(0, self._on_processing_complete)

    def _on_processing_complete(self):
        """Handle completion of processing."""
        self._disable_controls(False)

        # Update text displays
        results = self.app.get_current_results()
        if results["transcript"]:
            self.original_text.delete(1.0, tk.END)
            self.original_text.insert(1.0, results["transcript"])

        if results["enhanced_text"]:
            self.enhanced_text.delete(1.0, tk.END)
            self.enhanced_text.insert(1.0, results["enhanced_text"])
            self._enable_export_buttons(True)

    def _update_api_status(self):
        """Update API key status display."""
        if self.app.config.has_valid_api_key():
            self.api_status_label.configure(text="API Key: ✅ Configured")
            self.api_button.configure(text="Update API Key")
        else:
            self.api_status_label.configure(text="API Key: ❌ Not configured")
            self.api_button.configure(text="Configure API Key")

    def _disable_controls(self, disabled: bool):
        """Enable/disable UI controls during processing."""
        state = "disabled" if disabled else "normal"
        self.record_button.configure(state=state)
        self.load_file_button.configure(state=state)
        self.recording_mode.configure(state=state)

    def _enable_export_buttons(self, enabled: bool):
        """Enable/disable export buttons."""
        state = "normal" if enabled else "disabled"
        self.export_md_button.configure(state=state)
        self.export_word_button.configure(state=state)
        self.export_txt_button.configure(state=state)

    def _load_ui_config(self):
        """Load configuration values into UI."""
        # Set initial values from config
        self.recording_mode.set(self.app.get_config_value("recording_mode", "microphone"))
        self.language_var.set(self.app.get_config_value("language", "id-ID"))

        # Set checkboxes
        self.ai_enhancement_var.select() if self.app.get_config_value("use_ai_enhancement", True) else self.ai_enhancement_var.deselect()
        self.auto_process_var.select() if self.app.get_config_value("auto_process", True) else self.auto_process_var.deselect()
        self.create_summary_var.select() if self.app.get_config_value("create_summary", False) else self.create_summary_var.deselect()

        # Update folder button
        folder = self.app.get_config_value("output_folder", "")
        if folder:
            self.folder_button.configure(text=f"...{folder[-20:]}")

    # Timer functions
    def _start_timer(self):
        """Start the recording timer."""
        self.elapsed_time = 0
        self.timer_running = True
        self._update_timer()

    def _stop_timer(self):
        """Stop the recording timer."""
        self.timer_running = False

    def _update_timer(self):
        """Update timer display."""
        if self.timer_running:
            hours = self.elapsed_time // 3600
            minutes = (self.elapsed_time % 3600) // 60
            seconds = self.elapsed_time % 60
            self.timer_var.set(f"{hours:02d}:{minutes:02d}:{seconds:02d}")
            self.elapsed_time += 1
            self.root.after(1000, self._update_timer)

    def _on_closing(self):
        """Handle window closing."""
        if self.app.recording:
            if messagebox.askokcancel("Quit", "Recording in progress. Stop and quit?"):
                self.app.stop_recording()
                self.app.cleanup()
                self.root.destroy()
        else:
            self.app.cleanup()
            self.root.destroy()

    def run(self):
        """Start the UI main loop."""
        self.root.mainloop()

class APIKeyDialog:
    """Dialog for API key input."""

    def __init__(self, parent, current_key: str = ""):
        self.result = None

        # Create dialog window
        self.dialog = ctk.CTkToplevel(parent)
        self.dialog.title("Configure Groq API Key")
        self.dialog.geometry("500x300")
        self.dialog.resizable(False, False)

        # Make dialog modal
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Center dialog
        self.dialog.geometry("+%d+%d" % (parent.winfo_rootx() + 50, parent.winfo_rooty() + 50))

        self._setup_dialog(current_key)

    def _setup_dialog(self, current_key: str):
        """Setup dialog UI."""
        # Title
        title_label = ctk.CTkLabel(
            self.dialog,
            text="🔑 Configure Groq API Key",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(pady=20)

        # Instructions
        instructions = ctk.CTkTextbox(self.dialog, height=100, width=450)
        instructions.pack(padx=20, pady=10)
        instructions.insert("0.0",
            "Enter your Groq API key below. This key will be stored securely on your local machine.\n\n"
            "To get a free API key:\n"
            "1. Visit https://console.groq.com/\n"
            "2. Sign up or log in\n"
            "3. Go to API Keys section\n"
            "4. Create a new API key\n\n"
            "Your API key should start with 'gsk_'"
        )
        instructions.configure(state="disabled")

        # API Key entry
        self.api_entry = ctk.CTkEntry(
            self.dialog,
            placeholder_text="gsk_...",
            width=450,
            show="*"
        )
        self.api_entry.pack(padx=20, pady=10)

        if current_key:
            self.api_entry.insert(0, current_key)

        # Buttons
        button_frame = ctk.CTkFrame(self.dialog)
        button_frame.pack(fill="x", padx=20, pady=20)

        cancel_button = ctk.CTkButton(
            button_frame,
            text="Cancel",
            command=self._on_cancel,
            width=100
        )
        cancel_button.pack(side="right", padx=5)

        save_button = ctk.CTkButton(
            button_frame,
            text="Save",
            command=self._on_save,
            width=100
        )
        save_button.pack(side="right", padx=5)

        # Bind Enter key
        self.api_entry.bind("<Return>", lambda e: self._on_save())
        self.api_entry.focus()

    def _on_save(self):
        """Handle save button."""
        api_key = self.api_entry.get().strip()
        if api_key:
            self.result = api_key
            self.dialog.destroy()
        else:
            messagebox.showwarning("Warning", "Please enter an API key")

    def _on_cancel(self):
        """Handle cancel button."""
        self.result = None
        self.dialog.destroy()

    def get_api_key(self) -> Optional[str]:
        """Get the entered API key."""
        self.dialog.wait_window()
        return self.result

    def run(self):
        """Run the application main loop."""
        try:
            logger.info("Starting EchoScribe AI application...")
            self.root.mainloop()
        except KeyboardInterrupt:
            logger.info("Application interrupted by user")
        except Exception as e:
            logger.error(f"Error in application main loop: {e}")
        finally:
            self._cleanup()

    def _cleanup(self):
        """Cleanup resources before closing."""
        try:
            if hasattr(self, 'app') and self.app:
                # Stop any ongoing recording
                if hasattr(self.app, 'audio_recorder') and self.app.audio_recorder:
                    if hasattr(self.app.audio_recorder, 'cleanup'):
                        self.app.audio_recorder.cleanup()

                # Stop visualization if running
                if hasattr(self, 'viz_after_id') and self.viz_after_id:
                    self.root.after_cancel(self.viz_after_id)

            logger.info("Cleanup completed successfully")
        except Exception as e:
            logger.error(f"Error during cleanup: {e}")

def main():
    """Main entry point for the UI."""
    app = EchoScribeUI()
    app.run()

if __name__ == "__main__":
    main()
