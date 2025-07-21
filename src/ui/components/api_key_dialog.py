# -*- coding: utf-8 -*-
"""
API Key Dialog Component for EchoScribe AI
Secure API key entry dialog with validation and testing.
Extracted from monolithic system (lines 372-577).
"""

import tkinter as tk
from tkinter import messagebox
try:
    import customtkinter as ctk
except ImportError:
    ctk = None
import logging
import threading
from typing import Optional, Callable

logger = logging.getLogger(__name__)

class ApiKeyDialog:
    """
    Secure API key entry dialog with validation and testing functionality.
    Extracted from monolithic system with complete security features.
    """

    def __init__(self,
                 parent,
                 title: str = "API Key Configuration",
                 current_key: str = "",
                 test_callback: Optional[Callable[[str], bool]] = None,
                 save_callback: Optional[Callable[[str], bool]] = None):
        """
        Initialize API key dialog.

        Args:
            parent: Parent window
            title: Dialog title
            current_key: Current API key (will be masked)
            test_callback: Function to test API key validity
            save_callback: Function to save API key
        """
        self.parent = parent
        self.title = title
        self.current_key = current_key
        self.test_callback = test_callback
        self.save_callback = save_callback

        # Dialog state
        self.result = None
        self.dialog = None
        self.key_entry = None
        self.test_button = None
        self.save_button = None
        self.cancel_button = None
        self.status_label = None
        self.show_key_var = None
        self.show_key_checkbox = None

        # Testing state
        self.testing_in_progress = False

        self.create_dialog()

    def create_dialog(self):
        """Create the API key dialog with professional styling."""
        try:
            # Create dialog window
            self.dialog = tk.Toplevel(self.parent)
            self.dialog.title(self.title)
            self.dialog.geometry("500x300")
            self.dialog.resizable(False, False)
            self.dialog.grab_set()  # Make dialog modal

            # Center dialog on parent
            self._center_dialog()

            # Configure dialog styling
            if ctk:
                self.dialog.configure(fg_color=("#F0F0F0", "#2D2D2D"))
            else:
                self.dialog.configure(bg='#2D2D2D')

            # Create main frame
            if ctk:
                main_frame = ctk.CTkFrame(self.dialog, fg_color="transparent")
            else:
                main_frame = tk.Frame(self.dialog, bg='#2D2D2D')
            main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

            # Title label
            if ctk:
                title_label = ctk.CTkLabel(
                    main_frame,
                    text="Groq API Key Configuration",
                    font=ctk.CTkFont(size=18, weight="bold"),
                    text_color=("#000000", "#FFFFFF")
                )
            else:
                title_label = tk.Label(
                    main_frame,
                    text="Groq API Key Configuration",
                    bg='#2D2D2D',
                    fg='#FFFFFF',
                    font=('Arial', 14, 'bold')
                )
            title_label.pack(pady=(0, 20))

            # Instructions
            instructions = (
                "Enter your Groq API key to enable AI transcription and enhancement features.\n"
                "Your API key will be stored securely on your local machine.\n\n"
                "Get your free API key at: https://console.groq.com/keys"
            )

            if ctk:
                instructions_label = ctk.CTkLabel(
                    main_frame,
                    text=instructions,
                    font=ctk.CTkFont(size=12),
                    text_color=("#333333", "#CCCCCC"),
                    justify="left",
                    wraplength=450
                )
            else:
                instructions_label = tk.Label(
                    main_frame,
                    text=instructions,
                    bg='#2D2D2D',
                    fg='#CCCCCC',
                    font=('Arial', 10),
                    justify=tk.LEFT,
                    wraplength=450
                )
            instructions_label.pack(pady=(0, 20))

            # API Key entry frame
            if ctk:
                key_frame = ctk.CTkFrame(main_frame)
            else:
                key_frame = tk.Frame(main_frame, bg='#3A3A3A', relief=tk.RAISED, bd=1)
            key_frame.pack(fill=tk.X, pady=(0, 15))

            # API Key label
            if ctk:
                key_label = ctk.CTkLabel(
                    key_frame,
                    text="API Key:",
                    font=ctk.CTkFont(size=12, weight="bold")
                )
            else:
                key_label = tk.Label(
                    key_frame,
                    text="API Key:",
                    bg='#3A3A3A',
                    fg='#FFFFFF',
                    font=('Arial', 10, 'bold')
                )
            key_label.pack(anchor=tk.W, padx=15, pady=(15, 5))

            # API Key entry
            if ctk:
                self.key_entry = ctk.CTkEntry(
                    key_frame,
                    width=400,
                    height=35,
                    font=ctk.CTkFont(size=11),
                    show="*",
                    placeholder_text="gsk_..."
                )
            else:
                self.key_entry = tk.Entry(
                    key_frame,
                    width=50,
                    font=('Courier', 10),
                    show="*",
                    bg='#1E1E1E',
                    fg='#FFFFFF',
                    insertbackground='#FFFFFF',
                    relief=tk.SUNKEN,
                    bd=2
                )

            # Set current key if provided
            if self.current_key:
                self.key_entry.insert(0, self.current_key)

            self.key_entry.pack(padx=15, pady=(0, 10))

            # Show/Hide key checkbox
            self.show_key_var = tk.BooleanVar(value=False)
            if ctk:
                self.show_key_checkbox = ctk.CTkCheckBox(
                    key_frame,
                    text="Show API Key",
                    variable=self.show_key_var,
                    command=self._toggle_key_visibility,
                    font=ctk.CTkFont(size=10)
                )
            else:
                self.show_key_checkbox = tk.Checkbutton(
                    key_frame,
                    text="Show API Key",
                    variable=self.show_key_var,
                    command=self._toggle_key_visibility,
                    bg='#3A3A3A',
                    fg='#CCCCCC',
                    selectcolor='#1E1E1E',
                    font=('Arial', 9)
                )
            self.show_key_checkbox.pack(anchor=tk.W, padx=15, pady=(0, 15))

            # Status label
            if ctk:
                self.status_label = ctk.CTkLabel(
                    main_frame,
                    text="",
                    font=ctk.CTkFont(size=11),
                    text_color=("#666666", "#AAAAAA")
                )
            else:
                self.status_label = tk.Label(
                    main_frame,
                    text="",
                    bg='#2D2D2D',
                    fg='#AAAAAA',
                    font=('Arial', 9)
                )
            self.status_label.pack(pady=(0, 20))

            # Button frame
            if ctk:
                button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
            else:
                button_frame = tk.Frame(main_frame, bg='#2D2D2D')
            button_frame.pack(fill=tk.X)

            # Test button
            if ctk:
                self.test_button = ctk.CTkButton(
                    button_frame,
                    text="Test API Key",
                    width=120,
                    height=35,
                    font=ctk.CTkFont(size=12),
                    fg_color=("#007ACC", "#005A9E"),
                    hover_color=("#4DA6FF", "#007ACC"),
                    command=self._test_api_key
                )
            else:
                self.test_button = tk.Button(
                    button_frame,
                    text="Test API Key",
                    width=12,
                    bg='#007ACC',
                    fg='#FFFFFF',
                    font=('Arial', 10),
                    relief=tk.RAISED,
                    bd=2,
                    command=self._test_api_key
                )
            self.test_button.pack(side=tk.LEFT, padx=(0, 10))

            # Save button
            if ctk:
                self.save_button = ctk.CTkButton(
                    button_frame,
                    text="Save & Apply",
                    width=120,
                    height=35,
                    font=ctk.CTkFont(size=12),
                    fg_color=("#4ECDC4", "#3A9B94"),
                    hover_color=("#7FFFD4", "#4ECDC4"),
                    command=self._save_api_key
                )
            else:
                self.save_button = tk.Button(
                    button_frame,
                    text="Save & Apply",
                    width=12,
                    bg='#4ECDC4',
                    fg='#000000',
                    font=('Arial', 10),
                    relief=tk.RAISED,
                    bd=2,
                    command=self._save_api_key
                )
            self.save_button.pack(side=tk.LEFT, padx=(0, 10))

            # Cancel button
            if ctk:
                self.cancel_button = ctk.CTkButton(
                    button_frame,
                    text="Cancel",
                    width=100,
                    height=35,
                    font=ctk.CTkFont(size=12),
                    fg_color=("#666666", "#555555"),
                    hover_color=("#888888", "#777777"),
                    command=self._cancel
                )
            else:
                self.cancel_button = tk.Button(
                    button_frame,
                    text="Cancel",
                    width=10,
                    bg='#666666',
                    fg='#FFFFFF',
                    font=('Arial', 10),
                    relief=tk.RAISED,
                    bd=2,
                    command=self._cancel
                )
            self.cancel_button.pack(side=tk.RIGHT)

            # Bind events
            self.key_entry.bind('<Return>', lambda e: self._test_api_key())
            self.key_entry.bind('<KeyRelease>', self._on_key_change)
            self.dialog.bind('<Escape>', lambda e: self._cancel())

            # Focus on entry
            self.key_entry.focus_set()

            logger.info("API Key dialog created successfully")

        except Exception as e:
            logger.error(f"Error creating API key dialog: {e}")

    def _center_dialog(self):
        """Center dialog on parent window."""
        try:
            self.dialog.update_idletasks()

            # Get dialog dimensions
            dialog_width = self.dialog.winfo_reqwidth()
            dialog_height = self.dialog.winfo_reqheight()

            # Get parent dimensions and position
            parent_x = self.parent.winfo_rootx()
            parent_y = self.parent.winfo_rooty()
            parent_width = self.parent.winfo_width()
            parent_height = self.parent.winfo_height()

            # Calculate center position
            x = parent_x + (parent_width // 2) - (dialog_width // 2)
            y = parent_y + (parent_height // 2) - (dialog_height // 2)

            # Ensure dialog stays on screen
            x = max(0, x)
            y = max(0, y)

            self.dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        except Exception as e:
            logger.debug(f"Error centering dialog: {e}")

    def _toggle_key_visibility(self):
        """Toggle API key visibility."""
        try:
            if self.show_key_var.get():
                if ctk and hasattr(self.key_entry, 'configure'):
                    self.key_entry.configure(show="")
                elif hasattr(self.key_entry, 'config'):
                    self.key_entry.config(show="")
            else:
                if ctk and hasattr(self.key_entry, 'configure'):
                    self.key_entry.configure(show="*")
                elif hasattr(self.key_entry, 'config'):
                    self.key_entry.config(show="*")

        except Exception as e:
            logger.debug(f"Error toggling key visibility: {e}")

    def _on_key_change(self, event=None):
        """Handle API key entry changes."""
        try:
            key = self.key_entry.get().strip()

            # Basic validation
            if not key:
                self._update_status("Enter your Groq API key", "info")
            elif not key.startswith("gsk_"):
                self._update_status("Groq API keys should start with 'gsk_'", "warning")
            elif len(key) < 20:
                self._update_status("API key appears too short", "warning")
            else:
                self._update_status("API key format looks valid", "success")

        except Exception as e:
            logger.debug(f"Error handling key change: {e}")

    def _update_status(self, message: str, status_type: str = "info"):
        """Update status label with colored message."""
        try:
            if not self.status_label:
                return

            # Color mapping
            colors = {
                "info": ("#666666", "#AAAAAA"),
                "warning": ("#FF8C00", "#FFA500"),
                "error": ("#FF6B6B", "#FF4444"),
                "success": ("#4ECDC4", "#7FFFD4")
            }

            color = colors.get(status_type, colors["info"])

            if ctk and hasattr(self.status_label, 'configure'):
                self.status_label.configure(text=message, text_color=color)
            elif hasattr(self.status_label, 'config'):
                self.status_label.config(text=message, fg=color[1])

        except Exception as e:
            logger.debug(f"Error updating status: {e}")

    def _test_api_key(self):
        """Test API key validity."""
        try:
            if self.testing_in_progress:
                return

            key = self.key_entry.get().strip()

            if not key:
                self._update_status("Please enter an API key first", "warning")
                return

            if not key.startswith("gsk_"):
                self._update_status("Invalid API key format", "error")
                return

            # Disable test button
            self.testing_in_progress = True
            if ctk and hasattr(self.test_button, 'configure'):
                self.test_button.configure(text="Testing...", state="disabled")
            elif hasattr(self.test_button, 'config'):
                self.test_button.config(text="Testing...", state=tk.DISABLED)

            self._update_status("Testing API key...", "info")

            # Test in background thread
            test_thread = threading.Thread(target=self._test_api_key_background, args=(key,), daemon=True)
            test_thread.start()

        except Exception as e:
            logger.error(f"Error testing API key: {e}")
            self._test_complete(False, f"Error testing API key: {e}")

    def _test_api_key_background(self, key: str):
        """Test API key in background thread."""
        try:
            success = False
            message = "API key test failed"

            if self.test_callback:
                try:
                    success = self.test_callback(key)
                    if success:
                        message = "API key is valid and working!"
                    else:
                        message = "API key test failed - please check your key"
                except Exception as e:
                    message = f"API key test error: {str(e)[:50]}..."
            else:
                # Basic format validation if no test callback
                if key.startswith("gsk_") and len(key) > 20:
                    success = True
                    message = "API key format is valid (cannot test without connection)"
                else:
                    message = "Invalid API key format"

            # Update UI in main thread
            self.dialog.after(0, lambda: self._test_complete(success, message))

        except Exception as e:
            logger.error(f"Error in background API key test: {e}")
            self.dialog.after(0, lambda: self._test_complete(False, f"Test error: {e}"))

    def _test_complete(self, success: bool, message: str):
        """Handle test completion."""
        try:
            self.testing_in_progress = False

            # Re-enable test button
            if ctk and hasattr(self.test_button, 'configure'):
                self.test_button.configure(text="Test API Key", state="normal")
            elif hasattr(self.test_button, 'config'):
                self.test_button.config(text="Test API Key", state=tk.NORMAL)

            # Update status
            status_type = "success" if success else "error"
            self._update_status(message, status_type)

        except Exception as e:
            logger.debug(f"Error completing test: {e}")

    def _save_api_key(self):
        """Save API key."""
        try:
            key = self.key_entry.get().strip()

            if not key:
                self._update_status("Please enter an API key", "warning")
                return

            if not key.startswith("gsk_"):
                self._update_status("Invalid API key format", "error")
                return

            # Save API key
            success = True
            if self.save_callback:
                try:
                    success = self.save_callback(key)
                except Exception as e:
                    success = False
                    logger.error(f"Error saving API key: {e}")
                    self._update_status(f"Error saving: {e}", "error")
                    return

            if success:
                self.result = key
                self._update_status("API key saved successfully!", "success")

                # Close dialog after short delay
                self.dialog.after(1000, self._close_dialog)
            else:
                self._update_status("Failed to save API key", "error")

        except Exception as e:
            logger.error(f"Error saving API key: {e}")
            self._update_status("Error saving API key", "error")

    def _cancel(self):
        """Cancel dialog."""
        self.result = None
        self._close_dialog()

    def _close_dialog(self):
        """Close dialog safely."""
        try:
            if self.dialog:
                self.dialog.grab_release()
                self.dialog.destroy()

        except Exception as e:
            logger.debug(f"Error closing dialog: {e}")

    def show(self) -> Optional[str]:
        """Show dialog and return result."""
        try:
            # Wait for dialog to complete
            self.dialog.wait_window()
            return self.result

        except Exception as e:
            logger.error(f"Error showing dialog: {e}")
            return None
