# -*- coding: utf-8 -*-
"""
Value Tracking Slider Component for EchoScribe AI
Enhanced slider with real-time value display and tooltip functionality.
Extracted from monolithic system (lines 299-370).
"""

import tkinter as tk
try:
    import customtkinter as ctk
except ImportError:
    ctk = None
import logging
from typing import Optional, Callable

logger = logging.getLogger(__name__)

class ValueTrackingSlider:
    """
    Enhanced slider with real-time value display and professional styling.
    Extracted from monolithic system with complete functionality.
    """

    def __init__(self,
                 parent,
                 from_: float = 0.0,
                 to: float = 1.0,
                 initial_value: float = 0.5,
                 orientation: str = "horizontal",
                 label: str = "Value",
                 format_string: str = "{:.2f}",
                 callback: Optional[Callable[[float], None]] = None,
                 width: int = 200,
                 height: int = 20):
        """
        Initialize value tracking slider with advanced features.

        Args:
            parent: Parent widget
            from_: Minimum value
            to: Maximum value
            initial_value: Initial slider value
            orientation: "horizontal" or "vertical"
            label: Label text for the slider
            format_string: Format string for value display
            callback: Callback function called when value changes
            width: Slider width
            height: Slider height
        """
        self.parent = parent
        self.from_ = from_
        self.to = to
        self.orientation = orientation
        self.label_text = label
        self.format_string = format_string
        self.callback = callback
        self.width = width
        self.height = height

        # Current value
        self.current_value = initial_value

        # UI components
        self.frame = None
        self.label = None
        self.slider = None
        self.value_label = None
        self.tooltip = None

        # Tooltip state
        self.tooltip_visible = False
        self.tooltip_window = None

        self.create_widgets()

    def create_widgets(self):
        """Create the slider widgets with professional styling."""
        try:
            # Create main frame
            if ctk:
                self.frame = ctk.CTkFrame(self.parent, fg_color="transparent")
            else:
                self.frame = tk.Frame(self.parent, bg='#2D2D2D')

            self.frame.pack(fill=tk.X, padx=5, pady=2)

            # Create label
            label_text = f"{self.label_text}:"
            if ctk:
                self.label = ctk.CTkLabel(
                    self.frame,
                    text=label_text,
                    font=ctk.CTkFont(size=12),
                    width=80,
                    anchor="w"
                )
            else:
                self.label = tk.Label(
                    self.frame,
                    text=label_text,
                    bg='#2D2D2D',
                    fg='#FFFFFF',
                    font=('Arial', 10),
                    width=12,
                    anchor="w"
                )

            self.label.pack(side=tk.LEFT, padx=(5, 10))

            # Create value display label
            if ctk:
                self.value_label = ctk.CTkLabel(
                    self.frame,
                    text=self.format_string.format(self.current_value),
                    font=ctk.CTkFont(size=11, weight="bold"),
                    width=60,
                    fg_color=("#E0E0E0", "#3A3A3A"),
                    corner_radius=6
                )
            else:
                self.value_label = tk.Label(
                    self.frame,
                    text=self.format_string.format(self.current_value),
                    bg='#3A3A3A',
                    fg='#FFFFFF',
                    font=('Arial', 9, 'bold'),
                    width=8,
                    relief=tk.RAISED,
                    bd=1
                )

            self.value_label.pack(side=tk.RIGHT, padx=(10, 5))

            # Create slider
            if ctk:
                self.slider = ctk.CTkSlider(
                    self.frame,
                    from_=self.from_,
                    to=self.to,
                    number_of_steps=100,
                    orientation=self.orientation,
                    width=self.width,
                    height=self.height,
                    command=self._on_value_change,
                    button_color=("#007ACC", "#005A9E"),
                    button_hover_color=("#4DA6FF", "#007ACC"),
                    progress_color=("#007ACC", "#005A9E")
                )
                self.slider.set(self.current_value)
            else:
                self.slider = tk.Scale(
                    self.frame,
                    from_=self.from_,
                    to=self.to,
                    orient=tk.HORIZONTAL if self.orientation == "horizontal" else tk.VERTICAL,
                    resolution=0.01,
                    length=self.width,
                    bg='#2D2D2D',
                    fg='#FFFFFF',
                    troughcolor='#1E1E1E',
                    activebackground='#007ACC',
                    highlightbackground='#2D2D2D',
                    command=self._on_value_change,
                    showvalue=0
                )
                self.slider.set(self.current_value)

            self.slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

            # Bind tooltip events
            self._bind_tooltip_events()

            logger.debug(f"ValueTrackingSlider created: {self.label_text}")

        except Exception as e:
            logger.error(f"Error creating ValueTrackingSlider: {e}")

    def _bind_tooltip_events(self):
        """Bind tooltip events to slider."""
        try:
            if self.slider:
                self.slider.bind("<Enter>", self._show_tooltip)
                self.slider.bind("<Leave>", self._hide_tooltip)
                self.slider.bind("<Motion>", self._update_tooltip)

                if hasattr(self.slider, 'bind'):
                    # For CustomTkinter, also bind to button events
                    self.slider.bind("<Button-1>", self._show_tooltip)
                    self.slider.bind("<ButtonRelease-1>", self._hide_tooltip)

        except Exception as e:
            logger.debug(f"Error binding tooltip events: {e}")

    def _on_value_change(self, value):
        """Handle slider value change."""
        try:
            # Convert string value if needed
            if isinstance(value, str):
                value = float(value)

            self.current_value = value

            # Update value display
            if self.value_label:
                formatted_value = self.format_string.format(value)
                if ctk and hasattr(self.value_label, 'configure'):
                    self.value_label.configure(text=formatted_value)
                elif hasattr(self.value_label, 'config'):
                    self.value_label.config(text=formatted_value)

            # Update tooltip if visible
            if self.tooltip_visible:
                self._update_tooltip_content()

            # Call callback if provided
            if self.callback:
                try:
                    self.callback(value)
                except Exception as e:
                    logger.debug(f"Error in slider callback: {e}")

        except Exception as e:
            logger.error(f"Error handling slider value change: {e}")

    def _show_tooltip(self, event=None):
        """Show tooltip with current value."""
        try:
            if self.tooltip_window or self.tooltip_visible:
                return

            self.tooltip_visible = True

            # Create tooltip window
            self.tooltip_window = tk.Toplevel(self.parent)
            self.tooltip_window.wm_overrideredirect(True)
            self.tooltip_window.configure(bg='#2D2D2D', highlightbackground='#007ACC', highlightthickness=1)

            # Position tooltip
            x = self.slider.winfo_rootx() + 20
            y = self.slider.winfo_rooty() - 30
            self.tooltip_window.geometry(f"+{x}+{y}")

            # Create tooltip content
            self._update_tooltip_content()

        except Exception as e:
            logger.debug(f"Error showing tooltip: {e}")

    def _hide_tooltip(self, event=None):
        """Hide tooltip."""
        try:
            self.tooltip_visible = False
            if self.tooltip_window:
                self.tooltip_window.destroy()
                self.tooltip_window = None

        except Exception as e:
            logger.debug(f"Error hiding tooltip: {e}")

    def _update_tooltip(self, event=None):
        """Update tooltip position and content."""
        try:
            if self.tooltip_visible and self.tooltip_window:
                # Update position
                x = self.slider.winfo_rootx() + 20
                y = self.slider.winfo_rooty() - 30
                self.tooltip_window.geometry(f"+{x}+{y}")

                # Update content
                self._update_tooltip_content()

        except Exception as e:
            logger.debug(f"Error updating tooltip: {e}")

    def _update_tooltip_content(self):
        """Update tooltip content with current value."""
        try:
            if not self.tooltip_window:
                return

            # Clear existing content
            for widget in self.tooltip_window.winfo_children():
                widget.destroy()

            # Create tooltip text
            tooltip_text = f"{self.label_text}: {self.format_string.format(self.current_value)}"

            if ctk:
                tooltip_label = ctk.CTkLabel(
                    self.tooltip_window,
                    text=tooltip_text,
                    font=ctk.CTkFont(size=10),
                    fg_color="transparent",
                    text_color=("#000000", "#FFFFFF")
                )
            else:
                tooltip_label = tk.Label(
                    self.tooltip_window,
                    text=tooltip_text,
                    bg='#2D2D2D',
                    fg='#FFFFFF',
                    font=('Arial', 9),
                    padx=5,
                    pady=2
                )

            tooltip_label.pack()

        except Exception as e:
            logger.debug(f"Error updating tooltip content: {e}")

    def get_value(self) -> float:
        """Get current slider value."""
        return self.current_value

    def set_value(self, value: float):
        """Set slider value programmatically."""
        try:
            if self.from_ <= value <= self.to:
                self.current_value = value
                if self.slider:
                    if ctk and hasattr(self.slider, 'set'):
                        self.slider.set(value)
                    elif hasattr(self.slider, 'set'):
                        self.slider.set(value)

                # Update display
                self._on_value_change(value)

        except Exception as e:
            logger.error(f"Error setting slider value: {e}")

    def configure(self, **kwargs):
        """Configure slider properties."""
        try:
            if 'from_' in kwargs:
                self.from_ = kwargs['from_']
                if self.slider:
                    if ctk and hasattr(self.slider, 'configure'):
                        self.slider.configure(from_=self.from_)
                    elif hasattr(self.slider, 'config'):
                        self.slider.config(from_=self.from_)

            if 'to' in kwargs:
                self.to = kwargs['to']
                if self.slider:
                    if ctk and hasattr(self.slider, 'configure'):
                        self.slider.configure(to=self.to)
                    elif hasattr(self.slider, 'config'):
                        self.slider.config(to=self.to)

            if 'label' in kwargs:
                self.label_text = kwargs['label']
                if self.label:
                    label_text = f"{self.label_text}:"
                    if ctk and hasattr(self.label, 'configure'):
                        self.label.configure(text=label_text)
                    elif hasattr(self.label, 'config'):
                        self.label.config(text=label_text)

        except Exception as e:
            logger.error(f"Error configuring slider: {e}")

    def destroy(self):
        """Clean up slider resources."""
        try:
            self._hide_tooltip()
            if self.frame:
                self.frame.destroy()

        except Exception as e:
            logger.debug(f"Error destroying slider: {e}")
