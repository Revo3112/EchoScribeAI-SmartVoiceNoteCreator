# -*- coding: utf-8 -*-
"""
Audio Visualization Module for EchoScribe AI
Extracted from monolithic system with full real-time visualization support.
Supports waveform, bars, spectrum, and fill visualization modes.
"""

import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import threading
import queue
import time
import logging
import tkinter as tk
from typing import Optional, Callable

logger = logging.getLogger(__name__)

class AudioVisualizer:
    """
    Real-time audio visualization with multiple modes.
    Extracted from monolithic system (lines 1000-1500).
    """

    def __init__(self, parent_frame, status_callback: Optional[Callable[[str], None]] = None):
        self.parent_frame = parent_frame
        self.status_callback = status_callback or (lambda x: None)

        # Visualization state
        self.viz_mode = tk.StringVar(value="waveform")
        self.viz_enabled = tk.BooleanVar(value=True)
        self.viz_sensitivity = tk.DoubleVar(value=1.0)
        self.viz_running = False

        # Audio data management
        self.audio_queue = queue.Queue(maxsize=100)
        self.viz_data = np.zeros(100)
        self.spectrum_data = np.zeros((50, 50))
        self.idle_time = 0
        self.idle_spectrum_data = np.zeros((50, 50))
        self.prev_bar_data = np.zeros(20)

        # Matplotlib components
        self.viz_fig = None
        self.viz_ax = None
        self.viz_canvas = None
        self.viz_thread = None

        # Setup visualization
        self.setup_visualization()

    def setup_visualization(self):
        """Setup matplotlib visualization with proper error handling."""
        try:
            # Set matplotlib backend and configuration
            try:
                matplotlib.use('TkAgg')
            except ImportError as e:
                logger.error(f"❌ Matplotlib not available: {e}")
                self.setup_placeholder_visualization()
                return

            # Set matplotlib configuration to avoid font warnings
            plt.rcParams['font.family'] = ['Arial', 'DejaVu Sans', 'sans-serif']
            plt.rcParams['font.size'] = 10
            plt.rcParams['axes.unicode_minus'] = False

            # Create matplotlib figure with dark theme
            self.viz_fig, self.viz_ax = plt.subplots(figsize=(10, 2), facecolor='#2B2B2B')
            self.viz_fig.patch.set_facecolor('#2B2B2B')

            # Configure plot with elegant colors
            self.viz_ax.set_xlim(0, 100)
            self.viz_ax.set_ylim(-1, 1)
            self.viz_ax.set_facecolor('#1E1E1E')
            self.viz_ax.axis('off')

            # Initialize visualization elements
            self.viz_line, = self.viz_ax.plot([], [], color='#007ACC', linewidth=2, alpha=0.8)
            self.viz_bars = None
            self.viz_fill = None
            self.viz_spectrum_image = None

            # Create canvas for tkinter
            self.viz_canvas = FigureCanvasTkAgg(self.viz_fig, self.parent_frame)
            canvas_widget = self.viz_canvas.get_tk_widget()
            canvas_widget.configure(bg='#2B2B2B')
            canvas_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            # Add welcome message
            self.viz_ax.text(
                0.5, 0.5,
                'Audio Visualization Ready\\nSelect mode and start recording',
                transform=self.viz_ax.transAxes,
                ha='center', va='center',
                fontsize=12, color='#CCCCCC', alpha=0.8
            )
            self.viz_canvas.draw()

            # Start visualization thread
            self.start_visualization_update()

            logger.info("✅ Audio visualization setup successful")

        except ImportError as e:
            logger.error(f"❌ Error importing visualization libraries: {e}")
            self.setup_placeholder_visualization()
        except Exception as e:
            logger.error(f"❌ Error setting up audio visualization: {e}")
            self.setup_placeholder_visualization()

    def setup_placeholder_visualization(self):
        """Fallback placeholder with elegant styling."""
        try:
            import customtkinter as ctk

            placeholder_frame = ctk.CTkFrame(self.parent_frame, fg_color=("#F0F0F0", "#2D2D2D"))
            placeholder_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            viz_label = ctk.CTkLabel(
                placeholder_frame,
                text="Audio Visualization\\n(Real-time waveform will appear during recording)\\n\\nMatplotlib required for visualization\\nInstall with: pip install matplotlib",
                font=ctk.CTkFont(size=12),
                text_color=("#666666", "#AAAAAA"),
                justify="center"
            )
            viz_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        except ImportError:
            # Fallback to standard tkinter
            placeholder_frame = tk.Frame(self.parent_frame, bg='#2D2D2D')
            placeholder_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            viz_label = tk.Label(
                placeholder_frame,
                text="Audio Visualization\\n(Matplotlib required)",
                bg='#2D2D2D', fg='#AAAAAA',
                font=('Arial', 12),
                justify=tk.CENTER
            )
            viz_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    def start_visualization_update(self):
        """Start real-time visualization update thread."""
        if hasattr(self, 'viz_canvas') and self.viz_canvas:
            self.viz_running = True
            try:
                self.viz_thread = threading.Thread(target=self.update_visualization_loop, daemon=True)
                self.viz_thread.start()
                logger.info("Visualization thread started successfully")
            except Exception as e:
                logger.error(f"Failed to start visualization thread: {e}")
                self.viz_running = False

    def update_visualization_loop(self):
        """Loop for real-time visualization updates."""
        while self.viz_running:
            try:
                if not self.viz_enabled.get():
                    time.sleep(0.1)
                    continue

                # Check if main thread is still alive
                try:
                    if not self.parent_frame.winfo_exists():
                        logger.info("Main window closed, stopping visualization thread")
                        break
                except tk.TclError:
                    logger.info("Main window closed, stopping visualization thread")
                    break

                if not self.audio_queue.empty():
                    try:
                        # Get audio data from queue with timeout
                        audio_chunk = self.audio_queue.get(timeout=0.1)

                        # Update visualization based on mode
                        current_mode = self.viz_mode.get()
                        if current_mode == "waveform":
                            self.update_waveform_visualization(audio_chunk)
                        elif current_mode == "bars":
                            self.update_bars_visualization(audio_chunk)
                        elif current_mode == "spectrum":
                            self.update_spectrum_visualization(audio_chunk)
                        elif current_mode == "fill":
                            self.update_fill_visualization(audio_chunk)

                    except queue.Empty:
                        pass  # Normal timeout, continue loop
                    except Exception as e:
                        logger.debug(f"Error processing audio chunk: {e}")
                else:
                    # Show idle state
                    self.show_idle_visualization()

                time.sleep(0.05)  # 20 FPS update

            except Exception as e:
                logger.error(f"Error in visualization update loop: {e}")
                time.sleep(0.5)  # Longer sleep on error

        logger.info("Visualization thread ended")

    def update_waveform_visualization(self, audio_chunk):
        """Update waveform visualization."""
        try:
            if not hasattr(self, 'viz_canvas') or not hasattr(self, 'viz_sensitivity'):
                return

            if len(audio_chunk) > 0:
                # Apply sensitivity
                sensitivity = self.viz_sensitivity.get()

                # Normalize audio data
                normalized = np.array(audio_chunk, dtype=np.float32) / 32768.0 * sensitivity

                # Resample to fit display
                if len(normalized) > 100:
                    step = len(normalized) // 100
                    self.viz_data = normalized[::step][:100]
                else:
                    if len(normalized) < 100:
                        padded = np.zeros(100)
                        padded[:len(normalized)] = normalized
                        self.viz_data = padded
                    else:
                        self.viz_data = normalized[:100]

                # Clear and reset plot with proper styling
                self.viz_ax.clear()
                self.viz_ax.set_xlim(0, 100)
                self.viz_ax.set_facecolor('#1E1E1E')
                self.viz_ax.axis('off')

                # Update y-axis limits dynamically
                max_amplitude = np.max(np.abs(self.viz_data))
                if max_amplitude > 0:
                    self.viz_ax.set_ylim(-max_amplitude * 1.2, max_amplitude * 1.2)
                else:
                    self.viz_ax.set_ylim(-1, 1)

                # Update colors based on intensity with elegant gradient
                if max_amplitude > 0.7:
                    color = '#FF6B6B'  # Red for loud
                    glow_color = '#FF9999'
                elif max_amplitude > 0.3:
                    color = '#4ECDC4'  # Teal for medium
                    glow_color = '#7FFFD4'
                else:
                    color = '#007ACC'  # Blue for quiet
                    glow_color = '#4DA6FF'

                # Plot data with glow effect
                x_data = np.arange(len(self.viz_data))

                # Main line
                self.viz_ax.plot(x_data, self.viz_data, color=color, linewidth=2, alpha=0.9)

                # Glow effect
                self.viz_ax.plot(x_data, self.viz_data, color=glow_color, linewidth=4, alpha=0.3)

                # Add zero line
                self.viz_ax.axhline(y=0, color='#555555', alpha=0.5, linewidth=1)

                # Thread-safe canvas update
                if hasattr(self, 'viz_canvas'):
                    self.parent_frame.after(0, self._safe_canvas_draw)

        except Exception as e:
            logger.error(f"Error updating waveform: {e}")

    def update_bars_visualization(self, audio_chunk):
        """Update bar visualization."""
        try:
            if len(audio_chunk) > 0:
                # Apply sensitivity
                sensitivity = self.viz_sensitivity.get()

                # FFT for frequency analysis
                fft = np.fft.fft(audio_chunk * sensitivity)
                freqs = np.abs(fft[:len(fft)//2])

                # Resample to fit bars (20 bars for nice display)
                num_bars = 20
                if len(freqs) > num_bars:
                    step = len(freqs) // num_bars
                    bar_data = freqs[::step][:num_bars]
                else:
                    bar_data = np.pad(freqs, (0, max(0, num_bars - len(freqs))), 'constant')[:num_bars]

                # Normalize
                if np.max(bar_data) > 0:
                    bar_data = bar_data / np.max(bar_data)

                # Clear with proper styling
                self.viz_ax.clear()
                self.viz_ax.set_xlim(-0.5, num_bars - 0.5)
                self.viz_ax.set_ylim(0, 1.1)
                self.viz_ax.set_facecolor('#1E1E1E')
                self.viz_ax.axis('off')

                # Smoothing
                if hasattr(self, 'prev_bar_data') and self.prev_bar_data is not None:
                    smoothing_factor = 0.7
                    bar_data = smoothing_factor * self.prev_bar_data + (1 - smoothing_factor) * bar_data
                self.prev_bar_data = bar_data.copy()

                # Color gradient
                colors = []
                for i, height in enumerate(bar_data):
                    if height > 0.7:
                        colors.append('#FF6B6B')  # Red for high frequencies
                    elif height > 0.4:
                        colors.append('#4ECDC4')  # Cyan for medium
                    else:
                        colors.append('#007ACC')  # Blue for low

                # Create bars with glow effect
                bars = self.viz_ax.bar(
                    range(len(bar_data)),
                    bar_data,
                    color=colors,
                    width=0.8,
                    alpha=0.8,
                    edgecolor='white',
                    linewidth=0.5
                )

                # Add glow effect for bars
                self.viz_ax.bar(
                    range(len(bar_data)),
                    bar_data,
                    color=colors,
                    width=1.0,
                    alpha=0.3
                )

                # Thread-safe redraw
                if hasattr(self, 'viz_canvas'):
                    self.parent_frame.after(0, self._safe_canvas_draw)

        except Exception as e:
            logger.error(f"Error updating bars: {e}")

    def update_spectrum_visualization(self, audio_chunk):
        """Update spectrum visualization (waterfall)."""
        try:
            if len(audio_chunk) > 0:
                # Apply sensitivity
                sensitivity = self.viz_sensitivity.get()

                # Create spectrogram-like visualization
                fft = np.fft.fft(audio_chunk * sensitivity)
                spectrum = np.abs(fft[:len(fft)//2])

                # Initialize spectrum_data if not exists
                if not hasattr(self, 'spectrum_data') or self.spectrum_data is None:
                    self.spectrum_data = np.zeros((50, 50))

                # Shift existing data
                self.spectrum_data = np.roll(self.spectrum_data, -1, axis=1)

                # Add new column
                spectrum_height = self.spectrum_data.shape[0]
                if len(spectrum) > spectrum_height:
                    step = len(spectrum) // spectrum_height
                    self.spectrum_data[:, -1] = spectrum[::step][:spectrum_height]
                else:
                    padded_spectrum = np.zeros(spectrum_height)
                    padded_spectrum[:len(spectrum)] = spectrum
                    self.spectrum_data[:, -1] = padded_spectrum

                # Normalize
                if np.max(self.spectrum_data) > 0:
                    normalized_data = self.spectrum_data / np.max(self.spectrum_data)
                else:
                    normalized_data = self.spectrum_data

                # Clear and redraw with styling
                self.viz_ax.clear()
                self.viz_ax.set_facecolor('#1E1E1E')
                self.viz_ax.axis('off')

                # Use elegant colormap
                im = self.viz_ax.imshow(
                    normalized_data,
                    aspect='auto',
                    cmap='plasma',  # Beautiful purple-pink-yellow gradient
                    origin='lower',
                    alpha=0.8,
                    interpolation='bilinear'
                )

                # Thread-safe redraw
                if hasattr(self, 'viz_canvas'):
                    self.parent_frame.after(0, self._safe_canvas_draw)

        except Exception as e:
            logger.error(f"Error updating spectrum: {e}")

    def update_fill_visualization(self, audio_chunk):
        """Update filled area visualization."""
        try:
            if len(audio_chunk) > 0:
                # Apply sensitivity
                sensitivity = self.viz_sensitivity.get()

                # Normalize audio data
                normalized = np.array(audio_chunk, dtype=np.float32) / 32768.0 * sensitivity

                # Resample to fit display
                if len(normalized) > 100:
                    step = len(normalized) // 100
                    self.viz_data = normalized[::step][:100]
                else:
                    if len(normalized) < 100:
                        padded = np.zeros(100)
                        padded[:len(normalized)] = normalized
                        self.viz_data = padded
                    else:
                        self.viz_data = normalized[:100]

                # Clear with proper styling
                self.viz_ax.clear()
                self.viz_ax.set_xlim(0, 100)
                self.viz_ax.set_facecolor('#1E1E1E')
                self.viz_ax.axis('off')

                # Update y-axis limits
                max_amplitude = np.max(np.abs(self.viz_data))
                if max_amplitude > 0:
                    self.viz_ax.set_ylim(-max_amplitude * 1.2, max_amplitude * 1.2)
                else:
                    self.viz_ax.set_ylim(-1, 1)

                # Update colors based on intensity
                if max_amplitude > 0.7:
                    fill_color = '#FF6B6B'  # Red for loud
                    glow_color = '#FF9999'
                elif max_amplitude > 0.3:
                    fill_color = '#4ECDC4'  # Teal for medium
                    glow_color = '#7FFFD4'
                else:
                    fill_color = '#007ACC'  # Blue for quiet
                    glow_color = '#4DA6FF'

                x_data = np.arange(len(self.viz_data))

                # Create filled area with gradient effect
                # Main fill
                self.viz_ax.fill_between(
                    x_data,
                    0,
                    self.viz_data,
                    alpha=0.7,
                    color=fill_color
                )

                # Glow effect
                self.viz_ax.fill_between(
                    x_data,
                    0,
                    self.viz_data,
                    alpha=0.3,
                    color=glow_color
                )

                # Add center line and grid lines
                self.viz_ax.axhline(y=0, color='#FFFFFF', alpha=0.6, linewidth=1)
                if max_amplitude > 0:
                    self.viz_ax.axhline(y=max_amplitude * 0.5, color='#888888', alpha=0.3, linewidth=0.5, linestyle='--')
                    self.viz_ax.axhline(y=-max_amplitude * 0.5, color='#888888', alpha=0.3, linewidth=0.5, linestyle='--')

                # Thread-safe redraw
                if hasattr(self, 'viz_canvas'):
                    self.parent_frame.after(0, self._safe_canvas_draw)

        except Exception as e:
            logger.error(f"Error updating fill visualization: {e}")

    def show_idle_visualization(self):
        """Show idle state visualization."""
        try:
            if not self.viz_enabled.get():
                return

            # Check if canvas still exists
            if not hasattr(self, 'viz_ax') or not self.viz_ax:
                return

            self.idle_time += 0.1
            x = np.linspace(0, 100, 100)
            current_mode = self.viz_mode.get()

            self.viz_ax.clear()
            self.viz_ax.set_facecolor('#1E1E1E')
            self.viz_ax.axis('off')

            if current_mode == "waveform":
                self.viz_ax.set_xlim(0, 100)
                self.viz_ax.set_ylim(-0.3, 0.3)

                y1 = 0.15 * np.sin(0.2 * x + self.idle_time) * np.exp(-0.02 * x)
                y2 = 0.08 * np.sin(0.5 * x + self.idle_time * 1.5)
                y = y1 + y2

                self.viz_ax.plot(x, y, color='#007ACC', linewidth=2, alpha=0.6)
                self.viz_ax.plot(x, y, color='#4DA6FF', linewidth=4, alpha=0.3)
                self.viz_ax.axhline(y=0, color='#555555', alpha=0.5, linewidth=1)

            elif current_mode == "fill":
                self.viz_ax.set_xlim(0, 100)
                self.viz_ax.set_ylim(-0.3, 0.3)

                amplitude = 0.12 * (1 + 0.4 * np.sin(self.idle_time))
                y = amplitude * np.sin(0.3 * x + self.idle_time)

                self.viz_ax.fill_between(x, 0, y, alpha=0.5, color='#007ACC')
                self.viz_ax.fill_between(x, 0, y, alpha=0.2, color='#4DA6FF')
                self.viz_ax.axhline(y=0, color='#FFFFFF', alpha=0.6, linewidth=1)

            elif current_mode == "bars":
                self.viz_ax.set_xlim(-0.5, 19.5)
                self.viz_ax.set_ylim(0, 0.3)

                bar_data = 0.08 * np.random.random(20) * (1 + 0.3 * np.sin(self.idle_time + np.arange(20) * 0.3))
                colors = ['#007ACC' if i % 3 == 0 else '#4ECDC4' if i % 3 == 1 else '#FF6B6B' for i in range(20)]

                self.viz_ax.bar(range(20), bar_data, color=colors, width=0.8, alpha=0.6)
                self.viz_ax.bar(range(20), bar_data, color=colors, width=1.0, alpha=0.3)

            elif current_mode == "spectrum":
                if not hasattr(self, 'idle_spectrum_data') or self.idle_spectrum_data is None:
                    self.idle_spectrum_data = np.random.random((50, 50)) * 0.1

                self.idle_spectrum_data = np.roll(self.idle_spectrum_data, -1, axis=1)
                self.idle_spectrum_data[:, -1] = np.random.random(50) * 0.08 * (1 + 0.5 * np.sin(self.idle_time))

                self.viz_ax.imshow(
                    self.idle_spectrum_data,
                    aspect='auto',
                    cmap='viridis',
                    origin='lower',
                    alpha=0.4,
                    interpolation='bilinear'
                )

            # Thread-safe redraw
            if hasattr(self, 'viz_canvas'):
                self.parent_frame.after(0, self._safe_canvas_draw)

        except Exception as e:
            logger.debug(f"Error in idle visualization: {e}")

    def change_visualization_mode(self, mode):
        """Change visualization mode."""
        try:
            valid_modes = ["waveform", "bars", "spectrum", "fill"]
            if mode not in valid_modes:
                logger.warning(f"Invalid visualization mode: {mode}")
                mode = "waveform"

            self.viz_mode.set(mode)

            if hasattr(self, 'viz_ax') and self.viz_ax:
                self.viz_ax.clear()
                self.viz_ax.set_facecolor('#1E1E1E')
                self.viz_ax.axis('off')

                # Reset mode-specific data
                if mode == "bars":
                    self.prev_bar_data = np.zeros(20)
                    self.viz_ax.set_xlim(-0.5, 19.5)
                    self.viz_ax.set_ylim(0, 1.1)
                elif mode == "spectrum":
                    self.spectrum_data = np.zeros((50, 50))
                    self.idle_spectrum_data = np.zeros((50, 50))
                elif mode in ["waveform", "fill"]:
                    self.viz_data = np.zeros(100)
                    self.viz_ax.set_xlim(0, 100)
                    self.viz_ax.set_ylim(-1, 1)

                # Show mode change message
                self.viz_ax.text(
                    0.5, 0.5,
                    f'{mode.title()} Mode\\nReady for audio...',
                    transform=self.viz_ax.transAxes,
                    ha='center', va='center',
                    fontsize=12, color='#CCCCCC', alpha=0.8
                )

                # Thread-safe canvas draw
                if hasattr(self, 'viz_canvas'):
                    self.parent_frame.after(0, self._safe_canvas_draw)

            logger.info(f"Visualization mode changed to: {mode}")

        except Exception as e:
            logger.error(f"Error changing visualization mode: {e}")

    def toggle_visualization(self):
        """Toggle visualization on/off."""
        try:
            if self.viz_enabled.get():
                self.viz_running = True
                if not hasattr(self, 'viz_thread') or not self.viz_thread.is_alive():
                    self.start_visualization_update()

                if hasattr(self, 'viz_ax') and self.viz_ax:
                    self.viz_ax.clear()
                    self.viz_ax.set_facecolor('#1E1E1E')
                    self.viz_ax.axis('off')
                    self.viz_ax.set_xlim(0, 100)
                    self.viz_ax.set_ylim(-1, 1)

                    self.viz_ax.text(
                        0.5, 0.5,
                        'Visualization Enabled\\nReady for audio...',
                        transform=self.viz_ax.transAxes,
                        ha='center', va='center',
                        fontsize=12, color='#4ECDC4', alpha=0.8
                    )

                    if hasattr(self, 'viz_canvas'):
                        self.parent_frame.after(0, self._safe_canvas_draw)
            else:
                self.viz_running = False
                if hasattr(self, 'viz_ax') and self.viz_ax:
                    self.viz_ax.clear()
                    self.viz_ax.set_facecolor('#1E1E1E')
                    self.viz_ax.axis('off')
                    self.viz_ax.set_xlim(0, 100)
                    self.viz_ax.set_ylim(-1, 1)

                    self.viz_ax.text(
                        0.5, 0.5,
                        'Visualization Disabled\\nClick to enable',
                        transform=self.viz_ax.transAxes,
                        ha='center', va='center',
                        fontsize=12, color='#888888'
                    )

                    if hasattr(self, 'viz_canvas'):
                        self.parent_frame.after(0, self._safe_canvas_draw)

            logger.info(f"Visualization {'enabled' if self.viz_enabled.get() else 'disabled'}")

        except Exception as e:
            logger.error(f"Error toggling visualization: {e}")

    def _safe_canvas_draw(self):
        """Thread-safe canvas drawing method."""
        try:
            # Check if canvas and root still exist
            if (hasattr(self, 'viz_canvas') and
                self.viz_canvas and
                hasattr(self, 'parent_frame') and
                self.parent_frame.winfo_exists()):

                # Use draw_idle() for thread safety
                self.viz_canvas.draw_idle()
        except tk.TclError as e:
            # Handle TclError when window is closed
            logger.debug(f"TclError in canvas draw (window likely closed): {e}")
        except Exception as e:
            logger.debug(f"Error in safe canvas draw: {e}")

    def add_audio_data(self, audio_data):
        """Add audio data to visualization queue."""
        try:
            self.audio_queue.put_nowait(audio_data)
        except queue.Full:
            # Queue is full, skip this frame
            pass

    def stop_visualization(self):
        """Stop the visualization thread."""
        self.viz_running = False
        if hasattr(self, 'viz_thread') and self.viz_thread and self.viz_thread.is_alive():
            self.viz_thread.join(timeout=1.0)

    def cleanup(self):
        """Clean up visualization resources."""
        self.stop_visualization()
        if hasattr(self, 'viz_fig') and self.viz_fig:
            plt.close(self.viz_fig)
