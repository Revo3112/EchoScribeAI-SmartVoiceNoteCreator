# -*- coding: utf-8 -*-
"""
Modifikasi fungsi perekaman audio sistem (Versi 4 - Dual Input).
Menambahkan identifikasi perangkat mikrofon dan loopback.
"""

import soundcard as sc
import numpy as np
import threading
import wave
import logging
import time
from tkinter import messagebox # Assuming tkinter is used based on original code
import queue # Added for potential future use with threads
import tkinter as tk
from tkinter import ttk

# Setup basic logging
try:
    logger = logging.getLogger(__name__)
    if not logger.hasHandlers():
        handler = logging.StreamHandler()
        formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
except NameError:
    import sys
    logging.basicConfig(level=logging.INFO, stream=sys.stdout,
                        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
    logger = logging.getLogger(__name__)

# Placeholder for the main application class structure
class AudioRecorderApp:
    def __init__(self, root):
        self.root = root
        self.status_var = None
        self.stop_recording_flag = False
        self.recording = False # Flag to indicate if any recording is active
        self.use_extended_recording = None
        self.chunk_size = None
        # Separate queues/lists for visualization or data handling if needed
        self.loopback_audio_queue = None
        self.mic_audio_queue = None
        self.viz_enabled = None
        # Separate filenames
        self.loopback_output_filename_base = "system_audio"
        self.mic_output_filename_base = "mic_audio"
        self.loopback_chunk_count = 0
        self.mic_chunk_count = 0
        # Audio parameters
        self.default_samplerate = 48000
        self.default_blocksize = 1024
        self.sample_width = 2 # For 16-bit WAV output
        # Device storage
        self.loopback_device = None
        self.mic_device = None
        # Thread storage
        self.loopback_thread = None
        self.mic_thread = None
        # Frame storage
        self.loopback_frames = []
        self.mic_frames = []

    def _update_status(self, message):
        if self.status_var and self.root:
            self.root.after(0, lambda: self.status_var.set(message))
        logger.info(message)

    def get_audio_duration_from_frames(self, frames_list, rate, channels, sample_width):
        total_bytes = sum(len(f) for f in frames_list)
        if channels == 0 or sample_width == 0 or rate == 0: return 0 # Avoid division by zero
        num_frames = total_bytes / (channels * sample_width)
        duration = num_frames / rate
        return duration

    def _save_audio_to_file(self, frames_list, rate, channels, sample_width, filename):
        if not frames_list:
            logger.warning(f"No frames to save for {filename}")
            return
        if channels == 0:
             logger.warning(f"Cannot save {filename} with 0 channels detected.")
             # Fallback or default? For now, just warn and skip.
             # channels = 1 # Or maybe try to infer again?
             return

        try:
            with wave.open(filename, "wb") as wf:
                wf.setnchannels(channels)
                wf.setsampwidth(sample_width)
                wf.setframerate(rate)
                wf.writeframes(b"".join(frames_list))
            logger.info(f"Audio saved to {filename}")
            self._update_status(f"Saved: {filename.split("/")[-1]}")
        except Exception as e:
            logger.error(f"Error saving WAV file {filename}: {e}", exc_info=True)
            self._update_status(f"Error saving {filename.split("/")[-1]}")

    # --- Separate saving functions for clarity ---
    def _save_loopback_chunk(self, frames_list, rate, channels, sample_width):
        self.loopback_chunk_count += 1
        filename = f"{self.loopback_output_filename_base}_chunk_{self.loopback_chunk_count}.wav"
        self._save_audio_to_file(frames_list, rate, channels, sample_width, filename)

    def _save_mic_chunk(self, frames_list, rate, channels, sample_width):
        self.mic_chunk_count += 1
        filename = f"{self.mic_output_filename_base}_chunk_{self.mic_chunk_count}.wav"
        self._save_audio_to_file(frames_list, rate, channels, sample_width, filename)

    def _save_loopback_final(self, frames_list, rate, channels, sample_width):
        filename = f"{self.loopback_output_filename_base}_final.wav"
        self._save_audio_to_file(frames_list, rate, channels, sample_width, filename)

    def _save_mic_final(self, frames_list, rate, channels, sample_width):
        filename = f"{self.mic_output_filename_base}_final.wav"
        self._save_audio_to_file(frames_list, rate, channels, sample_width, filename)
    # --- End separate saving functions ---

    def process_audio_thread(self):
        # This might need adjustment if processing both files
        logger.info("Starting post-processing thread (placeholder)...")
        time.sleep(1)
        logger.info("Post-processing finished (placeholder).")
        self._update_status("Processing complete.")

    def _show_enhanced_system_audio_troubleshooting(self, error_message):
        logger.warning(f"Troubleshooting needed: {error_message}")
        if self.root:
             self.root.after(0, lambda: messagebox.showerror("Audio Recording Error", f"Could not record audio. Please check audio devices and settings.\n\nError: {error_message}"))

    # --- Thread target for loopback recording ---
    def _record_loopback_thread_target(self):
        """Target function for the loopback recording thread."""
        rate = self.default_samplerate
        blocksize = self.default_blocksize
        sample_width = self.sample_width
        actual_channels = 0 # Detect channels from data
        frames = [] # Local list for this thread

        try:
            logger.info(f"Loopback Thread: Starting recorder for {self.loopback_device.name}")
            with self.loopback_device.recorder(samplerate=rate, blocksize=blocksize) as recorder:
                while not self.stop_recording_flag:
                    try:
                        data_np = recorder.record(numframes=blocksize)
                        if data_np.size == 0:
                            continue

                        if not frames: # First valid frame
                            actual_channels = data_np.shape[1] if data_np.ndim > 1 else 1
                            logger.info(f"Loopback Thread: Detected {actual_channels} channel(s).")

                        data_int16 = (data_np * 32767).astype(np.int16)
                        data_bytes = data_int16.tobytes()
                        frames.append(data_bytes)

                        # Optional: Put data in queue for visualization/other processing
                        # if self.loopback_audio_queue and self.viz_enabled.get(): etc.

                        # Chunking logic (if enabled)
                        if self.use_extended_recording and self.use_extended_recording.get() and len(frames) > 0:
                            chunk_duration = self.get_audio_duration_from_frames(frames, rate, actual_channels, sample_width)
                            if chunk_duration >= self.chunk_size.get():
                                self._save_loopback_chunk(frames, rate, actual_channels, sample_width)
                                frames = [] # Reset local frames

                    except Exception as loop_e:
                        logger.error(f"Loopback Thread: Error during recording loop: {loop_e}", exc_info=True)
                        time.sleep(0.01)

            logger.info("Loopback Thread: Recording loop finished.")
            # Store collected frames
            self.loopback_frames = frames
            # Save final chunk/file if needed (handled in main thread after join)

        except Exception as thread_e:
            logger.error(f"Loopback Thread: Failed to start or run recorder: {thread_e}", exc_info=True)
            self._update_status(f"ERROR: Loopback recording failed: {thread_e}")
            # Optionally signal main thread or set an error flag

    # --- Thread target for microphone recording ---
    def _record_microphone_thread_target(self):
        """Target function for the microphone recording thread."""
        rate = self.default_samplerate
        blocksize = self.default_blocksize
        sample_width = self.sample_width
        actual_channels = 0 # Detect channels from data
        frames = [] # Local list for this thread

        try:
            logger.info(f"Microphone Thread: Starting recorder for {self.mic_device.name}")
            with self.mic_device.recorder(samplerate=rate, blocksize=blocksize) as recorder:
                while not self.stop_recording_flag:
                    try:
                        data_np = recorder.record(numframes=blocksize)
                        if data_np.size == 0:
                            continue

                        if not frames: # First valid frame
                            actual_channels = data_np.shape[1] if data_np.ndim > 1 else 1
                            logger.info(f"Microphone Thread: Detected {actual_channels} channel(s).")

                        data_int16 = (data_np * 32767).astype(np.int16)
                        data_bytes = data_int16.tobytes()
                        frames.append(data_bytes)

                        # Optional: Put data in queue for visualization/other processing
                        # if self.mic_audio_queue and self.viz_enabled.get(): etc.

                        # Chunking logic (if enabled)
                        if self.use_extended_recording and self.use_extended_recording.get() and len(frames) > 0:
                            chunk_duration = self.get_audio_duration_from_frames(frames, rate, actual_channels, sample_width)
                            if chunk_duration >= self.chunk_size.get():
                                self._save_mic_chunk(frames, rate, actual_channels, sample_width)
                                frames = [] # Reset local frames

                    except Exception as loop_e:
                        logger.error(f"Microphone Thread: Error during recording loop: {loop_e}", exc_info=True)
                        time.sleep(0.01)

            logger.info("Microphone Thread: Recording loop finished.")
            # Store collected frames
            self.mic_frames = frames
            # Save final chunk/file if needed (handled in main thread after join)

        except Exception as thread_e:
            logger.error(f"Microphone Thread: Failed to start or run recorder: {thread_e}", exc_info=True)
            self._update_status(f"ERROR: Microphone recording failed: {thread_e}")
            # Optionally signal main thread or set an error flag

    # --- Main function to start dual recording ---
    def start_dual_recording(self):
        """
        Identifies devices and starts parallel recording threads for loopback and microphone.
        """
        self._update_status("Initializing Dual Audio Capture (SoundCard V4)...")
        self.recording = False # Reset recording status
        self.loopback_device = None
        self.mic_device = None
        self.loopback_frames = [] # Clear previous frames
        self.mic_frames = []      # Clear previous frames
        self.loopback_chunk_count = 0 # Reset counters
        self.mic_chunk_count = 0
        rate = self.default_samplerate # Use consistent rate
        sample_width = self.sample_width

        # 1. Identify Loopback Device
        try:
            default_speaker = sc.default_speaker()
            if not default_speaker:
                self._show_enhanced_system_audio_troubleshooting("Cannot find default speaker.")
                return
            logger.info(f"Default speaker found: {default_speaker.name} (ID: {default_speaker.id})")

            # Use speaker ID for potentially better matching on some systems
            self.loopback_device = sc.get_microphone(id=str(default_speaker.id), include_loopback=True)
            if not self.loopback_device:
                 # Fallback: try finding any loopback device if ID match fails
                 logger.warning(f"Could not find loopback mic by speaker ID, trying generic loopback search...")
                 all_mics = sc.all_microphones(include_loopback=True)
                 for mic in all_mics:
                     # Heuristic: Check if 'loopback' or 'monitor' is in the name (less reliable)
                     if 'loopback' in mic.name.lower() or 'monitor' in mic.name.lower():
                         self.loopback_device = mic
                         logger.info(f"Found potential loopback fallback: {mic.name}")
                         break
                 if not self.loopback_device:
                     self._show_enhanced_system_audio_troubleshooting(f"No loopback device found for {default_speaker.name}.")
                     return

            logger.info(f"Loopback device selected: {self.loopback_device.name}")

        except Exception as e_loopback:
            self._show_enhanced_system_audio_troubleshooting(f"Error finding loopback device: {e_loopback}")
            return

        # 2. Identify Microphone Device
        try:
            self.mic_device = sc.default_microphone()
            if not self.mic_device:
                self._show_enhanced_system_audio_troubleshooting("Cannot find default microphone.")
                return
            logger.info(f"Microphone device selected: {self.mic_device.name}")
        except Exception as e_mic:
            self._show_enhanced_system_audio_troubleshooting(f"Error finding microphone device: {e_mic}")
            return

        # --- Devices identified, proceed to start threads ---
        self._update_status(f"Rec: Loopback [{self.loopback_device.name}] + Mic [{self.mic_device.name}]")
        self.stop_recording_flag = False
        self.recording = True # Set recording active flag

        # 3. Create and Start Threads
        self.loopback_thread = threading.Thread(target=self._record_loopback_thread_target, daemon=True)
        self.mic_thread = threading.Thread(target=self._record_microphone_thread_target, daemon=True)

        self._update_status("🔴 Recording Loopback & Microphone...")
        self.loopback_thread.start()
        self.mic_thread.start()

        # Optional: Add a mechanism to monitor thread health if needed

    # --- Function to stop recording ---
    def stop_dual_recording(self):
        """Signals recording threads to stop and waits for them to finish."""
        if not self.recording:
            logger.info("No recording is currently active.")
            return

        self._update_status("Stopping recording...")
        self.stop_recording_flag = True

        # Wait for threads to complete
        if self.loopback_thread and self.loopback_thread.is_alive():
            logger.info("Waiting for loopback thread to finish...")
            self.loopback_thread.join() # Wait indefinitely
            logger.info("Loopback thread finished.")
        if self.mic_thread and self.mic_thread.is_alive():
            logger.info("Waiting for microphone thread to finish...")
            self.mic_thread.join() # Wait indefinitely
            logger.info("Microphone thread finished.")

        self.recording = False # Mark recording as inactive
        self._update_status("Recording stopped. Saving final files...")

        # 4. Save final audio data (after threads have finished)
        rate = self.default_samplerate
        sample_width = self.sample_width

        # Determine channels for saving (use detected channels if available)
        loopback_channels = 0
        if self.loopback_frames:
            try: # Re-check channel count from first frame just in case
                 first_frame_np = np.frombuffer(self.loopback_frames[0], dtype=np.int16)
                 loopback_channels = first_frame_np.reshape(-1, self.default_blocksize).shape[0] if first_frame_np.ndim > 0 else 1 # Infer based on blocksize? Risky. Better rely on detection in thread.
                 # A better way might be to pass the detected channel count back from the thread
                 # For now, let's assume the detection in the thread worked or use a default if empty
                 # We need a more robust way to get the channel count back from the thread.
                 # Let's try inferring from the first frame's shape if possible.
                 temp_data = np.frombuffer(self.loopback_frames[0], dtype=np.int16)
                 # Simple check: if total samples / blocksize > 1, assume stereo? Very heuristic.
                 if temp_data.size > self.default_blocksize: loopback_channels = 2
                 else: loopback_channels = 1
                 logger.info(f"Final check: Loopback channels inferred as {loopback_channels}")
            except Exception:
                 logger.warning("Could not infer loopback channels, defaulting to 1")
                 loopback_channels = 1 # Default fallback

        mic_channels = 0
        if self.mic_frames:
            try:
                 temp_data = np.frombuffer(self.mic_frames[0], dtype=np.int16)
                 if temp_data.size > self.default_blocksize: mic_channels = 2
                 else: mic_channels = 1
                 logger.info(f"Final check: Mic channels inferred as {mic_channels}")
            except Exception:
                 logger.warning("Could not infer mic channels, defaulting to 1")
                 mic_channels = 1 # Default fallback

        # Save remaining frames if chunking wasn't used or if there are leftovers
        if not (self.use_extended_recording and self.use_extended_recording.get()):
            self._save_loopback_final(self.loopback_frames, rate, loopback_channels, sample_width)
            self._save_mic_final(self.mic_frames, rate, mic_channels, sample_width)
        else:
            # Save any remaining frames as the last chunk
            if self.loopback_frames:
                self._save_loopback_chunk(self.loopback_frames, rate, loopback_channels, sample_width)
            if self.mic_frames:
                 self._save_mic_chunk(self.mic_frames, rate, mic_channels, sample_width)

        # Clear frames after saving
        self.loopback_frames = []
        self.mic_frames = []

        # Optional: Start post-processing
        # self.root.after(100, lambda: threading.Thread(target=self.process_audio_thread, daemon=True).start())
        self._update_status("Dual recording finished and saved.")


# Example Usage (if run standalone)
if __name__ == '__main__':
    root = tk.Tk()
    root.title("SoundCard Dual Recorder Test V4")
    root.geometry("350x250")

    app = AudioRecorderApp(root)

    app.status_var = tk.StringVar(root, "Idle")
    app.use_extended_recording = tk.BooleanVar(root, False)
    app.chunk_size = tk.IntVar(root, 10)
    # Queues not used in this basic example, but could be for visualization
    # app.loopback_audio_queue = queue.Queue(maxsize=100)
    # app.mic_audio_queue = queue.Queue(maxsize=100)
    app.viz_enabled = tk.BooleanVar(root, False)

    status_label = ttk.Label(root, textvariable=app.status_var)
    status_label.pack(pady=10)

    def start_rec_action():
        # Call the new start function
        if not app.recording:
            # Run device identification and thread starting in main thread initially
            app.start_dual_recording()
        else:
            app._update_status("Already recording!")

    def stop_rec_action():
        # Call the new stop function
        if app.recording:
            # Run the stop logic (flag setting, joining, saving) in main thread or separate thread?
            # Running join in main thread will block UI. Better to run stop in a separate thread.
            app._update_status("Stop requested...")
            stop_thread = threading.Thread(target=app.stop_dual_recording, daemon=True)
            stop_thread.start()
        else:
             app._update_status("Not currently recording.")

    start_button = ttk.Button(root, text="Start Dual Recording", command=start_rec_action)
    start_button.pack(pady=5)

    stop_button = ttk.Button(root, text="Stop Recording", command=stop_rec_action)
    stop_button.pack(pady=5)

    chk_extended = ttk.Checkbutton(root, text="Use Chunking", variable=app.use_extended_recording)
    chk_extended.pack(pady=5)

    root.mainloop()
