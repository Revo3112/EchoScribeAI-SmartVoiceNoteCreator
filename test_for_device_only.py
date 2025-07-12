# -*- coding: utf-8 -*-
"""
Modifikasi fungsi perekaman audio sistem menggunakan library soundcard (Versi 3).
Perbaikan berdasarkan feedback error pengguna dan contoh penggunaan loopback yang benar.
"""

import soundcard as sc
import numpy as np
import threading
import wave
import logging
import time
from tkinter import messagebox # Assuming tkinter is used based on original code

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
        self.recording = False
        self.use_extended_recording = None
        self.chunk_size = None
        self.audio_queue = None
        self.viz_enabled = None
        self.output_filename_base = "system_audio_recording"
        self.chunk_count = 0
        self.default_samplerate = 48000
        self.default_channels = 2 # Default to stereo, recorder might adjust
        self.default_blocksize = 1024
        self.sample_width = 2

    def _update_status(self, message):
        if self.status_var and self.root:
            self.root.after(0, lambda: self.status_var.set(message))
        logger.info(message)

    def get_audio_duration_from_frames(self, frames_list, rate, channels, sample_width):
        total_bytes = sum(len(f) for f in frames_list)
        num_frames = total_bytes / (channels * sample_width)
        duration = num_frames / rate
        return duration

    def _save_audio_to_file(self, frames_list, rate, channels, sample_width, filename):
        try:
            with wave.open(filename, "wb") as wf:
                wf.setnchannels(channels)
                wf.setsampwidth(sample_width)
                wf.setframerate(rate)
                wf.writeframes(b"".join(frames_list))
            logger.info(f"Audio saved to {filename}")
            self._update_status(f"Saved: {filename.split('/')[-1]}")
        except Exception as e:
            logger.error(f"Error saving WAV file {filename}: {e}", exc_info=True)
            self._update_status(f"Error saving {filename.split('/')[-1]}")

    def _save_system_audio_chunk_optimized(self, frames_list, rate, channels, sample_width):
        self.chunk_count += 1
        filename = f"{self.output_filename_base}_chunk_{self.chunk_count}.wav"
        # Ensure channels match the actual recorded data if possible
        # For simplicity, we use the requested default channels here
        self._save_audio_to_file(frames_list, rate, channels, sample_width, filename)

    def _save_system_audio_to_file_optimized(self, frames_list, rate, channels, sample_width):
        filename = f"{self.output_filename_base}_final.wav"
        self._save_audio_to_file(frames_list, rate, channels, sample_width, filename)

    def process_audio_thread(self):
        logger.info("Starting post-processing thread (placeholder)...")
        time.sleep(1)
        logger.info("Post-processing finished (placeholder).")
        self._update_status("Processing complete.")

    def _show_enhanced_system_audio_troubleshooting(self, error_message):
        logger.warning(f"Troubleshooting needed: {error_message}")
        if self.root:
             self.root.after(0, lambda: messagebox.showerror("Audio Recording Error", f"Could not record system audio. Please ensure audio is playing and check audio settings.\n\nError: {error_message}"))

    # --- REVISED record_system_audio_soundcard (V3 - Correct Loopback Method) ---
    def record_system_audio_soundcard(self):
        """
        Record system audio (loopback) using the CORRECT soundcard method (V3).
        Uses get_microphone with include_loopback=True.
        """
        self._update_status("Initializing system audio capture (SoundCard V3)...")
        frames = []
        rate = self.default_samplerate
        blocksize = self.default_blocksize
        sample_width = self.sample_width
        actual_channels = self.default_channels # Start with default, update if possible

        try:
            # 1. Get the default speaker to find its name/ID
            default_speaker = sc.default_speaker()
            if not default_speaker:
                error_msg = "SoundCard could not find default speaker."
                logger.error(error_msg)
                self._update_status("ERROR: Cannot find default speaker")
                self._show_enhanced_system_audio_troubleshooting(error_msg)
                return

            logger.info(f"Default speaker found: {default_speaker.name}")

            # 2. Get the microphone corresponding to the default speaker WITH loopback enabled
            # The ID is crucial for some backends (like PulseAudio) to find the correct monitor source
            loopback_mic = sc.get_microphone(id=str(default_speaker.id), include_loopback=True)

            if not loopback_mic:
                 error_msg = f"Could not find loopback microphone for speaker: {default_speaker.name}"
                 logger.error(error_msg)
                 self._update_status(f"ERROR: No loopback for {default_speaker.name}")
                 self._show_enhanced_system_audio_troubleshooting(error_msg)
                 return

            logger.info(f"Loopback microphone found: {loopback_mic.name}")
            self._update_status(f"Recording loopback from: {loopback_mic.name} @ {rate}Hz")

            # 3. Start recording using the loopback microphone's recorder
            self.recording = True
            self.stop_recording_flag = False
            self.chunk_count = 0
            frames = []

            self._update_status("🔴 Recording system audio (SoundCard V3)...")

            # Use the loopback microphone object to get the recorder
            with loopback_mic.recorder(samplerate=rate, blocksize=blocksize) as recorder:
                # Update actual channels based on the recorder if possible (might vary)
                # Note: recorder object itself might not expose channels directly,
                # data shape will tell us later.

                while not self.stop_recording_flag and self.recording:
                    try:
                        data_np = recorder.record(numframes=blocksize)
                        if data_np.size == 0:
                            continue

                        # Determine actual channels from first valid data chunk
                        if not frames: # Only check on the first frame
                            if data_np.ndim > 1:
                                actual_channels = data_np.shape[1]
                            else: # Mono
                                actual_channels = 1
                            logger.info(f"Detected {actual_channels} channel(s) in recording.")
                            # Update default if it differs? Or just use actual_channels for saving.
                            # self.default_channels = actual_channels

                        # Convert float NumPy array to 16-bit integer bytes
                        data_int16 = (data_np * 32767).astype(np.int16)
                        data_bytes = data_int16.tobytes()
                        frames.append(data_bytes)

                        # Update visualization (if applicable)
                        if hasattr(self, 'audio_queue') and hasattr(self, 'viz_enabled') and self.viz_enabled.get():
                            if not self.audio_queue.full():
                                self.audio_queue.put_nowait(data_np)

                        # Extended recording chunking
                        if self.use_extended_recording and self.use_extended_recording.get() and len(frames) > 0:
                            # Use actual_channels determined from data for saving
                            chunk_duration = self.get_audio_duration_from_frames(frames, rate, actual_channels, sample_width)
                            if chunk_duration >= self.chunk_size.get():
                                self._save_system_audio_chunk_optimized(frames, rate, actual_channels, sample_width)
                                frames = []

                    except Exception as e:
                        logger.error(f"Error during recording loop: {e}", exc_info=True)
                        time.sleep(0.01)

            logger.info("Recording loop finished.")
            self._update_status("Finishing recording...")

        except Exception as e:
            error_msg = f"SoundCard Recording Error (V3): {e}"
            logger.error(error_msg, exc_info=True)
            self._update_status(f"ERROR: {error_msg}")
            self._show_enhanced_system_audio_troubleshooting(str(e))
            self.recording = False
            return
        finally:
            self.recording = False

        # 4. Save the recorded audio
        if len(frames) > 0:
             # Ensure actual_channels is set correctly if recording was very short
            if actual_channels == self.default_channels and frames:
                 try:
                     first_frame_np = np.frombuffer(frames[0], dtype=np.int16).astype(np.float32) / 32767
                     # Simple check for stereo based on expected blocksize vs frame length
                     expected_mono_len = blocksize
                     expected_stereo_len = blocksize * 2
                     if len(first_frame_np) == expected_stereo_len:
                         actual_channels = 2
                         logger.info("Confirmed 2 channels for saving.")
                     elif len(first_frame_np) == expected_mono_len:
                         actual_channels = 1
                         logger.info("Confirmed 1 channel for saving.")
                 except Exception as e_ch_check:
                     logger.warning(f"Could not confirm channel count from first frame: {e_ch_check}")

            if self.use_extended_recording and self.use_extended_recording.get():
                # Save the last chunk if any frames remain
                self._save_system_audio_chunk_optimized(frames, rate, actual_channels, sample_width)
            else:
                # Save the entire recording as one file
                self._save_system_audio_to_file_optimized(frames, rate, actual_channels, sample_width)

            self.root.after(100, lambda: threading.Thread(target=self.process_audio_thread, daemon=True).start())
        else:
            logger.warning("No audio frames captured with SoundCard V3.")
            self._update_status("No audio captured")
            if self.root:
                self.root.after(0, lambda: messagebox.showwarning(
                    "No Audio Captured",
                    "No system audio was detected or captured using SoundCard. Make sure audio is playing."
                ))

# Example Usage (if run standalone)
if __name__ == '__main__':
    import tkinter as tk
    from tkinter import ttk
    import queue

    root = tk.Tk()
    root.title("SoundCard Recorder Test V3")
    root.geometry("300x200")

    app = AudioRecorderApp(root)

    app.status_var = tk.StringVar(root, "Idle")
    app.use_extended_recording = tk.BooleanVar(root, False)
    app.chunk_size = tk.IntVar(root, 10)
    app.audio_queue = queue.Queue(maxsize=100)
    app.viz_enabled = tk.BooleanVar(root, False)

    status_label = ttk.Label(root, textvariable=app.status_var)
    status_label.pack(pady=10)

    def start_rec():
        if not app.recording:
            app.recording = True
            app.stop_recording_flag = False
            rec_thread = threading.Thread(target=app.record_system_audio_soundcard, daemon=True)
            rec_thread.start()

    def stop_rec():
        app.stop_recording_flag = True

    start_button = ttk.Button(root, text="Start Recording System Audio", command=start_rec)
    start_button.pack(pady=5)

    stop_button = ttk.Button(root, text="Stop Recording", command=stop_rec)
    stop_button.pack(pady=5)

    chk_extended = ttk.Checkbutton(root, text="Use Chunking", variable=app.use_extended_recording)
    chk_extended.pack(pady=5)

    root.mainloop()
