# -*- coding: utf-8 -*-
"""
Modifikasi fungsi perekaman audio sistem (Versi 5 - Dual Input with Mixing).
Menambahkan mixing audio loopback dan mikrofon ke satu file output.
"""

import soundcard as sc
import numpy as np
import threading
import wave
import logging
import time
from tkinter import messagebox # Assuming tkinter is used based on original code
import queue
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
        self.recording = False
        # --- Mixing related ---
        self.mix_output = True # Flag to enable mixing
        self.mixed_output_filename_base = "mixed_audio"
        # Keep separate filenames in case mixing is disabled
        self.loopback_output_filename_base = "system_audio"
        self.mic_output_filename_base = "mic_audio"
        # --- End Mixing related ---
        self.use_extended_recording = None # Chunking is complex with mixing, disable for now
        self.chunk_size = None
        self.loopback_audio_queue = None
        self.mic_audio_queue = None
        self.viz_enabled = None
        self.loopback_chunk_count = 0 # Not used if mixing final output
        self.mic_chunk_count = 0      # Not used if mixing final output
        self.default_samplerate = 48000
        self.default_blocksize = 1024
        self.sample_width = 2 # For 16-bit WAV output
        self.loopback_device = None
        self.mic_device = None
        self.loopback_thread = None
        self.mic_thread = None
        self.loopback_frames = []
        self.mic_frames = []
        # Store detected channels
        self.loopback_channels = 0
        self.mic_channels = 0

    def _update_status(self, message):
        if self.status_var and self.root:
            self.root.after(0, lambda: self.status_var.set(message))
        logger.info(message)

    def get_audio_duration_from_frames(self, frames_list, rate, channels, sample_width):
        total_bytes = sum(len(f) for f in frames_list)
        if channels == 0 or sample_width == 0 or rate == 0: return 0
        num_frames = total_bytes / (channels * sample_width)
        duration = num_frames / rate
        return duration

    def _save_audio_to_file(self, audio_data_bytes, rate, channels, sample_width, filename):
        """Saves raw audio bytes to a WAV file."""
        if not audio_data_bytes:
            logger.warning(f"No audio data bytes to save for {filename}")
            return
        if channels == 0:
             logger.warning(f"Cannot save {filename} with 0 channels.")
             return

        try:
            with wave.open(filename, "wb") as wf:
                wf.setnchannels(channels)
                wf.setsampwidth(sample_width)
                wf.setframerate(rate)
                wf.writeframes(audio_data_bytes)
            logger.info(f"Audio saved to {filename}")
            self._update_status(f"Saved: {filename.split("/")[-1]}")
        except Exception as e:
            logger.error(f"Error saving WAV file {filename}: {e}", exc_info=True)
            self._update_status(f"Error saving {filename.split("/")[-1]}")

    # Chunk saving functions are removed as mixing is done at the end

    def process_audio_thread(self):
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
        rate = self.default_samplerate
        blocksize = self.default_blocksize
        frames = []
        detected_channels = 0
        try:
            logger.info(f"Loopback Thread: Starting recorder for {self.loopback_device.name}")
            with self.loopback_device.recorder(samplerate=rate, blocksize=blocksize) as recorder:
                while not self.stop_recording_flag:
                    try:
                        data_np = recorder.record(numframes=blocksize)
                        if data_np.size == 0:
                            continue

                        if not frames: # First valid frame
                            detected_channels = data_np.shape[1] if data_np.ndim > 1 else 1
                            self.loopback_channels = detected_channels # Store detected channels
                            logger.info(f"Loopback Thread: Detected {detected_channels} channel(s).")

                        data_int16 = (data_np * 32767).astype(np.int16)
                        data_bytes = data_int16.tobytes()
                        frames.append(data_bytes)

                    except Exception as loop_e:
                        logger.error(f"Loopback Thread: Error during recording loop: {loop_e}", exc_info=True)
                        time.sleep(0.01)

            logger.info("Loopback Thread: Recording loop finished.")
            self.loopback_frames = frames # Store collected frames

        except Exception as thread_e:
            logger.error(f"Loopback Thread: Failed to start or run recorder: {thread_e}", exc_info=True)
            self._update_status(f"ERROR: Loopback recording failed: {thread_e}")

    # --- Thread target for microphone recording ---
    def _record_microphone_thread_target(self):
        rate = self.default_samplerate
        blocksize = self.default_blocksize
        frames = []
        detected_channels = 0
        try:
            logger.info(f"Microphone Thread: Starting recorder for {self.mic_device.name}")
            with self.mic_device.recorder(samplerate=rate, blocksize=blocksize) as recorder:
                while not self.stop_recording_flag:
                    try:
                        data_np = recorder.record(numframes=blocksize)
                        if data_np.size == 0:
                            continue

                        if not frames: # First valid frame
                            detected_channels = data_np.shape[1] if data_np.ndim > 1 else 1
                            self.mic_channels = detected_channels # Store detected channels
                            logger.info(f"Microphone Thread: Detected {detected_channels} channel(s).")

                        data_int16 = (data_np * 32767).astype(np.int16)
                        data_bytes = data_int16.tobytes()
                        frames.append(data_bytes)

                    except Exception as loop_e:
                        logger.error(f"Microphone Thread: Error during recording loop: {loop_e}", exc_info=True)
                        time.sleep(0.01)

            logger.info("Microphone Thread: Recording loop finished.")
            self.mic_frames = frames # Store collected frames

        except Exception as thread_e:
            logger.error(f"Microphone Thread: Failed to start or run recorder: {thread_e}", exc_info=True)
            self._update_status(f"ERROR: Microphone recording failed: {thread_e}")

    # --- Function to mix and save audio ---
    def _mix_and_save_audio(self):
        """Mixes loopback and mic audio frames and saves to a single file."""
        logger.info("Starting audio mixing process...")
        rate = self.default_samplerate
        sample_width = self.sample_width

        if not self.loopback_frames or not self.mic_frames:
            logger.warning("One or both audio sources have no frames. Cannot mix.")
            # Optionally save the one that does have frames?
            if self.loopback_frames:
                 filename = f"{self.loopback_output_filename_base}_final.wav"
                 self._save_audio_to_file(b"".join(self.loopback_frames), rate, self.loopback_channels, sample_width, filename)
            if self.mic_frames:
                 filename = f"{self.mic_output_filename_base}_final.wav"
                 self._save_audio_to_file(b"".join(self.mic_frames), rate, self.mic_channels, sample_width, filename)
            return

        try:
            # 1. Combine frames and convert to numpy int16 arrays
            loopback_bytes = b"".join(self.loopback_frames)
            mic_bytes = b"".join(self.mic_frames)

            loopback_audio = np.frombuffer(loopback_bytes, dtype=np.int16)
            mic_audio = np.frombuffer(mic_bytes, dtype=np.int16)

            # Use detected channels, default to 1 if detection failed (channel=0)
            loopback_ch = self.loopback_channels if self.loopback_channels > 0 else 1
            mic_ch = self.mic_channels if self.mic_channels > 0 else 1
            logger.info(f"Mixing: Loopback channels={loopback_ch}, Mic channels={mic_ch}")

            # 2. Determine target channels (prefer stereo if either is stereo)
            target_channels = 2 if (loopback_ch == 2 or mic_ch == 2) else 1
            logger.info(f"Mixing: Target channels={target_channels}")

            # 3. Reshape and convert mono to stereo if necessary
            if loopback_ch == 1 and target_channels == 2:
                logger.info("Mixing: Converting loopback to stereo.")
                loopback_audio = np.repeat(loopback_audio[:, np.newaxis], 2, axis=1).flatten()
            elif loopback_ch == 2 and target_channels == 1:
                 logger.warning("Mixing: Converting loopback from stereo to mono (averaging channels). This might lose information.")
                 loopback_audio = loopback_audio.reshape(-1, 2).mean(axis=1).astype(np.int16)
            # Reshape even if channels match target, to ensure correct dimensions
            try: loopback_audio = loopback_audio.reshape(-1, target_channels)
            except ValueError as e: logger.error(f"Error reshaping loopback audio: {e}. Length: {len(loopback_audio)}"); return

            if mic_ch == 1 and target_channels == 2:
                logger.info("Mixing: Converting mic to stereo.")
                mic_audio = np.repeat(mic_audio[:, np.newaxis], 2, axis=1).flatten()
            elif mic_ch == 2 and target_channels == 1:
                 logger.warning("Mixing: Converting mic from stereo to mono (averaging channels). This might lose information.")
                 mic_audio = mic_audio.reshape(-1, 2).mean(axis=1).astype(np.int16)
            # Reshape even if channels match target
            try: mic_audio = mic_audio.reshape(-1, target_channels)
            except ValueError as e: logger.error(f"Error reshaping mic audio: {e}. Length: {len(mic_audio)}"); return

            # 4. Equalize length (trim longer audio)
            min_len = min(len(loopback_audio), len(mic_audio))
            if len(loopback_audio) != len(mic_audio):
                logger.warning(f"Mixing: Audio lengths differ (Loopback: {len(loopback_audio)}, Mic: {len(mic_audio)}). Trimming to {min_len}.")
                loopback_audio = loopback_audio[:min_len]
                mic_audio = mic_audio[:min_len]

            # 5. Convert to float for mixing
            loopback_float = loopback_audio.astype(np.float32) / 32767.0
            mic_float = mic_audio.astype(np.float32) / 32767.0

            # 6. Mix (simple addition - might need normalization)
            mixed_float = loopback_float + mic_float

            # 7. Normalize to prevent clipping (check max absolute value)
            max_abs = np.max(np.abs(mixed_float))
            if max_abs > 1.0:
                logger.warning(f"Mixing: Potential clipping detected (max abs: {max_abs:.2f}). Normalizing.")
                mixed_float /= max_abs
            else:
                logger.info(f"Mixing: Max absolute value is {max_abs:.2f}, no normalization needed.")

            # 8. Convert back to int16
            mixed_int16 = (mixed_float * 32767).astype(np.int16)

            # 9. Convert back to bytes
            mixed_bytes = mixed_int16.tobytes()

            # 10. Save the mixed audio
            filename = f"{self.mixed_output_filename_base}_final.wav"
            self._save_audio_to_file(mixed_bytes, rate, target_channels, sample_width, filename)
            logger.info("Mixing process completed successfully.")

        except Exception as mix_e:
            logger.error(f"Error during audio mixing: {mix_e}", exc_info=True)
            self._update_status(f"ERROR: Audio mixing failed: {mix_e}")
            # Optionally save unmixed files as fallback?
            # filename_lb = f"{self.loopback_output_filename_base}_final_unmixed.wav"
            # self._save_audio_to_file(b"".join(self.loopback_frames), rate, self.loopback_channels, sample_width, filename_lb)
            # filename_mic = f"{self.mic_output_filename_base}_final_unmixed.wav"
            # self._save_audio_to_file(b"".join(self.mic_frames), rate, self.mic_channels, sample_width, filename_mic)

    # --- Main function to start dual recording ---
    def start_dual_recording(self):
        self._update_status("Initializing Dual Audio Capture (SoundCard V5 - Mixing)...")
        self.recording = False
        self.loopback_device = None
        self.mic_device = None
        self.loopback_frames = []
        self.mic_frames = []
        self.loopback_channels = 0 # Reset detected channels
        self.mic_channels = 0
        # Chunking is disabled for mixing version
        # self.loopback_chunk_count = 0
        # self.mic_chunk_count = 0

        # 1. Identify Loopback Device (same as V4)
        try:
            default_speaker = sc.default_speaker()
            if not default_speaker: self._show_enhanced_system_audio_troubleshooting("Cannot find default speaker."); return
            logger.info(f"Default speaker found: {default_speaker.name} (ID: {default_speaker.id})")
            self.loopback_device = sc.get_microphone(id=str(default_speaker.id), include_loopback=True)
            if not self.loopback_device:
                 logger.warning(f"Could not find loopback mic by speaker ID, trying generic loopback search...")
                 all_mics = sc.all_microphones(include_loopback=True)
                 for mic in all_mics:
                     if 'loopback' in mic.name.lower() or 'monitor' in mic.name.lower():
                         self.loopback_device = mic
                         logger.info(f"Found potential loopback fallback: {mic.name}")
                         break
                 if not self.loopback_device:
                     self._show_enhanced_system_audio_troubleshooting(f"No loopback device found for {default_speaker.name}."); return
            logger.info(f"Loopback device selected: {self.loopback_device.name}")
        except Exception as e_loopback: self._show_enhanced_system_audio_troubleshooting(f"Error finding loopback device: {e_loopback}"); return

        # 2. Identify Microphone Device (same as V4)
        try:
            self.mic_device = sc.default_microphone()
            if not self.mic_device: self._show_enhanced_system_audio_troubleshooting("Cannot find default microphone."); return
            logger.info(f"Microphone device selected: {self.mic_device.name}")
        except Exception as e_mic: self._show_enhanced_system_audio_troubleshooting(f"Error finding microphone device: {e_mic}"); return

        # --- Devices identified, proceed to start threads ---
        self._update_status(f"Rec: Loopback [{self.loopback_device.name}] + Mic [{self.mic_device.name}]")
        self.stop_recording_flag = False
        self.recording = True

        # 3. Create and Start Threads
        self.loopback_thread = threading.Thread(target=self._record_loopback_thread_target, daemon=True)
        self.mic_thread = threading.Thread(target=self._record_microphone_thread_target, daemon=True)

        self._update_status("🔴 Recording Loopback & Microphone...")
        self.loopback_thread.start()
        self.mic_thread.start()

    # --- Function to stop recording ---
    def stop_dual_recording(self):
        if not self.recording: logger.info("No recording is currently active."); return
        self._update_status("Stopping recording...")
        self.stop_recording_flag = True

        # Wait for threads
        if self.loopback_thread and self.loopback_thread.is_alive(): logger.info("Waiting for loopback thread..."); self.loopback_thread.join(); logger.info("Loopback thread finished.")
        if self.mic_thread and self.mic_thread.is_alive(): logger.info("Waiting for microphone thread..."); self.mic_thread.join(); logger.info("Microphone thread finished.")

        self.recording = False
        self._update_status("Recording stopped. Mixing and saving final file...")

        # 4. Mix and Save final audio data
        # Chunking is disabled in this version, so we always mix the final result
        self._mix_and_save_audio()

        # Clear frames after saving/mixing
        self.loopback_frames = []
        self.mic_frames = []
        self._update_status("Dual recording finished and mixed file saved.")

# Example Usage (if run standalone)
if __name__ == '__main__':
    root = tk.Tk()
    root.title("SoundCard Dual Recorder Test V5 (Mixing)")
    root.geometry("350x200")

    app = AudioRecorderApp(root)

    app.status_var = tk.StringVar(root, "Idle")
    # Chunking checkbox removed as it's disabled for mixing
    # app.use_extended_recording = tk.BooleanVar(root, False)
    # app.chunk_size = tk.IntVar(root, 10)
    app.viz_enabled = tk.BooleanVar(root, False)

    status_label = ttk.Label(root, textvariable=app.status_var)
    status_label.pack(pady=10)

    def start_rec_action():
        if not app.recording:
            app.start_dual_recording()
        else:
            app._update_status("Already recording!")

    def stop_rec_action():
        if app.recording:
            app._update_status("Stop requested...")
            stop_thread = threading.Thread(target=app.stop_dual_recording, daemon=True)
            stop_thread.start()
        else:
             app._update_status("Not currently recording.")

    start_button = ttk.Button(root, text="Start Dual Recording (Mixed Output)", command=start_rec_action)
    start_button.pack(pady=5)

    stop_button = ttk.Button(root, text="Stop Recording", command=stop_rec_action)
    stop_button.pack(pady=5)

    # chk_extended = ttk.Checkbutton(root, text="Use Chunking", variable=app.use_extended_recording)
    # chk_extended.pack(pady=5)

    root.mainloop()
