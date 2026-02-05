import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import speech_recognition as sr
from pydub import AudioSegment  # pip install pydub
import subprocess
import platform

SUPPORTED_EXTENSIONS = (".wav", ".mp3", ".m4a")

def check_ffmpeg_installed():
    try:
        subprocess.run(["ffmpeg", "-version"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except FileNotFoundError:
        return False

def show_ffmpeg_install_instructions():
    system = platform.system()
    if system == "Windows":
        instructions = (
            "FFmpeg is not installed.\n\n"
            "1. Download FFmpeg from https://ffmpeg.org/download.html\n"
            "2. Extract the ZIP and copy the 'bin' folder path.\n"
            "3. Add that path to your System Environment Variables under 'Path'."
        )
    elif system == "Darwin":
        instructions = (
            "FFmpeg is not installed.\n\n"
            "Install it using Homebrew:\n"
            "1. Open Terminal\n"
            "2. Run: brew install ffmpeg"
        )
    elif system == "Linux":
        instructions = (
            "FFmpeg is not installed.\n\n"
            "Install it using your package manager:\n"
            "Debian/Ubuntu: sudo apt install ffmpeg\n"
            "Fedora: sudo dnf install ffmpeg"
        )
    else:
        instructions = "FFmpeg is not installed and your OS could not be identified."

    messagebox.showerror("FFmpeg Not Found", instructions)

def log_message(message):
    log_text.config(state=tk.NORMAL)
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)
    log_text.config(state=tk.DISABLED)
    root.update_idletasks()

def convert_to_pcm_wav(input_path, output_path):
    try:
        audio = AudioSegment.from_file(input_path)
        audio = audio.set_channels(1).set_frame_rate(16000)
        audio.export(output_path, format="wav")
        return output_path
    except Exception as e:
        log_message(f"Conversion failed for {input_path}: {e}")
        return None

def transcribe_audio_files(input_folder):
    recognizer = sr.Recognizer()
    # Adjust recognizer settings for better accuracy
    recognizer.energy_threshold = 300
    recognizer.dynamic_energy_threshold = True
    
    files = [f for f in os.listdir(input_folder) if f.lower().endswith(SUPPORTED_EXTENSIONS)]

    transcript_folder = os.path.join(input_folder, "Transcripts")
    os.makedirs(transcript_folder, exist_ok=True)

    progress_bar["maximum"] = len(files)
    progress_bar["value"] = 0
    root.update_idletasks()

    for idx, filename in enumerate(files):
        original_path = os.path.join(input_folder, filename)
        converted_path = os.path.join(input_folder, "converted_" + os.path.splitext(filename)[0] + ".wav")

        valid_path = convert_to_pcm_wav(original_path, converted_path)
        if not valid_path:
            continue

        txt_path = os.path.join(transcript_folder, os.path.splitext(filename)[0] + ".txt")

        try:
            # Get audio duration
            audio = AudioSegment.from_wav(valid_path)
            duration_seconds = len(audio) / 1000.0
            log_message(f"🎵 Processing {filename} ({duration_seconds:.1f} seconds)")
            
            # For long files, chunk them into smaller segments (30 seconds each)
            chunk_duration_ms = 30000  # 30 seconds in milliseconds
            full_transcript = []
            
            with sr.AudioFile(valid_path) as source:
                if duration_seconds <= 30:
                    # Short file - process all at once
                    audio_data = recognizer.record(source)
                    text = recognizer.recognize_google(audio_data, language="en-US")
                    full_transcript.append(text)
                else:
                    # Long file - process in chunks
                    log_message(f"⏱️ File is long ({duration_seconds:.1f}s), processing in chunks...")
                    chunk_num = 0
                    while True:
                        # Record a chunk (30 seconds)
                        audio_data = recognizer.record(source, duration=30)
                        if len(audio_data.frame_data) == 0:
                            break
                        
                        chunk_num += 1
                        try:
                            text = recognizer.recognize_google(audio_data, language="en-US")
                            full_transcript.append(text)
                            log_message(f"  ✓ Chunk {chunk_num} transcribed")
                        except sr.UnknownValueError:
                            log_message(f"  ⚠️ Chunk {chunk_num}: no speech detected")
                        except sr.RequestError as e:
                            log_message(f"  ⚠️ Chunk {chunk_num}: API error ({e})")
                            # Continue with other chunks even if one fails
            
            # Write the combined transcript
            if full_transcript:
                with open(txt_path, "w", encoding="utf-8") as f:
                    f.write(" ".join(full_transcript))
                log_message(f"✅ Transcribed: {filename}")
            else:
                log_message(f"⚠️ No transcript generated for {filename}")
                
        except sr.RequestError as e:
            error_msg = str(e).lower()
            if "bad request" in error_msg:
                log_message(f"❌ {filename}: Audio too long or incompatible format. Try a shorter file.")
            else:
                log_message(f"❌ {filename}: API error ({e})")
        except sr.UnknownValueError:
            log_message(f"❌ {filename}: No speech detected or audio unintelligible")
        except Exception as e:
            log_message(f"❌ {filename}: Unexpected error ({e})")
        finally:
            if os.path.exists(converted_path):
                os.remove(converted_path)

        progress_bar["value"] = idx + 1
        root.update_idletasks()

def browse_folder():
    folder_selected = filedialog.askdirectory()
    folder_entry.delete(0, tk.END)
    folder_entry.insert(0, folder_selected)

def start_transcription():
    log_text.config(state=tk.NORMAL)
    log_text.delete(1.0, tk.END)
    log_text.config(state=tk.DISABLED)

    if not check_ffmpeg_installed():
        show_ffmpeg_install_instructions()
        return

    folder = folder_entry.get().strip()
    if not folder:
        folder = os.path.join(os.getcwd(), "Audio")
        messagebox.showinfo("Default Folder", f"No folder selected. Using default: {folder}")

    if not os.path.exists(folder):
        messagebox.showerror("Folder Not Found", f"The folder '{folder}' does not exist.")
        return

    log_message(f"🔍 Starting transcription in folder: {folder}")
    transcribe_audio_files(folder)
    log_message("✅ Transcription completed.")

# GUI setup
root = tk.Tk()
root.title("Audio to Text Converter")

tk.Label(root, text="Select Input Folder (Audio folder is the default):").pack(pady=5)
folder_entry = tk.Entry(root, width=50)
folder_entry.pack(padx=10)

tk.Button(root, text="Browse", command=browse_folder).pack(pady=5)
tk.Button(root, text="Start Transcription", command=start_transcription).pack(pady=10)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.pack(pady=10)

log_frame = tk.Frame(root)
log_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

log_text = tk.Text(log_frame, height=12, wrap=tk.WORD, state=tk.DISABLED)
log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = ttk.Scrollbar(log_frame, command=log_text.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
log_text.config(yscrollcommand=scrollbar.set)

root.mainloop()