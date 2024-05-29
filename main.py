import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
from pptx import Presentation
import openai
from pydub import AudioSegment
from pathlib import Path
import pygame

# Set up OpenAI API
openai.api_key = 'my-api-key'

# Initialize pygame mixer
pygame.mixer.init()

# Ensure 'audios' directory exists
os.makedirs('audios', exist_ok=True)

def extract_text_from_pdf(file_path):
    with pdfplumber.open(file_path) as pdf:
        text = '\n'.join([page.extract_text() for page in pdf.pages if page.extract_text() is not None])
    return text

def extract_text_from_pptx(file_path):
    prs = Presentation(file_path)
    text = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text.append(shape.text)
    return '\n'.join(text)

def text_to_speech(text, filename):
    response = openai.audio.speech.create(
        model="tts-1",
        voice="alloy",
        input="WHat the sigma",
    )
    audio_file_path = Path('audios') / f"{filename}.mp3"
    response.write_to_file(audio_file_path)
    return audio_file_path

def handle_file_upload():
    file_path = filedialog.askopenfilename(
        filetypes=[("PDF files", "*.pdf"), ("PPTX files", "*.pptx")]
    )
    if not file_path:
        return

    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == ".pdf":
        text = extract_text_from_pdf(file_path)
    elif file_extension == ".pptx":
        text = extract_text_from_pptx(file_path)
    else:
        messagebox.showerror("Error", "Unsupported file type")
        return

    if not text:
        messagebox.showerror("Error", "No text found in file")
        return

    filename = os.path.splitext(os.path.basename(file_path))[0]
    audio_file_path = text_to_speech(text, filename)
    messagebox.showinfo("Success", f"Audio file created: {audio_file_path}")
    refresh_audio_list()

def refresh_audio_list():
    listbox.delete(0, tk.END)
    for filename in os.listdir('audios'):
        listbox.insert(tk.END, filename)

def play_audio():
    selected_file = listbox.get(tk.ACTIVE)
    if selected_file:
        audio_file_path = os.path.join('audios', selected_file)
        pygame.mixer.music.load(audio_file_path)
        pygame.mixer.music.play()

def delete_audio():
    selected_file = listbox.get(tk.ACTIVE)
    if selected_file:
        audio_file_path = os.path.join('audios', selected_file)
        os.remove(audio_file_path)
        refresh_audio_list()

# Set up the Tkinter GUI
root = tk.Tk()
root.title("PDF/PPTX to Audio Converter")

upload_button = tk.Button(root, text="Upload PDF/PPTX", command=handle_file_upload)
upload_button.pack(pady=10)

listbox = tk.Listbox(root)
listbox.pack(fill=tk.BOTH, expand=True)

refresh_button = tk.Button(root, text="Refresh List", command=refresh_audio_list)
refresh_button.pack(pady=5)

play_button = tk.Button(root, text="Play Audio", command=play_audio)
play_button.pack(pady=5)

delete_button = tk.Button(root, text="Delete Audio", command=delete_audio)
delete_button.pack(pady=5)

refresh_audio_list()

# Ensure pygame is properly cleaned up on exit
def on_closing():
    pygame.mixer.quit()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()
