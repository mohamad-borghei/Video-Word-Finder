import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from mp4_to_mp3.mp4tomp3 import extract_audio_from_video
from processing.searchinmovie import process_audio_file, save_results
from preprocessing.preprocessing import clean_word, keep_only_numbers
from processing.WordPatternMatching import process_multiple_words


# Function to select video file path and set it in the related entry field
def select_video_path():
    # Open file dialog to choose the video file and set its path in video_path
    video_path.set(filedialog.askopenfilename(filetypes=[("Video Files", "*.mp4")]))


# Function to select output folder and set it in the related entry field
def select_output_folder():
    # Open file dialog to choose the output folder and set its path in output_folder
    output_folder.set(filedialog.askdirectory())


# Main function to execute various processes on the video file
def run_main():
    # Retrieve input values from the GUI
    video_file = video_path.get()
    target_keyword = keyword_entry.get()
    chunk_duration = chunk_entry.get()
    output_dir = output_folder.get()

    # Check that all inputs are provided
    if not video_file or not target_keyword or not chunk_duration or not output_dir:
        # Show a warning message if inputs are incomplete
        messagebox.showwarning("Incomplete Input", "Please enter all inputs.")
        return

    progress_bar["value"] = 0  # Set initial value for the progress bar
    root.update_idletasks()  # Update the GUI to show the progress bar

    # Clean and standardize the keyword and chunk duration
    target_keyword = clean_word(target_keyword)
    chunk_duration = int(keep_only_numbers(chunk_duration))

    # Display processed values in the GUI
    target_keyword_label.config(text=f"Final Keyword: {target_keyword}")
    chunk_duration_label.config(text=f"Chunk Duration: {chunk_duration} seconds")

    # Update progress bar to 20%
    progress_bar["value"] = 20
    root.update_idletasks()

    # Extract audio from the video and save the path in audio_path
    audio_path = extract_audio_from_video(video_file, output_dir)
    progress_bar["value"] = 50  # Update progress bar to 50%
    root.update_idletasks()

    # Process audio file and extract words and sentences
    words, sentences = process_audio_file(audio_path, chunk_duration=chunk_duration)
    progress_bar["value"] = 80  # Update progress bar to 80%
    root.update_idletasks()

    # Save processed results in Excel files
    output_excel_words = os.path.join(output_dir, "words.xlsx")
    output_excel_sentences = os.path.join(output_dir, "sentences.xlsx")
    save_results(words, sentences, output_excel_words, output_excel_sentences, keyword=target_keyword)

    # Set progress bar to 100% and update GUI
    progress_bar["value"] = 100
    root.update_idletasks()

    # Display success message after processing is complete
    messagebox.showinfo("Successful Operation", "Processing and saving results completed successfully.")

    # Call function to process similar words
    show_similar_words(target_keyword, output_excel_words, output_dir)


# Function to display a window to ask about matching similar words
def show_similar_words(target_keyword, output_excel_words, output_folder):
    # Internal function to handle the user’s choice on matching similar words
    def on_similar_words_decision():
        # Get user’s choice from the similar_words_var variable
        matching = similar_words_var.get()
        if matching == 'y':  # If user confirms
            input_words = [target_keyword]  # Set the list of words to process
            # Process similar words and display success message
            process_multiple_words(input_words, output_excel_words, output_folder)
            messagebox.showinfo("Similar Words Matching", "Results for similar words processed successfully.")
        # Close the window after user’s choice
        similar_words_window.destroy()

    # Create a new window for selecting similar words matching
    similar_words_window = tk.Toplevel(root)
    similar_words_window.title("Similar Words Matching")

    # Add widgets for similar words matching question and confirmation options
    tk.Label(similar_words_window, text="Do you want to view results for similar words as well?").pack()
    similar_words_var = tk.StringVar(value="n")  # Variable to store user’s choice
    tk.Radiobutton(similar_words_window, text="Yes", variable=similar_words_var, value="y").pack()
    tk.Radiobutton(similar_words_window, text="No", variable=similar_words_var, value="n").pack()
    tk.Button(similar_words_window, text="Confirm", command=on_similar_words_decision).pack()


# Create the main window
root = tk.Tk()
root.title("Video Processing")

# Create a Notebook to add different tabs
notebook = ttk.Notebook(root)
notebook.grid(row=0, column=0, sticky="nsew")

# Create two frames for different tabs: Processing and Creator Info
main_frame = ttk.Frame(notebook)
info_frame = ttk.Frame(notebook)

# Add tabs to the Notebook
notebook.add(main_frame, text="Video Processing")
notebook.add(info_frame, text="About Creator")

# Video Processing frame
video_path = tk.StringVar()  # Variable to store the video file path
output_folder = tk.StringVar()  # Variable to store the output folder path

# Widgets for the Video Processing frame
tk.Label(main_frame, text="Video File Path:").grid(row=0, column=0, sticky="w")
tk.Entry(main_frame, textvariable=video_path, width=50).grid(row=0, column=1)
tk.Button(main_frame, text="Select File", command=select_video_path).grid(row=0, column=2)

# Widget for selecting output folder
tk.Label(main_frame, text="Output Folder Path:").grid(row=1, column=0, sticky="w")
tk.Entry(main_frame, textvariable=output_folder, width=50).grid(row=1, column=1)
tk.Button(main_frame, text="Select Folder", command=select_output_folder).grid(row=1, column=2)

# Widget for entering the keyword
tk.Label(main_frame, text="Keyword:").grid(row=2, column=0, sticky="w")
keyword_entry = tk.Entry(main_frame, width=50)
keyword_entry.grid(row=2, column=1, columnspan=2)

# Widget for entering the chunk duration for processing
tk.Label(main_frame, text="Chunk Duration (seconds):").grid(row=3, column=0, sticky="w")
chunk_entry = tk.Entry(main_frame, width=50)
chunk_entry.grid(row=3, column=1, columnspan=2)

# Button to start processing
tk.Button(main_frame, text="Start Processing", command=run_main).grid(row=4, column=1, pady=10)

# Progress bar for showing processing progress
progress_bar = ttk.Progressbar(main_frame, orient="horizontal", length=400, mode="determinate")
progress_bar.grid(row=5, column=0, columnspan=3, pady=10)

# Labels to display the processed values of the keyword and chunk duration
target_keyword_label = tk.Label(main_frame, text="Final Keyword: ")
target_keyword_label.grid(row=6, column=0, columnspan=3, sticky="w")

chunk_duration_label = tk.Label(main_frame, text="Chunk Duration: ")
chunk_duration_label.grid(row=7, column=0, columnspan=3, sticky="w")

# Creator Info frame
# Display creator info including name, email, and Telegram ID
tk.Label(info_frame, text="Creator: Mohammad Javad Borghei").pack(pady=5)
tk.Label(info_frame, text="Email: mohamad.b76@gmail.com").pack(pady=5)
tk.Label(info_frame, text="Telegram ID: https://t.me/Mohamad_borghei").pack(pady=5)

# Run the main window
root.mainloop()
