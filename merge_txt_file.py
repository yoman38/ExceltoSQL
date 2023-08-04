import tkinter as tk
from tkinter import filedialog
from pathlib import Path

# Create a root Tkinter window and hide it (we only want the file dialog)
root = tk.Tk()
root.withdraw()

# Open a file dialog and allow the user to select multiple .txt files
file_paths = filedialog.askopenfilenames(
    title="Select files", 
    filetypes=(("Text files", "*.txt"), ("all files", "*.*")),
)

# Open the output file
with open('merged.txt', 'w') as output_file:
    # For each selected file...
    for file_path in file_paths:
        # Open the file and append its contents to the output file
        with open(file_path, 'r') as input_file:
            output_file.write(input_file.read())
            
print(f"Merged {len(file_paths)} files into merged.txt")

# Close the Tkinter window
root.destroy()
