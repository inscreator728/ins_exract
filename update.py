import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import easyocr
import os
import threading

# Global variable for stopping the process
stop_process = False

def extract_text_from_image(image_path):
    """
    Extract text from an image file using EasyOCR.

    Args:
        image_path (str): The file path to the image.

    Returns:
        str: The extracted text.
    """
    try:
        reader = easyocr.Reader(["en"], gpu=False)
        results = reader.readtext(image_path)
        text = "\n".join([result[1] for result in results])
        return text
    except Exception as e:
        return f"Error: {str(e)}"

def process_bulk_images(folder_path):
    """
    Process all images in the selected folder, extract text, and save to a single file.

    Args:
        folder_path (str): The folder containing images.
    """
    global stop_process

    if not os.path.isdir(folder_path):
        messagebox.showerror("Error", "Selected folder is invalid.")
        return

    files = [f for f in os.listdir(folder_path) if f.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff"))]
    total_files = len(files)
    if total_files == 0:
        messagebox.showinfo("No Images Found", "No valid image files found in the selected folder.")
        return

    progress_bar["maximum"] = total_files
    progress_label.config(text=f"Progress: 0/{total_files} (0%)")

    output_file_path = os.path.join(folder_path, "extracted_texts.txt")

    with open(output_file_path, "w", encoding="utf-8") as output_file:
        for i, file_name in enumerate(files):
            if stop_process:
                messagebox.showinfo("Process Stopped", "The extraction process was canceled.")
                progress_bar["value"] = 0
                progress_label.config(text="Progress: Canceled")
                return

            file_path = os.path.join(folder_path, file_name)
            extracted_text = extract_text_from_image(file_path)

            # Save text to the output file
            output_file.write(f"{file_name}:\n{extracted_text}\n{'-'*40}\n")

            # Update progress bar and preview image
            progress_bar["value"] = i + 1
            percentage = ((i + 1) / total_files) * 100
            progress_label.config(text=f"Progress: {i + 1}/{total_files} ({percentage:.2f}%)")

            img = Image.open(file_path)
            img.thumbnail((250, 250))
            img_tk = ImageTk.PhotoImage(img)
            image_label.config(image=img_tk)
            image_label.image = img_tk

            root.update_idletasks()

    messagebox.showinfo("Success", f"All texts have been extracted and saved to {output_file_path}")
    os.startfile(folder_path)  # Open the folder where the file is saved

def open_folder():
    global stop_process
    stop_process = False
    folder_path = filedialog.askdirectory(title="Select Folder")
    if folder_path:
        threading.Thread(target=process_bulk_images, args=(folder_path,), daemon=True).start()

def cancel_process():
    global stop_process
    stop_process = True

def open_single_file():
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff")])
    if file_path:
        text = extract_text_from_image(file_path)
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, text)

        # Display the selected image
        img = Image.open(file_path)
        img.thumbnail((250, 250))
        img_tk = ImageTk.PhotoImage(img)
        image_label.config(image=img_tk)
        image_label.image = img_tk

def copy_to_clipboard():
    extracted_text = result_text.get(1.0, tk.END).strip()
    if extracted_text:
        root.clipboard_clear()
        root.clipboard_append(extracted_text)
        root.update()
        messagebox.showinfo("Success", "Text copied to clipboard!")

# Initialize the GUI
root = tk.Tk()
root.title("Bulk Image to Text Extractor")
root.geometry("600x800")

# Buttons for single file and bulk processing
single_file_button = tk.Button(root, text="Extract Single Image", command=open_single_file, font=("Arial", 12), bg="#4CAF50", fg="white")
single_file_button.pack(pady=10)

bulk_folder_button = tk.Button(root, text="Extract Bulk Images", command=open_folder, font=("Arial", 12), bg="#FF5722", fg="white")
bulk_folder_button.pack(pady=10)

cancel_button = tk.Button(root, text="Cancel Process", command=cancel_process, font=("Arial", 12), bg="#F44336", fg="white")
cancel_button.pack(pady=10)

# Progress bar and label
progress_bar = ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate")
progress_bar.pack(pady=10)

progress_label = tk.Label(root, text="Progress: 0/0 (0%)", font=("Arial", 10))
progress_label.pack(pady=5)

# Label to display the image
image_label = tk.Label(root)
image_label.pack(pady=10)

# Text widget to display the extracted text
result_text = tk.Text(root, wrap=tk.WORD, height=10, width=70)
result_text.pack(pady=10)

# Button to copy text to clipboard
copy_button = tk.Button(root, text="Copy to Clipboard", command=copy_to_clipboard, font=("Arial", 12), bg="#008CBA", fg="white")
copy_button.pack(pady=10)

# Start the GUI loop
root.mainloop()
