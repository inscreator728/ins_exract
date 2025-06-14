import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import easyocr
import os
from fpdf import FPDF
import threading

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

def save_text_as_filename(image_path, folder_path):
    """
    Extract text from an image and save it as a file named after the text.

    Args:
        image_path (str): Path to the image file.
        folder_path (str): Path to the folder to save the file.
    """
    extracted_text = extract_text_from_image(image_path)
    if extracted_text.strip():
        # Create a valid filename from the extracted text
        filename = extracted_text.split('\n')[0][:50].replace(" ", "_").replace("/", "_") + ".txt"
        save_path = os.path.join(folder_path, filename)
        try:
            with open(save_path, "w", encoding="utf-8") as file:
                file.write(extracted_text)
            messagebox.showinfo("Success", f"File saved as: {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file: {str(e)}")
    else:
        messagebox.showwarning("No Text Found", "No text detected in the image to save.")

def save_text_with_preview(image_path, folder_path):
    """
    Save the extracted text from an image to a file, with a preview of the text.

    Args:
        image_path (str): Path to the image file.
        folder_path (str): Path to the folder to save the file.
    """
    extracted_text = extract_text_from_image(image_path)
    if extracted_text.strip():
        filename = extracted_text.split('\n')[0][:50].replace(" ", "_").replace("/", "_") + ".txt"
        save_path = os.path.join(folder_path, filename)
        with open(save_path, "w", encoding="utf-8") as file:
            file.write(extracted_text)
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, extracted_text)
        messagebox.showinfo("Success", f"Text saved as: {filename}")
    else:
        messagebox.showwarning("No Text Found", "No text detected in the image.")

def process_save_text_as_filename():
    folder_path = filedialog.askdirectory(title="Select Save Folder")
    if folder_path:
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff")])
        if file_path:
            save_text_with_preview(file_path, folder_path)

def process_bulk_images(folder_path, save_as_pdf=False):
    """
    Process all images in the selected folder, extract text, and save to files.

    Args:
        folder_path (str): The folder containing images.
        save_as_pdf (bool): Whether to save the extracted texts as a PDF.
    """
    if not os.path.isdir(folder_path):
        messagebox.showerror("Error", "Selected folder is invalid.")
        return

    files = [f for f in os.listdir(folder_path) if f.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff"))]
    total_files = len(files)
    if total_files == 0:
        messagebox.showinfo("No Images Found", "No valid image files found in the selected folder.")
        return

    progress_bar["maximum"] = total_files

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    for i, file_name in enumerate(files):
        file_path = os.path.join(folder_path, file_name)
        extracted_text = extract_text_from_image(file_path)

        # Save text as .txt file
        text_file_path = os.path.join(folder_path, f"{os.path.splitext(file_name)[0]}.txt")
        with open(text_file_path, "w", encoding="utf-8") as text_file:
            text_file.write(extracted_text)

        if save_as_pdf:
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, f"{file_name}:\n\n{extracted_text}")

        # Update progress bar
        progress_bar["value"] = i + 1
        root.update_idletasks()

    if save_as_pdf:
        pdf_file_path = os.path.join(folder_path, "extracted_texts.pdf")
        pdf.output(pdf_file_path)
        messagebox.showinfo("Success", f"Texts saved as PDF at {pdf_file_path}")
    else:
        messagebox.showinfo("Success", "Texts extracted and saved as .txt files.")

def open_file_or_folder():
    path = filedialog.askdirectory(title="Select Folder")
    if path:
        save_as_pdf = messagebox.askyesno("Save as PDF", "Do you want to save the extracted texts as a PDF?")
        threading.Thread(target=process_bulk_images, args=(path, save_as_pdf), daemon=True).start()
        add_to_history(path)

def open_single_file():
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff")])
    if file_path:
        text = extract_text_from_image(file_path)
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, text)

        # Display the selected image
        img = Image.open(file_path)
        img.thumbnail((250, 250))  # Resize for display
        img_tk = ImageTk.PhotoImage(img)
        image_label.config(image=img_tk)
        image_label.image = img_tk
        add_to_history(file_path)

def copy_to_clipboard():
    extracted_text = result_text.get(1.0, tk.END).strip()
    if extracted_text:
        root.clipboard_clear()
        root.clipboard_append(extracted_text)
        root.update()  # Now it stays in the clipboard
        messagebox.showinfo("Success", "Text copied to clipboard!")

def add_to_history(path):
    history_list.insert(0, path)
    if history_list.size() > 10:
        history_list.delete(10, tk.END)

def view_history():
    selected = history_list.get(tk.ACTIVE)
    if selected:
        messagebox.showinfo("History Item", f"Selected Path: {selected}")

# Initialize the GUI
root = tk.Tk()
root.title("Enhanced Bulk Image to Text Extractor")
root.geometry("500x750")

# Buttons for single file and bulk processing
single_file_button = tk.Button(root, text="Extract Single Image", command=open_single_file, font=("Arial", 12), bg="#4CAF50", fg="white")
single_file_button.pack(pady=10)

bulk_folder_button = tk.Button(root, text="Extract Bulk Images", command=open_file_or_folder, font=("Arial", 12), bg="#FF5722", fg="white")
bulk_folder_button.pack(pady=10)

save_as_filename_button = tk.Button(root, text="Save Text as Filename", command=process_save_text_as_filename, font=("Arial", 12), bg="#FFC107", fg="black")
save_as_filename_button.pack(pady=10)

# Progress bar for bulk processing
progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.pack(pady=10)

# Label to display the image
image_label = tk.Label(root)
image_label.pack(pady=10)

# Text widget to display the extracted text
result_text = tk.Text(root, wrap=tk.WORD, height=10, width=50)
result_text.pack(pady=10)

# Button to copy text to clipboard
copy_button = tk.Button(root, text="Copy to Clipboard", command=copy_to_clipboard, font=("Arial", 12), bg="#008CBA", fg="white")
copy_button.pack(pady=10)

# History section
history_frame = tk.Frame(root)
history_frame.pack(pady=10)

history_label = tk.Label(history_frame, text="History (Last 10 Paths):", font=("Arial", 10))
history_label.pack(anchor="w")

history_list = tk.Listbox(history_frame, height=10, width=60)
history_list.pack(pady=5)

view_history_button = tk.Button(history_frame, text="View Selected Path", command=view_history, font=("Arial", 10), bg="#9C27B0", fg="white")
view_history_button.pack(pady=5)

# Start the GUI loop
root.mainloop()
