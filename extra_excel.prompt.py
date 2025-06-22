import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import easyocr
import os
import threading
import pandas as pd

# Global variable for stopping the process
stop_process = False

def extract_text_from_image(image_path, scan_mode):
    """
    Extract text from an image file using EasyOCR.

    Args:
        image_path (str): The file path to the image.
        scan_mode (str): Scan mode - normal, super_scan, or intense_scan.

    Returns:
        str: The extracted text.
    """
    try:
        if scan_mode == "normal":
            reader = easyocr.Reader(["en"], gpu=False)
        elif scan_mode == "super_scan":
            reader = easyocr.Reader(["en"], gpu=True)
        elif scan_mode == "intense_scan":
            reader = easyocr.Reader(["en"], gpu=True, model_storage_directory="high_precision_model")

        results = reader.readtext(image_path)
        text = "\n".join([result[1] for result in results])
        return text
    except Exception as e:
        return f"Error: {str(e)}"

def process_bulk_images(folder_path, save_option, scan_mode):
    """
    Process all images in the selected folder, extract text, and save based on the selected save option.

    Args:
        folder_path (str): The folder containing images.
        save_option (str): Save option selected by the user.
        scan_mode (str): Scan mode - normal, super_scan, or intense_scan.
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

    output_folder = os.path.join(folder_path, "Extracted_Texts")
    os.makedirs(output_folder, exist_ok=True)

    records = []  # For storing data for the Excel sheet

    for i, file_name in enumerate(files):
        if stop_process:
            messagebox.showinfo("Process Stopped", "The extraction process was canceled.")
            progress_bar["value"] = 0
            progress_label.config(text="Progress: Canceled")
            return

        file_path = os.path.join(folder_path, file_name)
        extracted_text = extract_text_from_image(file_path, scan_mode)

        # Save based on the selected option
        if save_option == "single_file":
            records.append({"File Name": file_name, "Extracted Text": extracted_text})
        elif save_option == "separate_files":
            subfolder = os.path.join(output_folder, os.path.splitext(file_name)[0])
            os.makedirs(subfolder, exist_ok=True)
            text_file_path = os.path.join(subfolder, "extracted_text.txt")
            with open(text_file_path, "w", encoding="utf-8") as text_file:
                text_file.write(extracted_text)
            records.append({"File Name": file_name, "Saved Path": text_file_path})
        elif save_option == "extracted_name":
            if extracted_text.strip():
                folder_name = extracted_text.split('\n')[0][:50].replace(" ", "_").replace("/", "_")
                folder_path = os.path.join(output_folder, folder_name)
                os.makedirs(folder_path, exist_ok=True)
                text_file_path = os.path.join(folder_path, "extracted_text.txt")
                with open(text_file_path, "w", encoding="utf-8") as text_file:
                    text_file.write(extracted_text)
                records.append({"Extracted Name": folder_name, "Saved Path": text_file_path})

        # Update progress bar and preview boxes
        progress_bar["value"] = i + 1
        percentage = ((i + 1) / total_files) * 100
        progress_label.config(text=f"Progress: {i + 1}/{total_files} ({percentage:.2f}%)")

        img = Image.open(file_path)
        img.thumbnail((250, 250))
        img_tk = ImageTk.PhotoImage(img)
        image_preview.config(image=img_tk)
        image_preview.image = img_tk

        extracted_text_preview.delete(1.0, tk.END)
        extracted_text_preview.insert(tk.END, extracted_text)

        root.update_idletasks()

    # Save records to an Excel file
    excel_file_path = os.path.join(output_folder, "extracted_texts.xlsx")
    df = pd.DataFrame(records)
    df.to_excel(excel_file_path, index=False)

    messagebox.showinfo("Success", f"All texts have been extracted and saved in {output_folder}")
    os.startfile(output_folder)  # Open the output folder

def start_bulk_processing():
    global stop_process
    stop_process = False
    folder_path = filedialog.askdirectory(title="Select Folder")
    if folder_path:
        save_option = save_option_var.get()
        scan_mode = scan_mode_var.get()
        threading.Thread(target=process_bulk_images, args=(folder_path, save_option, scan_mode), daemon=True).start()

def cancel_process():
    global stop_process
    stop_process = True

def open_single_file():
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff")])
    if file_path:
        scan_mode = scan_mode_var.get()
        text = extract_text_from_image(file_path, scan_mode)
        extracted_text_preview.delete(1.0, tk.END)
        extracted_text_preview.insert(tk.END, text)

        img = Image.open(file_path)
        img.thumbnail((250, 250))
        img_tk = ImageTk.PhotoImage(img)
        image_preview.config(image=img_tk)
        image_preview.image = img_tk

def copy_to_clipboard():
    extracted_text = extracted_text_preview.get(1.0, tk.END).strip()
    if extracted_text:
        root.clipboard_clear()
        root.clipboard_append(extracted_text)
        root.update()
        messagebox.showinfo("Success", "Text copied to clipboard!")

# Initialize the GUI
root = tk.Tk()
root.title("Enhanced Bulk Image to Text Extractor with Scan Modes")
root.geometry("900x700")

# Scan mode selector
scan_mode_var = tk.StringVar(value="normal")
scan_frame = tk.Frame(root)
scan_frame.pack(pady=10)

normal_scan_radio = tk.Radiobutton(scan_frame, text="Normal Scan", variable=scan_mode_var, value="normal")
normal_scan_radio.grid(row=0, column=0, padx=10)

super_scan_radio = tk.Radiobutton(scan_frame, text="Super Scan", variable=scan_mode_var, value="super_scan")
super_scan_radio.grid(row=0, column=1, padx=10)

intense_scan_radio = tk.Radiobutton(scan_frame, text="Intense Scan", variable=scan_mode_var, value="intense_scan")
intense_scan_radio.grid(row=0, column=2, padx=10)

# Save option selector
save_option_var = tk.StringVar(value="single_file")
options_frame = tk.Frame(root)
options_frame.pack(pady=10)

single_file_radio = tk.Radiobutton(options_frame, text="Save All in Single Excel Sheet", variable=save_option_var, value="single_file")
single_file_radio.grid(row=0, column=0, padx=10)

separate_files_radio = tk.Radiobutton(options_frame, text="Save Each in Separate Files", variable=save_option_var, value="separate_files")
separate_files_radio.grid(row=0, column=1, padx=10)

extracted_name_radio = tk.Radiobutton(options_frame, text="Save as Extracted Names in Excel", variable=save_option_var, value="extracted_name")
extracted_name_radio.grid(row=0, column=2, padx=10)

# Buttons for single file and bulk processing
single_file_button = tk.Button(root, text="Extract Single Image", command=open_single_file, font=("Arial", 12), bg="#4CAF50", fg="white")
single_file_button.pack(pady=10)

bulk_folder_button = tk.Button(root, text="Extract Bulk Images", command=start_bulk_processing, font=("Arial", 12), bg="#FF5722", fg="white")
bulk_folder_button.pack(pady=10)

cancel_button = tk.Button(root, text="Cancel Process", command=cancel_process, font=("Arial", 12), bg="#F44336", fg="white")
cancel_button.pack(pady=10)

# Progress bar and label
progress_bar = ttk.Progressbar(root, orient="horizontal", length=800, mode="determinate")
progress_bar.pack(pady=10)

progress_label = tk.Label(root, text="Progress: 0/0 (0%)", font=("Arial", 10))
progress_label.pack(pady=5)

# Preview frames
preview_frame = tk.Frame(root)
preview_frame.pack(pady=10, fill="both", expand=True)

# Left: Extracted text preview
extracted_text_preview = tk.Text(preview_frame, wrap=tk.WORD, height=20, width=40)
extracted_text_preview.grid(row=0, column=0, padx=10, sticky="nsew")

# Right: Image preview
image_preview = tk.Label(preview_frame)
image_preview.grid(row=0, column=1, padx=10, sticky="nsew")

# Configure grid weights for resizing
preview_frame.columnconfigure(0, weight=1)
preview_frame.columnconfigure(1, weight=1)

# Start the GUI loop
root.mainloop()