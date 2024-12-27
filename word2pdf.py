import os
import zipfile
from io import BytesIO
from PIL import Image, ImageOps
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading


def extract_images_from_word(input_path):
    """
    Extracts images from a Word document (.docx) and returns a list of image objects.

    Args:
        input_path (str): Path to the input Word (.docx) file.

    Returns:
        List of Pillow Image objects.
    """
    images = []
    file_names = []
    file_mapping = {}

    try:
        # Open the Word document as a zip file
        with zipfile.ZipFile(input_path, 'r') as docx_zip:
            # Look for image files in the 'word/media' folder
            for file in docx_zip.namelist():
                if file.startswith('word/media/'):
                    file_id = int(file.split(
                        '/')[-1].split('.')[0].replace('image', ''))
                    file_mapping[file_id] = file
                    file_names.append(file_id)
            file_names.sort()
            for file in file_names:
                file = file_mapping[file]
                # Extract the image data from the zip file
                image_data = docx_zip.read(file)

                # Open the image using Pillow
                image = Image.open(BytesIO(image_data))
                images.append(image)
    except Exception as e:
        raise RuntimeError(
            f"Error extracting images from Word document: {str(e)}")

    return images


def browse_file():
    """
    Opens a file dialog to select a Word document.
    """
    file_path = filedialog.askopenfilename(
        filetypes=[("Word Documents", "*.docx")])
    if file_path:
        entry_file.set(file_path)


def convert_to_pdf_threaded():
    """
    Runs the convert_to_pdf process in a separate thread to keep the GUI responsive.
    """
    threading.Thread(target=convert_to_pdf).start()


def convert_to_pdf():
    """
    Converts the selected Word document to PDF with optimized progress updates.
    """
    input_file = entry_file.get()
    if not input_file:
        messagebox.showerror("Error", "Please select a Word document.")
        return

    # Get the output PDF path (same directory, same name)
    output_file = os.path.splitext(input_file)[0] + ".pdf"

    try:
        # Update progress bar to show activity
        progress_bar.set(0.0)
        button_convert.configure(state="disabled")

        # Remove existing PDF file if it exists
        if os.path.exists(output_file):
            os.remove(output_file)

        # Extract images (track progress here)
        images = extract_images_from_word(input_file)
        total_images = len(images)
        if not total_images:
            raise ValueError("No images found in the Word document.")

        first_image_size = images[0].size
        normalized_images = []

        for img in images:
            # Resize to match the first image's dimensions
            # resized_img = img.resize(first_image_size, Image.LANCZOS)
            resized_img = ImageOps.fit(img, first_image_size, method=Image.LANCZOS)
            normalized_images.append(resized_img)

        # Save the PDF and update progress
        for idx, img in enumerate(normalized_images):
            if idx == 0:
                normalized_images[0].save(
                    output_file, save_all=True, append_images=normalized_images[1:], resolution=330.0, quality=100
                )
            progress_bar.set((idx + 1) / total_images)

        # # Save the PDF and update progress
        # for idx, img in enumerate(images):
        #     if idx == 0:
        #         img.save(output_file, save_all=True, append_images=images[1:], resolution=100.0, quality=100)
        #     progress_bar.set((idx + 1) / total_images)

        messagebox.showinfo(
            "Success", f"Successfully converted to {output_file}")
    except Exception as e:
        messagebox.showerror("Conversion Error", f"Error: {str(e)}")
    finally:
        button_convert.configure(state="normal")
        progress_bar.set(0.0)


if __name__ == "__main__":
    # Set up the main window
    app = ctk.CTk()
    app.title("High Quality Word to PDF Converter")
    app.geometry("440x200")
    app.resizable(False, False)

    # Input field to show selected file
    entry_file = ctk.StringVar()

    frame = ctk.CTkFrame(app)
    frame.pack(pady=20)

    # File selection entry and button
    entry = ctk.CTkEntry(frame, width=250, textvariable=entry_file,
                         placeholder_text="Select a .docx file")
    entry.grid(row=0, column=0, padx=10)
    button_browse = ctk.CTkButton(frame, text="Browse", command=browse_file)
    button_browse.grid(row=0, column=1)

    # Convert button
    button_convert = ctk.CTkButton(
        app, text="Convert to PDF", command=convert_to_pdf_threaded)
    button_convert.pack(pady=10)

    # Progress bar
    progress_bar = ctk.CTkProgressBar(app, width=300)
    progress_bar.pack(pady=10)
    progress_bar.set(0.0)

    app.mainloop()
