import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import fitz  # PyMuPDF
import cv2
import numpy as np
from pyzbar.pyzbar import decode
import win32com.client
from PIL import Image

# Configure appearance
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AutoArchiverApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AutoArchiver - Batch PDF QR Processor")
        self.geometry("800x600")

        # Configure grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Header
        self.header_frame = ctk.CTkFrame(self)
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        self.label_title = ctk.CTkLabel(self.header_frame, text="AutoArchiver: PDF -> Outlook", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_title.pack(padx=10, pady=10)

        # Controls
        self.controls_frame = ctk.CTkFrame(self)
        self.controls_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        self.btn_select_files = ctk.CTkButton(self.controls_frame, text="Select Files", command=self.select_files)
        self.btn_select_files.pack(side="left", padx=10, pady=10)

        self.btn_select_folder = ctk.CTkButton(self.controls_frame, text="Select Folder", command=self.select_folder)
        self.btn_select_folder.pack(side="left", padx=10, pady=10)

        self.btn_clear = ctk.CTkButton(self.controls_frame, text="Clear Log", fg_color="gray", command=self.clear_log)
        self.btn_clear.pack(side="right", padx=10, pady=10)

        # Log/Status Area
        self.log_textbox = ctk.CTkTextbox(self, width=760)
        self.log_textbox.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="nsew")
        
        # Processing State
        self.is_processing = False

    def log(self, message):
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")

    def clear_log(self):
        self.log_textbox.delete("0.0", "end")

    def select_files(self):
        if self.is_processing:
            return
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if file_paths:
            self.start_processing(file_paths)

    def select_folder(self):
        if self.is_processing:
            return
        folder_path = filedialog.askdirectory()
        if folder_path:
            files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
            if files:
                self.start_processing(files)
            else:
                self.log("No PDF files found in selected folder.")

    def start_processing(self, file_paths):
        self.is_processing = True
        self.log(f"--- Starting Batch of {len(file_paths)} files ---")
        
        # Run in thread to keep UI responsive
        thread = threading.Thread(target=self.process_files, args=(file_paths,))
        thread.start()

    def process_files(self, file_paths):
        outlook = None
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            self.log_textbox.after(0, self.log, f"ERROR: Could not connect to Outlook. {e}")
            self.is_processing = False
            return

        for file_path in file_paths:
            filename = os.path.basename(file_path)
            self.log_textbox.after(0, self.log, f"Processing: {filename}...")
            
            try:
                qr_text = self.extract_qr_from_pdf(file_path)
                
                if qr_text:
                    self.log_textbox.after(0, self.log, f"  > Found QR: {qr_text}")
                    self.create_outlook_mail(outlook, file_path, qr_text)
                    self.log_textbox.after(0, self.log, f"  > Draft Email Created.")
                else:
                    self.log_textbox.after(0, self.log, f"  > WARNING: No QR code found in {filename}")

            except Exception as e:
                self.log_textbox.after(0, self.log, f"  > ERROR: {str(e)}")

        self.is_processing = False
        self.log_textbox.after(0, self.log, "--- Batch Completed ---")

    def extract_qr_from_pdf(self, pdf_path):
        """
        Converts the first page of the PDF to an image and scans for QR codes.
        If not found on first page, tries others? For now, let's try all pages but stop at first find.
        """
        try:
            doc = fitz.open(pdf_path)
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                pix = page.get_pixmap(dpi=300) # High DPI for better QR detection
                
                # Convert to image format for opencv
                img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                
                # If RGB, openCV uses BGR
                if pix.n == 3:
                    img_data = cv2.cvtColor(img_data, cv2.COLOR_RGB2BGR)
                elif pix.n == 4: # RGBA
                    img_data = cv2.cvtColor(img_data, cv2.COLOR_RGBA2BGR)
                
                # Detect QR
                decoded_objects = decode(img_data)
                
                for obj in decoded_objects:
                    return obj.data.decode("utf-8")
                    
            return None
        except Exception as e:
            print(f"Error reading PDF {pdf_path}: {e}")
            raise e

    def create_outlook_mail(self, outlook, attachment_path, subject_text):
        try:
            mail = outlook.CreateItem(0) # 0 = olMailItem
            mail.Subject = subject_text
            mail.To = "sistematizacion@sparkgy.com"
            mail.Body = f"Adjunto el documento escaneado: {os.path.basename(attachment_path)}\n\nCodificado: {subject_text}"
            mail.Attachments.Add(attachment_path)
            mail.Save() # Save to Drafts
            # mail.Display() # Optional: Open the window
        except Exception as e:
            raise Exception(f"Failed to create Outlook Item: {e}")

if __name__ == "__main__":
    app = AutoArchiverApp()
    app.mainloop()
