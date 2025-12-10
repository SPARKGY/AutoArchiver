import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import fitz  # PyMuPDF
import cv2
import numpy as np
import win32com.client
from PIL import Image

# Set theme and color
ctk.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class AutoArchiverApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AutoArchiver - Spark Energy")
        self.geometry("900x600")

        # Set Icon
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sparkgy.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)

        # Grid Layout
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Sidebar (Left) ---
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="AutoArchiver", font=ctk.CTkFont(size=24, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        self.sidebar_desc = ctk.CTkLabel(self.sidebar_frame, text="PDF QR Scanner\n& Outlook Drafts", font=ctk.CTkFont(size=12))
        self.sidebar_desc.grid(row=1, column=0, padx=20, pady=(0, 20))

        self.btn_select_files = ctk.CTkButton(self.sidebar_frame, text="Select Files", command=self.select_files,
                                              fg_color="#106EBE", hover_color="#005A9E") # Outlook Blue-ish
        self.btn_select_files.grid(row=2, column=0, padx=20, pady=10)

        self.btn_select_folder = ctk.CTkButton(self.sidebar_frame, text="Select Folder", command=self.select_folder,
                                               fg_color="#106EBE", hover_color="#005A9E")
        self.btn_select_folder.grid(row=3, column=0, padx=20, pady=10)

        self.btn_clear = ctk.CTkButton(self.sidebar_frame, text="Clear Log", command=self.clear_log,
                                       fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"))
        self.btn_clear.grid(row=5, column=0, padx=20, pady=(10, 20))

        # --- Main Area (Right) ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_rowconfigure(2, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Status / Progress Bar
        self.status_label = ctk.CTkLabel(self.main_frame, text="Ready", anchor="w", font=ctk.CTkFont(size=14))
        self.status_label.grid(row=0, column=0, padx=0, pady=(0, 5), sticky="w")
        
        self.progressbar = ctk.CTkProgressBar(self.main_frame)
        self.progressbar.grid(row=1, column=0, padx=0, pady=(0, 20), sticky="ew")
        self.progressbar.set(0)

        # Log Area
        self.log_textbox = ctk.CTkTextbox(self.main_frame, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_textbox.grid(row=2, column=0, sticky="nsew")
        
        # Internal State
        self.is_processing = False

    def log(self, message):
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")

    def clear_log(self):
        self.log_textbox.delete("0.0", "end")
        self.progressbar.set(0)
        self.status_label.configure(text="Ready")

    def select_files(self):
        if self.is_processing: return
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if file_paths: self.start_processing(file_paths)

    def select_folder(self):
        if self.is_processing: return
        folder_path = filedialog.askdirectory()
        if folder_path:
            files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
            if files: self.start_processing(files)
            else: self.log("No PDF files found in selected folder.")

    def start_processing(self, file_paths):
        self.is_processing = True
        self.status_label.configure(text=f"Processing {len(file_paths)} files...")
        self.progressbar.set(0)
        self.log(f"\n--- Starting Batch of {len(file_paths)} files ---")
        
        thread = threading.Thread(target=self.process_files, args=(file_paths,))
        thread.start()

    def process_files(self, file_paths):
        outlook = None
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            self.update_ui_error(f"ERROR: Could not connect to Outlook. {e}")
            return

        total = len(file_paths)
        for i, file_path in enumerate(file_paths):
            filename = os.path.basename(file_path)
            self.update_log_safe(f"[{i+1}/{total}] Processing: {filename}...")
            
            try:
                qr_text = self.extract_qr_from_pdf(file_path)
                
                if qr_text:
                    self.update_log_safe(f"  > Found QR: {qr_text}")
                    self.create_outlook_mail(outlook, file_path, qr_text)
                    self.update_log_safe(f"  > Draft Email Created.")
                else:
                    self.update_log_safe(f"  > WARNING: No QR code found.")

            except Exception as e:
                self.update_log_safe(f"  > ERROR: {str(e)}")
            
            # Update Progress
            self.update_progress_safe((i + 1) / total)

        self.update_ui_done()

    def extract_qr_from_pdf(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
            pages_to_check = min(3, len(doc))
            detector = cv2.QRCodeDetector()

            for page_num in range(pages_to_check):
                page = doc.load_page(page_num)
                pix = page.get_pixmap(dpi=300) 
                
                # Convert img for opencv
                img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                
                if pix.n == 3: img_data = cv2.cvtColor(img_data, cv2.COLOR_RGB2BGR)
                elif pix.n == 4: img_data = cv2.cvtColor(img_data, cv2.COLOR_RGBA2BGR)
                else: img_data = cv2.cvtColor(img_data, cv2.COLOR_GRAY2BGR)

                data, bbox, _ = detector.detectAndDecode(img_data)
                if data: return data
            return None
        except Exception as e:
            raise e

    def create_outlook_mail(self, outlook, attachment_path, subject_text):
        mail = outlook.CreateItem(0) 
        mail.Subject = subject_text
        mail.To = "sistematizacion@sparkgy.com"
        mail.Body = f"Adjunto el documento escaneado: {os.path.basename(attachment_path)}\n\nCodificado: {subject_text}"
        mail.Attachments.Add(attachment_path)
        mail.Save() 

    # --- Thread Safe UI Updates ---
    def update_log_safe(self, message):
        self.log_textbox.after(0, self.log, message)

    def update_progress_safe(self, val):
        self.progressbar.after(0, self.progressbar.set, val)

    def update_ui_error(self, msg):
        self.log_textbox.after(0, self.log, msg)
        self.is_processing = False
        self.status_label.after(0, lambda: self.status_label.configure(text="Error"))

    def update_ui_done(self):
        self.is_processing = False
        self.log_textbox.after(0, self.log, "\n--- Batch Completed ---")
        self.status_label.after(0, lambda: self.status_label.configure(text="Done"))

if __name__ == "__main__":
    app = AutoArchiverApp()
    app.mainloop()
