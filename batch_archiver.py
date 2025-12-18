import sys
import traceback

try:
    import os
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox
    import customtkinter as ctk
    import fitz  # PyMuPDF
    import cv2
    import numpy as np
    import zxingcpp
    import win32com.client
    from PIL import Image, ImageTk
    from tkinterdnd2 import TkinterDnD, DND_FILES
    import requests
    import json
    import base64
    import shutil
except Exception as e:
    traceback.print_exc()
    input("CRITICAL IMPORT ERROR: Press Enter to exit...")
    sys.exit(1)

# Set theme and color
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class DndCTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class MainMenuFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.center_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.center_frame.grid(row=0, column=0, sticky="nsew")
        self.center_frame.grid_columnconfigure(0, weight=1)
        self.center_frame.grid_rowconfigure((0,1,2,3), weight=1)

        self.logo_label = ctk.CTkLabel(self.center_frame, text="AutoArchiver", font=ctk.CTkFont(size=40, weight="bold"))
        self.logo_label.grid(row=1, column=0, padx=20, pady=(20, 10))
        
        self.desc_label = ctk.CTkLabel(self.center_frame, text="Select Mode", font=ctk.CTkFont(size=16))
        self.desc_label.grid(row=2, column=0, padx=20, pady=(0, 20))

        self.btn_scanner = ctk.CTkButton(self.center_frame, text="Escanear QR en PDFs", width=250, height=50,
                                         font=ctk.CTkFont(size=18, weight="bold"),
                                         command=lambda: controller.show_frame("PDFScannerFrame"))
        self.btn_scanner.grid(row=3, column=0, padx=20, pady=10)

        self.btn_img2pdf = ctk.CTkButton(self.center_frame, text="Convertir Imágenes a PDF", width=250, height=50,
                                         font=ctk.CTkFont(size=18, weight="bold"),
                                         command=lambda: controller.show_frame("ImageToPDFFrame"))
        self.btn_img2pdf.grid(row=4, column=0, padx=20, pady=10)


class PDFScannerFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        # State Variables
        self.file_queue = []
        self.is_processing = False
        self.stop_event = threading.Event()
        self.current_pdf_doc = None
        self.current_page_pix = None
        self.zoom_level = 1.0
        self.pan_offset_x = 0
        self.pan_offset_y = 0
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.current_page_index = 0
        self.total_pages = 0

        # Layout Configuration
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Sidebar (Left) ---
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(8, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="PDF Scanner", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        self.sidebar_desc = ctk.CTkLabel(self.sidebar_frame, text="QR & Email", font=ctk.CTkFont(size=12))
        self.sidebar_desc.grid(row=1, column=0, padx=20, pady=(0, 20))

        # Controls
        self.switch_email_mode = ctk.CTkSwitch(self.sidebar_frame, text="Use EmailJS", state="disabled")
        self.switch_email_mode.grid(row=2, column=0, padx=20, pady=10)

        self.btn_select_files = ctk.CTkButton(self.sidebar_frame, text="Add Files...", command=self.select_files)
        self.btn_select_files.grid(row=3, column=0, padx=20, pady=10)

        self.btn_start = ctk.CTkButton(self.sidebar_frame, text="START CHECK", command=self.start_processing,
                                       fg_color="#2CC985", hover_color="#229C68", text_color="white") # Green
        self.btn_start.grid(row=4, column=0, padx=20, pady=10)

        self.btn_stop = ctk.CTkButton(self.sidebar_frame, text="STOP", command=self.stop_processing,
                                      fg_color="#D94040", hover_color="#A83232", state="disabled") # Red
        self.btn_stop.grid(row=5, column=0, padx=20, pady=10)

        self.btn_clear = ctk.CTkButton(self.sidebar_frame, text="Clear Queue", command=self.clear_queue,
                                       fg_color="transparent", border_width=1, text_color=("gray10", "#DCE4EE"))
        self.btn_clear.grid(row=6, column=0, padx=20, pady=(20, 10))

        # Back Button
        self.btn_back = ctk.CTkButton(self.sidebar_frame, text="← Menu", command=lambda: controller.show_frame("MainMenuFrame"),
                                      fg_color="transparent", border_width=1, text_color=("gray10", "#DCE4EE"))
        self.btn_back.grid(row=7, column=0, padx=20, pady=(20, 10))

        # --- Main Area (Right) ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.main_frame.grid_rowconfigure(0, weight=1) # Upper split
        self.main_frame.grid_rowconfigure(1, weight=0) # Log/Status
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Split Container (List + Preview)
        self.split_frame = ctk.CTkFrame(self.main_frame)
        self.split_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        self.split_frame.grid_columnconfigure(0, weight=1) # List
        self.split_frame.grid_columnconfigure(1, weight=3) # Preview
        self.split_frame.grid_rowconfigure(0, weight=1)

        # File List (Left of Split)
        self.file_listbox_frame = ctk.CTkFrame(self.split_frame, width=200)
        self.file_listbox_frame.grid(row=0, column=0, sticky="nsew", padx=(0,5))
        
        self.file_list_label = ctk.CTkLabel(self.file_listbox_frame, text="File Queue", font=ctk.CTkFont(weight="bold"))
        self.file_list_label.pack(pady=5)
        
        self.file_listbox = tk.Listbox(self.file_listbox_frame, bg="#2B2B2B", fg="#DCE4EE", 
                                       selectbackground="#1F6AA5", selectforeground="white", borderwidth=0, highlightthickness=0)
        self.file_listbox.pack(expand=True, fill="both", padx=5, pady=5)
        self.file_listbox.bind("<<ListboxSelect>>", self.on_list_select)

        # Image Viewer (Right of Split)
        self.viewer_frame = ctk.CTkFrame(self.split_frame, fg_color="#1A1A1A")
        self.viewer_frame.grid(row=0, column=1, sticky="nsew")
        self.viewer_frame.pack_propagate(False)

        # -- Toolbar --
        self.toolbar_frame = ctk.CTkFrame(self.viewer_frame, height=40)
        self.toolbar_frame.pack(side="top", fill="x", padx=2, pady=2)
        
        # Toolbar Buttons
        self.btn_prev = ctk.CTkButton(self.toolbar_frame, text="<", width=30, command=self.prev_page)
        self.btn_prev.pack(side="left", padx=2)
        
        self.lbl_page = ctk.CTkLabel(self.toolbar_frame, text="0 / 0", width=50)
        self.lbl_page.pack(side="left", padx=2)
        
        self.btn_next = ctk.CTkButton(self.toolbar_frame, text=">", width=30, command=self.next_page)
        self.btn_next.pack(side="left", padx=2)
        
        ctk.CTkLabel(self.toolbar_frame, text="|", width=10).pack(side="left", padx=5)

        self.btn_zoom_out = ctk.CTkButton(self.toolbar_frame, text="-", width=30, command=lambda: self.adjust_zoom(0.9))
        self.btn_zoom_out.pack(side="left", padx=2)
        
        self.btn_zoom_in = ctk.CTkButton(self.toolbar_frame, text="+", width=30, command=lambda: self.adjust_zoom(1.1))
        self.btn_zoom_in.pack(side="left", padx=2)

        self.btn_fit_width = ctk.CTkButton(self.toolbar_frame, text="Fit Width", width=70, command=self.fit_width)
        self.btn_fit_width.pack(side="left", padx=2)

        self.btn_fit_view = ctk.CTkButton(self.toolbar_frame, text="Fit Page", width=70, command=self.fit_view)
        self.btn_fit_view.pack(side="left", padx=2)
        
        # Viewer Canvas
        self.canvas = tk.Canvas(self.viewer_frame, bg="#333333", highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Canvas Controls
        self.canvas.bind("<ButtonPress-2>", self.on_drag_start)
        self.canvas.bind("<B2-Motion>", self.on_drag_motion)
        self.canvas.bind("<MouseWheel>", self.on_zoom)

        # Log & Status (Bottom)
        self.bottom_frame = ctk.CTkFrame(self.main_frame, height=150)
        self.bottom_frame.grid(row=1, column=0, sticky="ew")
        self.bottom_frame.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(self.bottom_frame, text="Ready. Drag & Drop PDF files here.", anchor="w")
        self.status_label.grid(row=0, column=0, sticky="w", padx=10, pady=(5,0))
        
        self.progressbar = ctk.CTkProgressBar(self.bottom_frame)
        self.progressbar.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        self.progressbar.set(0)

        self.log_textbox = ctk.CTkTextbox(self.bottom_frame, height=100, font=ctk.CTkFont(family="Consolas", size=11))
        self.log_textbox.grid(row=2, column=0, sticky="ew", padx=10, pady=(0,10))

        # DnD Registration - handled by controller/parent usually, but here we can try binding to self
        # Since 'self' is a Frame inside the Root, we need to bind to the Root or ensure the frame accepts drops.
        # TkinterDnD widgets need to be registered.
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.drop)

    # --- File Management ---
    def log(self, message):
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")

    def update_log_safe(self, message):
        self.after(0, lambda: self.log(message))

    def update_progress_safe(self, value):
        self.after(0, lambda: self.progressbar.set(value))

    def drop(self, event):
        if self.is_processing: return
        raw_files = self.split_file_list(event.data)
        count = 0
        for f in raw_files:
            if f.lower().endswith(".pdf") and f not in self.file_queue:
                self.file_queue.append(f)
                count += 1
        if count > 0:
            self.log(f"Added {count} files via Drag & Drop.")
            self.update_file_list_ui()

    def split_file_list(self, data):
        if data.startswith('{'):
            import re
            parts = re.findall(r'\{.*?\}|\S+', data) 
            return [p.strip('{}') for p in parts]
        return data.split()

    def select_files(self):
        if self.is_processing: return
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if file_paths:
            count = 0
            for f in file_paths:
                if f not in self.file_queue:
                    self.file_queue.append(f)
                    count += 1
            if count > 0:
                self.log(f"Added {count} files.")
                self.update_file_list_ui()

    def clear_queue(self):
        if self.is_processing: return
        self.file_queue = []
        self.update_file_list_ui()
        self.clear_canvas()
        self.log("Queue cleared.")

    def update_file_list_ui(self):
        self.file_listbox.delete(0, "end")
        for f in self.file_queue:
            self.file_listbox.insert("end", os.path.basename(f))
        self.status_label.configure(text=f"Ready. {len(self.file_queue)} files in queue.")

    # --- Visualization ---
    def on_list_select(self, event):
        selection = self.file_listbox.curselection()
        if not selection: return
        index = selection[0]
        file_path = self.file_queue[index]
        self.load_pdf_preview_async(file_path)

    def load_pdf_preview_async(self, file_path):
        threading.Thread(target=self._load_pdf_thread, args=(file_path,), daemon=True).start()

    def _load_pdf_thread(self, file_path):
        try:
            doc = fitz.open(file_path)
            self.total_pages = len(doc)
            self.current_page_index = 0
            self._load_page_image(doc, 0)
        except Exception as e:
            self.update_log_safe(f"Error previewing file: {e}")

    def _load_page_image(self, doc, page_index):
        try:
            page = doc.load_page(page_index) 
            pix = page.get_pixmap(dpi=150)
            mode = "RGB" if pix.n >= 3 else "L"
            if pix.n == 4: mode = "RGBA"
            pil_image = Image.frombytes(mode, (pix.w, pix.h), pix.samples)
            if mode == "RGBA":
                pil_image = pil_image.convert("RGB")
            self.after(0, self._on_pdf_loaded, doc, pil_image)
        except Exception as e:
            print(f"Page load error: {e}")

    def _on_pdf_loaded(self, doc, pil_image):
        if self.current_pdf_doc and self.current_pdf_doc != doc:
             try: self.current_pdf_doc.close()
             except: pass

        self.current_pdf_doc = doc
        self.pil_image = pil_image
        self.lbl_page.configure(text=f"{self.current_page_index + 1} / {self.total_pages}")
        self.zoom_level = 1.0
        self.fit_view()

    def prev_page(self):
        if not self.current_pdf_doc or self.current_page_index <= 0: return
        self.current_page_index -= 1
        threading.Thread(target=self._reload_current_page, daemon=True).start()

    def next_page(self):
        if not self.current_pdf_doc or self.current_page_index >= self.total_pages - 1: return
        self.current_page_index += 1
        threading.Thread(target=self._reload_current_page, daemon=True).start()

    def _reload_current_page(self):
         self._load_page_image(self.current_pdf_doc, self.current_page_index)

    def fit_view(self):
        if not hasattr(self, 'pil_image'): return
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        img_w, img_h = self.pil_image.size
        scale_w = canvas_w / img_w
        scale_h = canvas_h / img_h
        self.zoom_level = min(scale_w, scale_h) * 0.9
        self.display_image(reset_view=True)

    def fit_width(self):
        if not hasattr(self, 'pil_image'): return
        canvas_w = self.canvas.winfo_width()
        img_w, _ = self.pil_image.size
        self.zoom_level = (canvas_w / img_w) * 0.95
        self.display_image(reset_view=False)
        new_w = int(img_w * self.zoom_level)
        self.pan_offset_x = (canvas_w - new_w) // 2
        self.pan_offset_y = 20
        self._update_canvas_coords()

    def adjust_zoom(self, factor):
        self.zoom_level *= factor
        self.display_image()

    def display_image(self, reset_view=False):
        if not hasattr(self, 'pil_image'): return
        threading.Thread(target=self._resize_image_thread, args=(self.zoom_level, reset_view), daemon=True).start()

    def _resize_image_thread(self, zoom, reset_view):
        try:
            width, height = self.pil_image.size
            new_w = int(width * zoom)
            new_h = int(height * zoom)
            resized = self.pil_image.resize((new_w, new_h), Image.Resampling.BILINEAR)
            tk_image = ImageTk.PhotoImage(resized)
            self.after(0, self._update_canvas, tk_image, reset_view, new_w, new_h)
        except Exception as e:
            print(f"Resize error: {e}")

    def _update_canvas(self, tk_image, reset_view, new_w, new_h):
        self.tk_image = tk_image 
        self.canvas.delete("all")
        if reset_view:
            canvas_w = self.canvas.winfo_width()
            canvas_h = self.canvas.winfo_height()
            self.pan_offset_x = (canvas_w - new_w) // 2
            self.pan_offset_y = (canvas_h - new_h) // 2
        self._update_canvas_coords()

    def _update_canvas_coords(self):
        self.canvas.delete("all")
        self.image_id = self.canvas.create_image(
            self.pan_offset_x, 
            self.pan_offset_y, 
            image=self.tk_image, 
            anchor="nw"
        )

    def clear_canvas(self):
        self.canvas.delete("all")
        self.current_pdf_doc = None
        self.lbl_page.configure(text="0 / 0")

    def on_drag_start(self, event):
        self.drag_start_x = event.x
        self.drag_start_y = event.y

    def on_drag_motion(self, event):
        dx = event.x - self.drag_start_x
        dy = event.y - self.drag_start_y
        self.pan_offset_x += dx
        self.pan_offset_y += dy
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self._update_canvas_coords()
        
    def on_zoom(self, event):
        if not hasattr(self, 'pil_image'): return
        old_zoom = self.zoom_level
        if event.delta > 0:
            zoom_factor = 1.1
        else:
            zoom_factor = 0.9
            
        new_zoom = old_zoom * zoom_factor
        new_zoom = max(0.1, min(10.0, new_zoom))
        
        if new_zoom == old_zoom: return
        self.zoom_level = new_zoom
        
        verify_x = event.x - self.pan_offset_x
        verify_y = event.y - self.pan_offset_y
        new_verify_x = verify_x * zoom_factor
        new_verify_y = verify_y * zoom_factor
        
        self.pan_offset_x = int(event.x - new_verify_x)
        self.pan_offset_y = int(event.y - new_verify_y)
        self.display_image(reset_view=False)

    # --- Processing ---
    def start_processing(self):
        if not self.file_queue:
            messagebox.showinfo("Info", "No files in queue to process.")
            return

        self.is_processing = True
        self.stop_event.clear()
        self.btn_start.configure(state="disabled", text="RUNNING...")
        self.btn_stop.configure(state="normal")
        self.btn_select_files.configure(state="disabled")
        self.btn_clear.configure(state="disabled")
        self.btn_back.configure(state="disabled")
        
        threading.Thread(target=self.process_thread, daemon=True).start()

    def stop_processing(self):
        if self.is_processing:
            self.stop_event.set()
            self.log(">>> STOP REQUESTED. Finishing current file...")
            self.status_label.configure(text="Stopping...")

    def reset_ui_after_process(self):
        self.is_processing = False
        self.btn_start.after(0, lambda: self.btn_start.configure(state="normal", text="START CHECK"))
        self.btn_stop.after(0, lambda: self.btn_stop.configure(state="disabled"))
        self.btn_select_files.after(0, lambda: self.btn_select_files.configure(state="normal"))
        self.btn_clear.after(0, lambda: self.btn_clear.configure(state="normal"))
        self.btn_back.after(0, lambda: self.btn_back.configure(state="normal"))
        self.status_label.after(0, lambda: self.status_label.configure(text="Ready."))

    def extract_qr_from_pdf(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
            pages_to_check = min(5, len(doc))
            for dpi in [150, 300]:
                for page_num in range(pages_to_check):
                    page = doc.load_page(page_num)
                    pix = page.get_pixmap(dpi=dpi) 
                    img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                    
                    if pix.n == 3:
                        img = cv2.cvtColor(img_data, cv2.COLOR_RGB2GRAY)
                    elif pix.n == 4:
                        img = cv2.cvtColor(img_data, cv2.COLOR_RGBA2GRAY)
                    else:
                        img = img_data
                    
                    try:
                        results = zxingcpp.read_barcodes(img)
                        for res in results:
                            if res.text: return res.text
                    except: pass
                    
                    _, thresh = cv2.threshold(img, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
                    try:
                        results = zxingcpp.read_barcodes(thresh)
                        for res in results:
                            if res.text: return res.text
                    except: pass
            return None
        except Exception as e:
            print(f"QR Error: {e}")
            return None

    def create_outlook_mail(self, outlook, attachment_path, subject_text, attachment_name=None):
        try:
            mail = outlook.CreateItem(0) 
            mail.Subject = subject_text
            mail.To = "sistematizacion@sparkgy.com"
            mail.Body = f"Adjunto el documento escaneado: {attachment_name if attachment_name else os.path.basename(attachment_path)}\\n\\nCodificado: {subject_text}"
            
            if attachment_name and attachment_name != os.path.basename(attachment_path):
                # Outlook renaming via local copy to ensure correct name
                temp_dir = os.path.join(os.environ.get('TEMP', os.getcwd()), "AutoArchiver_Temp")
                if not os.path.exists(temp_dir): os.makedirs(temp_dir)
                
                temp_path = os.path.join(temp_dir, attachment_name)
                shutil.copy2(attachment_path, temp_path)
                
                mail.Attachments.Add(temp_path)
                
                # Cleanup handled after send? sending is async usually but in COM usually we can del after Add.
                # To be safe, we won't delete immediately if it causes locks, but python script ends.
                # Actually, Attachments.Add copies file content into msg structure for unsent draft? 
                # For safety, we keep it or rely on OS temp cleanup, but let's try delete.
                try: 
                    pass # Deleting immediately might fail if Outlook holds lock. 
                except: pass
            else:
                mail.Attachments.Add(attachment_path)
                
            mail.Send()
            return True
        except Exception as e:
            self.update_log_safe(f"  > Outlook Error: {e}")
            return False

    def send_via_emailjs(self, attachment_path, subject_text, attachment_name=None):
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(base_path, "config.json")
        try:
            with open(config_path, "r") as f:
                config = json.load(f)
        except:
             return False
        service_id = config.get("service_id")
        template_id = config.get("template_id")
        user_id = config.get("user_id")
        access_token = config.get("access_token")
        
        try:
            with open(attachment_path, "rb") as f:
                encoded_string = base64.b64encode(f.read()).decode('utf-8')
        except: return False
        
        final_name = attachment_name if attachment_name else os.path.basename(attachment_path)

        url = "https://api.emailjs.com/api/v1.0/email/send"
        payload = {
            "service_id": service_id,
            "template_id": template_id,
            "user_id": user_id,
            "accessToken": access_token,
            "template_params": {
                "qr_text": subject_text,
                "file_name": final_name,
                "content": encoded_string 
            }
        }
        try:
            response = requests.post(url, json=payload)
            if response.status_code == 200:
                return True
            else:
                self.update_log_safe(f"  > EmailJS Failed: {response.text}")
                return False
        except Exception as e:
            self.update_log_safe(f"  > Connection Error: {e}")
            return False

    def process_thread(self):
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(base_path, "config.json")
        
        # Check switch state (1=ON, 0=OFF)
        switch_on = self.switch_email_mode.get()
        # Use EmailJS ONLY if switch is ON and config exists
        use_emailjs = (switch_on == 1) and os.path.exists(config_path)

        outlook = None
        
        if not use_emailjs:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception as e:
                self.update_log_safe(f"CRITICAL ERROR: Outlook not found/config missing.")
                self.after(0, lambda: messagebox.showerror("Error", "Outlook not found or configuration missing.\nPlease ensure Outlook is installed and configured."))
                self.reset_ui_after_process()
                return

        total = len(self.file_queue)
        success_count = 0
        failed_filenames = []

        for i, file_path in enumerate(self.file_queue):
            if self.stop_event.is_set():
                break
            filename = os.path.basename(file_path)
            self.update_log_safe(f"[{i+1}/{total}] Processing: {filename}...")
            
            try:
                qr_text = self.extract_qr_from_pdf(file_path)
                if qr_text:
                    qr_text = qr_text.replace("https://sparkgy.github.io/AutoArchiver/?", "")
                    qr_text = qr_text.replace("&", ";")
                    
                    # SANITIZATION: Replace _ and - with space in subject
                    qr_text = qr_text.replace("_", " ").replace("-", " ")

                    self.update_log_safe(f"  > QR Found: {qr_text}")
                    
                    # Determine Attachment Name
                    target_name = filename
                    if "NOMBREARCHIVOESCANEADO=" in qr_text:
                        try:
                            # Parse content after tag
                            parts = qr_text.split("NOMBREARCHIVOESCANEADO=")
                            if len(parts) > 1:
                                raw_name = parts[1].strip()
                                # Sanitize filename: Replace _ and - with space
                                raw_name = raw_name.replace("_", " ").replace("-", " ")
                                
                                # Keep simplified safe characters list but allow spaces now
                                safe_name = "".join([c for c in raw_name if c.isalnum() or c in (' ', '.')]).strip()
                                
                                if safe_name:
                                    if not safe_name.lower().endswith(".pdf"):
                                        safe_name += ".pdf"
                                    target_name = safe_name
                                    self.update_log_safe(f"  > Renaming attachment to: {target_name}")
                        except Exception as e:
                            self.update_log_safe(f"  > Name parsing error: {e}")

                    if use_emailjs:
                        success = self.send_via_emailjs(file_path, qr_text, target_name)
                    else:
                        success = self.create_outlook_mail(outlook, file_path, qr_text, target_name)
                    
                    if success: 
                        self.update_log_safe(f"  > Email Sent.")
                        success_count += 1
                    else: 
                        self.update_log_safe(f"  > FAILED to send email.")
                        failed_filenames.append(filename)
                else:
                    self.update_log_safe(f"  > WARNING: No QR code found.")
                    failed_filenames.append(filename)
            except Exception as e:
                self.update_log_safe(f"  > ERROR: {str(e)}")
                failed_filenames.append(filename)

            self.update_progress_safe((i + 1) / total)
        
        # Batch Summary
        failed_count = len(failed_filenames)
        summary_msg = f"--- Batch Process Completed ---\nProcessed: {success_count}/{total} files successfully."
        if failed_count > 0:
            summary_msg += f"\nFailed ({failed_count}):\n"
            for f in failed_filenames:
                summary_msg += f" - {f}\n"
        
        self.update_log_safe(summary_msg)
        self.reset_ui_after_process()


class ImageToPDFFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        # State Variables
        self.file_queue = []
        self.is_processing = False
        self.stop_event = threading.Event()
        self.pil_image = None
        self.zoom_level = 1.0
        self.pan_offset_x = 0
        self.pan_offset_y = 0
        self.drag_start_x = 0
        self.drag_start_y = 0

        # Layout Configuration
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Sidebar (Left) ---
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(8, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="IMG to PDF", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        self.sidebar_desc = ctk.CTkLabel(self.sidebar_frame, text="Batch Convert", font=ctk.CTkFont(size=12))
        self.sidebar_desc.grid(row=1, column=0, padx=20, pady=(0, 20))

        # Controls
        self.btn_select_files = ctk.CTkButton(self.sidebar_frame, text="Add Images...", command=self.select_files)
        self.btn_select_files.grid(row=2, column=0, padx=20, pady=10)

        self.btn_start = ctk.CTkButton(self.sidebar_frame, text="CONVERT", command=self.start_processing,
                                       fg_color="#2CC985", hover_color="#229C68", text_color="white") # Green
        self.btn_start.grid(row=3, column=0, padx=20, pady=10)

        self.btn_stop = ctk.CTkButton(self.sidebar_frame, text="STOP", command=self.stop_processing,
                                      fg_color="#D94040", hover_color="#A83232", state="disabled") # Red
        self.btn_stop.grid(row=4, column=0, padx=20, pady=10)

        self.btn_clear = ctk.CTkButton(self.sidebar_frame, text="Clear Queue", command=self.clear_queue,
                                       fg_color="transparent", border_width=1, text_color=("gray10", "#DCE4EE"))
        self.btn_clear.grid(row=5, column=0, padx=20, pady=(20, 10))
        
        # Back Button
        self.btn_back = ctk.CTkButton(self.sidebar_frame, text="← Menu", command=lambda: controller.show_frame("MainMenuFrame"),
                                      fg_color="transparent", border_width=1, text_color=("gray10", "#DCE4EE"))
        self.btn_back.grid(row=6, column=0, padx=20, pady=(20, 10))

        # --- Main Area (Right) ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.main_frame.grid_rowconfigure(0, weight=1) # Upper split
        self.main_frame.grid_rowconfigure(1, weight=0) # Log/Status
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Split Container (List + Preview)
        self.split_frame = ctk.CTkFrame(self.main_frame)
        self.split_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        self.split_frame.grid_columnconfigure(0, weight=1) # List
        self.split_frame.grid_columnconfigure(1, weight=3) # Preview
        self.split_frame.grid_rowconfigure(0, weight=1)

        # File List (Left of Split)
        self.file_listbox_frame = ctk.CTkFrame(self.split_frame, width=200)
        self.file_listbox_frame.grid(row=0, column=0, sticky="nsew", padx=(0,5))
        
        self.file_list_label = ctk.CTkLabel(self.file_listbox_frame, text="File Queue", font=ctk.CTkFont(weight="bold"))
        self.file_list_label.pack(pady=5)
        
        self.file_listbox = tk.Listbox(self.file_listbox_frame, bg="#2B2B2B", fg="#DCE4EE", 
                                       selectbackground="#1F6AA5", selectforeground="white", borderwidth=0, highlightthickness=0)
        self.file_listbox.pack(expand=True, fill="both", padx=5, pady=5)
        self.file_listbox.bind("<<ListboxSelect>>", self.on_list_select)

        # Viewer Frame
        self.viewer_frame = ctk.CTkFrame(self.split_frame, fg_color="#1A1A1A")
        self.viewer_frame.grid(row=0, column=1, sticky="nsew")
        self.viewer_frame.pack_propagate(False)

        # Toolbar
        self.toolbar_frame = ctk.CTkFrame(self.viewer_frame, height=40)
        self.toolbar_frame.pack(side="top", fill="x", padx=2, pady=2)
        
        self.btn_zoom_out = ctk.CTkButton(self.toolbar_frame, text="-", width=30, command=lambda: self.adjust_zoom(0.9))
        self.btn_zoom_out.pack(side="left", padx=2)
        
        self.btn_zoom_in = ctk.CTkButton(self.toolbar_frame, text="+", width=30, command=lambda: self.adjust_zoom(1.1))
        self.btn_zoom_in.pack(side="left", padx=2)

        self.btn_fit_width = ctk.CTkButton(self.toolbar_frame, text="Fit Width", width=70, command=self.fit_width)
        self.btn_fit_width.pack(side="left", padx=2)
        
        # Viewer Canvas
        self.canvas = tk.Canvas(self.viewer_frame, bg="#333333", highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<ButtonPress-2>", self.on_drag_start)
        self.canvas.bind("<B2-Motion>", self.on_drag_motion)
        self.canvas.bind("<MouseWheel>", self.on_zoom)

        # Log & Status (Bottom)
        self.bottom_frame = ctk.CTkFrame(self.main_frame, height=150)
        self.bottom_frame.grid(row=1, column=0, sticky="ew")
        self.bottom_frame.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(self.bottom_frame, text="Ready. Drag & Drop Image files here.", anchor="w")
        self.status_label.grid(row=0, column=0, sticky="w", padx=10, pady=(5,0))
        
        self.progressbar = ctk.CTkProgressBar(self.bottom_frame)
        self.progressbar.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        self.progressbar.set(0)

        self.log_textbox = ctk.CTkTextbox(self.bottom_frame, height=100, font=ctk.CTkFont(family="Consolas", size=11))
        self.log_textbox.grid(row=2, column=0, sticky="ew", padx=10, pady=(0,10))

        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.drop)

    # --- UI Logic (Simplified copy of PDFScannerFrame) ---
    def log(self, message):
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")

    def update_log_safe(self, message):
        self.after(0, lambda: self.log(message))

    def update_progress_safe(self, value):
        self.after(0, lambda: self.progressbar.set(value))

    def drop(self, event):
        if self.is_processing: return
        raw_files = self.split_file_list(event.data)
        count = 0
        valid_exts = (".jpg", ".jpeg", ".png")
        for f in raw_files:
            if f.lower().endswith(valid_exts) and f not in self.file_queue:
                self.file_queue.append(f)
                count += 1
        if count > 0:
            self.log(f"Added {count} files via Drag & Drop.")
            self.update_file_list_ui()

    def split_file_list(self, data):
        if data.startswith('{'):
            import re
            parts = re.findall(r'\{.*?\}|\S+', data) 
            return [p.strip('{}') for p in parts]
        return data.split()

    def select_files(self):
        if self.is_processing: return
        file_paths = filedialog.askopenfilenames(filetypes=[("Images", "*.jpg;*.jpeg;*.png")])
        if file_paths:
            count = 0
            for f in file_paths:
                if f not in self.file_queue:
                    self.file_queue.append(f)
                    count += 1
            if count > 0:
                self.log(f"Added {count} files.")
                self.update_file_list_ui()

    def clear_queue(self):
        if self.is_processing: return
        self.file_queue = []
        self.update_file_list_ui()
        self.clear_canvas()
        self.log("Queue cleared.")

    def update_file_list_ui(self):
        self.file_listbox.delete(0, "end")
        for f in self.file_queue:
            self.file_listbox.insert("end", os.path.basename(f))
        self.status_label.configure(text=f"Ready. {len(self.file_queue)} files in queue.")

    def on_list_select(self, event):
        selection = self.file_listbox.curselection()
        if not selection: return
        index = selection[0]
        file_path = self.file_queue[index]
        self.load_image_preview(file_path)

    def load_image_preview(self, file_path):
        try:
            self.pil_image = Image.open(file_path)
            self.zoom_level = 1.0
            self.fit_view()
        except Exception as e:
            self.log(f"Error previewing file: {e}")

    # Standard Viewer Logic
    def fit_view(self):
        if not hasattr(self, 'pil_image'): return
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        img_w, img_h = self.pil_image.size
        scale_w = canvas_w / img_w
        scale_h = canvas_h / img_h
        self.zoom_level = min(scale_w, scale_h) * 0.9
        self.display_image(reset_view=True)

    def fit_width(self):
        if not hasattr(self, 'pil_image'): return
        canvas_w = self.canvas.winfo_width()
        img_w, _ = self.pil_image.size
        self.zoom_level = (canvas_w / img_w) * 0.95
        self.display_image(reset_view=False)
        new_w = int(img_w * self.zoom_level)
        self.pan_offset_x = (canvas_w - new_w) // 2
        self.pan_offset_y = 20
        self._update_canvas_coords()

    def adjust_zoom(self, factor):
        self.zoom_level *= factor
        self.display_image()

    def display_image(self, reset_view=False):
        if not hasattr(self, 'pil_image'): return
        threading.Thread(target=self._resize_image_thread, args=(self.zoom_level, reset_view), daemon=True).start()

    def _resize_image_thread(self, zoom, reset_view):
        try:
            width, height = self.pil_image.size
            new_w = int(width * zoom)
            new_h = int(height * zoom)
            resized = self.pil_image.resize((new_w, new_h), Image.Resampling.BILINEAR)
            tk_image = ImageTk.PhotoImage(resized)
            self.after(0, self._update_canvas, tk_image, reset_view, new_w, new_h)
        except Exception as e:
            print(f"Resize error: {e}")

    def _update_canvas(self, tk_image, reset_view, new_w, new_h):
        self.tk_image = tk_image 
        self.canvas.delete("all")
        if reset_view:
            canvas_w = self.canvas.winfo_width()
            canvas_h = self.canvas.winfo_height()
            self.pan_offset_x = (canvas_w - new_w) // 2
            self.pan_offset_y = (canvas_h - new_h) // 2
        self._update_canvas_coords()

    def _update_canvas_coords(self):
        self.canvas.delete("all")
        self.image_id = self.canvas.create_image(
            self.pan_offset_x, 
            self.pan_offset_y, 
            image=self.tk_image, 
            anchor="nw"
        )

    def clear_canvas(self):
        self.canvas.delete("all")
        self.pil_image = None
    
    def on_drag_start(self, event):
        self.drag_start_x = event.x
        self.drag_start_y = event.y

    def on_drag_motion(self, event):
        dx = event.x - self.drag_start_x
        dy = event.y - self.drag_start_y
        self.pan_offset_x += dx
        self.pan_offset_y += dy
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self._update_canvas_coords()

    def on_zoom(self, event):
        if not hasattr(self, 'pil_image'): return
        old_zoom = self.zoom_level
        zoom_factor = 1.1 if event.delta > 0 else 0.9
        new_zoom = old_zoom * zoom_factor
        new_zoom = max(0.1, min(10.0, new_zoom))
        if new_zoom == old_zoom: return
        self.zoom_level = new_zoom
        verify_x = event.x - self.pan_offset_x
        verify_y = event.y - self.pan_offset_y
        new_verify_x = verify_x * zoom_factor
        new_verify_y = verify_y * zoom_factor
        self.pan_offset_x = int(event.x - new_verify_x)
        self.pan_offset_y = int(event.y - new_verify_y)
        self.display_image(reset_view=False)

    # --- Processing ---
    def start_processing(self):
        if not self.file_queue:
            messagebox.showinfo("Info", "No files in queue to process.")
            return

        self.is_processing = True
        self.stop_event.clear()
        self.btn_start.configure(state="disabled", text="RUNNING...")
        self.btn_stop.configure(state="normal")
        self.btn_select_files.configure(state="disabled")
        self.btn_clear.configure(state="disabled")
        self.btn_back.configure(state="disabled")
        
        threading.Thread(target=self.process_thread, daemon=True).start()

    def stop_processing(self):
        if self.is_processing:
            self.stop_event.set()
            self.log(">>> STOP REQUESTED...")

    def reset_ui_after_process(self):
        self.is_processing = False
        self.btn_start.after(0, lambda: self.btn_start.configure(state="normal", text="CONVERT"))
        self.btn_stop.after(0, lambda: self.btn_stop.configure(state="disabled"))
        self.btn_select_files.after(0, lambda: self.btn_select_files.configure(state="normal"))
        self.btn_clear.after(0, lambda: self.btn_clear.configure(state="normal"))
        self.btn_back.after(0, lambda: self.btn_back.configure(state="normal"))
        self.status_label.after(0, lambda: self.status_label.configure(text="Ready."))

    def process_thread(self):
        # Determine output path
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        output_dir = os.path.join(base_path, "output")
        
        self.update_log_safe(f"Preparing output directory: {output_dir}")
        try:
            if os.path.exists(output_dir):
                shutil.rmtree(output_dir)
            os.makedirs(output_dir)
        except Exception as e:
            self.update_log_safe(f"CRITICAL ERROR: Output directory issue: {e}")
            self.reset_ui_after_process()
            return

        total = len(self.file_queue)
        for i, file_path in enumerate(self.file_queue):
            if self.stop_event.is_set():
                break
                
            filename = os.path.basename(file_path)
            self.update_log_safe(f"[{i+1}/{total}] Converting: {filename}...")
            
            try:
                # Convert logic
                img = Image.open(file_path)
                # Ensure RGB
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                name_root = os.path.splitext(filename)[0]
                pdf_path = os.path.join(output_dir, f"{name_root}.pdf")
                
                img.save(pdf_path, "PDF", resolution=100.0)
                self.update_log_safe(f"  > Saved to output/{name_root}.pdf")
                
            except Exception as e:
                self.update_log_safe(f"  > ERROR: {str(e)}")
            
            self.update_progress_safe((i + 1) / total)
            
        self.update_log_safe("--- Conversion Completed ---")
        self.reset_ui_after_process()


class AutoArchiverApp(DndCTk):
    def __init__(self):
        super().__init__()

        self.title("AutoArchiver - Spark Energy")
        self.geometry("1100x700")

        # Set Icon
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sparkgy.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)

        self.container = ctk.CTkFrame(self)
        self.container.pack(side="top", fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (MainMenuFrame, PDFScannerFrame, ImageToPDFFrame):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("MainMenuFrame")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()


if __name__ == "__main__":
    app = AutoArchiverApp()
    app.mainloop()
