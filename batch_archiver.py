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
    import win32com.client
    from PIL import Image, ImageTk
    from tkinterdnd2 import TkinterDnD, DND_FILES
    import requests
    import json
    import base64
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

class AutoArchiverApp(DndCTk):
    def __init__(self):
        super().__init__()

        self.title("AutoArchiver - Spark Energy")
        self.geometry("1100x700")

        # Set Icon
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sparkgy.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)

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

        # Layout Configuration
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Sidebar (Left) ---
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(7, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="AutoArchiver", font=ctk.CTkFont(size=24, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        self.sidebar_desc = ctk.CTkLabel(self.sidebar_frame, text="PDF QR Scanner\n& Outlook Drafts", font=ctk.CTkFont(size=12))
        self.sidebar_desc.grid(row=1, column=0, padx=20, pady=(0, 20))

        # Controls
        self.btn_select_files = ctk.CTkButton(self.sidebar_frame, text="Add Files...", command=self.select_files)
        self.btn_select_files.grid(row=2, column=0, padx=20, pady=10)

        self.btn_start = ctk.CTkButton(self.sidebar_frame, text="START CHECK", command=self.start_processing,
                                       fg_color="#2CC985", hover_color="#229C68", text_color="white") # Green
        self.btn_start.grid(row=3, column=0, padx=20, pady=10)

        self.btn_stop = ctk.CTkButton(self.sidebar_frame, text="STOP", command=self.stop_processing,
                                      fg_color="#D94040", hover_color="#A83232", state="disabled") # Red
        self.btn_stop.grid(row=4, column=0, padx=20, pady=10)

        self.btn_clear = ctk.CTkButton(self.sidebar_frame, text="Clear Queue", command=self.clear_queue,
                                       fg_color="transparent", border_width=1, text_color=("gray10", "#DCE4EE"))
        self.btn_clear.grid(row=5, column=0, padx=20, pady=(20, 10))

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
        
        self.viewer_label = ctk.CTkLabel(self.viewer_frame, text="Document Preview (Page 1)", font=ctk.CTkFont(size=10))
        self.viewer_label.pack(side="top", pady=2)
        
        self.canvas = tk.Canvas(self.viewer_frame, bg="#333333", highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Canvas Controls
        self.canvas.bind("<ButtonPress-1>", self.on_drag_start)
        self.canvas.bind("<B1-Motion>", self.on_drag_motion)
        self.canvas.bind("<MouseWheel>", self.on_zoom)  # Windows mouse wheel

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

        # DnD Registration
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.drop)

        # Internal Init
        self.update_file_list_ui()

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
        # tkinterdnd2 returns a brace-quoted string list sometimes
        if data.startswith('{'):
            # This is a robust regex-less split for Tcl list standard
            # but for simplicity in Windows paths often don't have spaces if handled by OS, 
            # or arc wrapped in {}. 
            # Simple parser:
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
        self.load_pdf_preview(file_path)

    def load_pdf_preview(self, file_path):
        try:
            doc = fitz.open(file_path)
            self.current_pdf_doc = doc
            page = doc.load_page(0) 
            pix = page.get_pixmap(dpi=150) # Standard preview DPI
            
            # Convert to PIL
            img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            if pix.n == 4: # RGBA
                mode = "RGBA"
            elif pix.n == 3: # RGB
                mode = "RGB"
            else:
                mode = "L" # Grayscale
                
            self.pil_image = Image.frombytes(mode, (pix.w, pix.h), pix.samples)
            if mode == "RGB": pass # OK
            
            self.zoom_level = 1.0
            self.pan_offset_x = 0
            self.pan_offset_y = 0
            self.display_image()
            
        except Exception as e:
            self.log(f"Error previewing file: {e}")

    def display_image(self):
        if not hasattr(self, 'pil_image'): return
        
        # Resize based on zoom
        width, height = self.pil_image.size
        new_w = int(width * self.zoom_level)
        new_h = int(height * self.zoom_level)
        
        resized = self.pil_image.resize((new_w, new_h), Image.Resampling.LANCZOS)
        self.tk_image = ImageTk.PhotoImage(resized)
        
        self.canvas.delete("all")
        # Center image approx
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        
        x = (canvas_w // 2) + self.pan_offset_x
        y = (canvas_h // 2) + self.pan_offset_y
        
        self.canvas.create_image(x, y, image=self.tk_image, anchor="center")

    def clear_canvas(self):
        self.canvas.delete("all")
        self.current_pdf_doc = None

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
        self.display_image()

    def on_zoom(self, event):
        if event.delta > 0:
            self.zoom_level *= 1.1
        else:
            self.zoom_level *= 0.9
        self.display_image()

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
        
        threading.Thread(target=self.process_thread, daemon=True).start()

    def stop_processing(self):
        if self.is_processing:
            self.stop_event.set()
            self.log(">>> STOP REQUESTED. Finishing current file...")
            self.status_label.configure(text="Stopping...")

    def reset_ui_after_process(self):
        self.is_processing = False
        # Schedule UI updates back on the main thread
        self.btn_start.after(0, lambda: self.btn_start.configure(state="normal", text="START CHECK"))
        self.btn_stop.after(0, lambda: self.btn_stop.configure(state="disabled"))
        self.btn_select_files.after(0, lambda: self.btn_select_files.configure(state="normal"))
        self.btn_clear.after(0, lambda: self.btn_clear.configure(state="normal"))
        self.status_label.after(0, lambda: self.status_label.configure(text="Ready."))

    # --- Visualization ---
    def on_list_select(self, event):
        selection = self.file_listbox.curselection()
        if not selection: return
        index = selection[0]
        file_path = self.file_queue[index]
        # Debounce or cancel previous? Simple way: just start new thread.
        # Ideally we'd have a 'current_request_id' to ignore old results.
        self.load_pdf_preview_async(file_path)

    def load_pdf_preview_async(self, file_path):
        threading.Thread(target=self._load_pdf_thread, args=(file_path,), daemon=True).start()

    def _load_pdf_thread(self, file_path):
        try:
            doc = fitz.open(file_path)
            page = doc.load_page(0) 
            pix = page.get_pixmap(dpi=150) # Standard preview DPI
            
            # Convert to PIL
            mode = "RGB" if pix.n >= 3 else "L"
            if pix.n == 4: mode = "RGBA"
                
            pil_image = Image.frombytes(mode, (pix.w, pix.h), pix.samples)
            if mode == "RGBA":
                pil_image = pil_image.convert("RGB")
            
            # Update UI on main thread
            self.after(0, self._on_pdf_loaded, doc, pil_image)
            
        except Exception as e:
            self.update_log_safe(f"Error previewing file: {e}")

    def _on_pdf_loaded(self, doc, pil_image):
        if self.current_pdf_doc:
            try: self.current_pdf_doc.close()
            except: pass
            
        self.current_pdf_doc = doc
        self.pil_image = pil_image
        self.zoom_level = 1.0
        self.pan_offset_x = 0
        self.pan_offset_y = 0
        
        self.display_image(reset_view=True)

    def display_image(self, reset_view=False):
        if not hasattr(self, 'pil_image'): return
        
        # Offload resizing to thread to keep UI responsive
        threading.Thread(target=self._resize_image_thread, args=(self.zoom_level, reset_view), daemon=True).start()

    def _resize_image_thread(self, zoom, reset_view):
        try:
            # Resize based on zoom
            width, height = self.pil_image.size
            new_w = int(width * zoom)
            new_h = int(height * zoom)
            
            # Use BILINEAR (Fast & Good enough)
            resized = self.pil_image.resize((new_w, new_h), Image.Resampling.BILINEAR)
            tk_image = ImageTk.PhotoImage(resized)
            
            # Update Canvas on Main Thread
            self.after(0, self._update_canvas, tk_image, reset_view, new_w, new_h)
        except Exception as e:
            print(f"Resize error: {e}")

    def _update_canvas(self, tk_image, reset_view, new_w, new_h):
        # Keep reference to avoid GC
        self.tk_image = tk_image 
        
        self.canvas.delete("all")
        
        if reset_view:
            canvas_w = self.canvas.winfo_width()
            canvas_h = self.canvas.winfo_height()
            self.pan_offset_x = (canvas_w - new_w) // 2
            self.pan_offset_y = (canvas_h - new_h) // 2

        self.image_id = self.canvas.create_image(
            self.pan_offset_x, 
            self.pan_offset_y, 
            image=self.tk_image, 
            anchor="nw"
        )

    def clear_canvas(self):
        self.canvas.delete("all")
        self.current_pdf_doc = None

    def on_drag_start(self, event):
        self.canvas.scan_mark(event.x, event.y)

    def on_drag_motion(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)
        # Update offsets to track current position for zoom centering
        # This is strictly not needed if we just trust scan_dragto visual, 
        # but to keep Zoom concentric we might need coords.
        # For simple pan speed, scan_dragto is native and fast.
        
    def on_zoom(self, event):
        old_zoom = self.zoom_level
        if event.delta > 0:
            self.zoom_level *= 1.1
        else:
            self.zoom_level *= 0.9
            
        # Limit zoom
        self.zoom_level = max(0.1, min(5.0, self.zoom_level))
        
        if old_zoom != self.zoom_level:
            # Simple zoom: Re-render centered (simplification for speed)
            # To do proper mouse-centered zoom requires complex offset math.
            # Sticking to simple center zoom or keeping current top-left.
            self.display_image() # Re-renders

    # ... (Processing methods remain same until extract_qr) ...

    def extract_qr_from_pdf(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
            # Check up to 5 pages
            pages_to_check = min(5, len(doc))
            detector = cv2.QRCodeDetector()

            # Try 150 DPI first (Faster, often better for large QRs), then 300 DPI
            for dpi in [150, 300]:
                for page_num in range(pages_to_check):
                    page = doc.load_page(page_num)
                    pix = page.get_pixmap(dpi=dpi) 
                    
                    # Convert to numpy
                    img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                    
                    # Convert to BGR/Gray
                    if pix.n == 3: # RGB
                        img = cv2.cvtColor(img_data, cv2.COLOR_RGB2BGR)
                    elif pix.n == 4: # RGBA
                        img = cv2.cvtColor(img_data, cv2.COLOR_RGBA2BGR)
                    else: # Gray or other
                        img = cv2.cvtColor(img_data, cv2.COLOR_GRAY2BGR)
                    
                    # 1. Standard Detect
                    data, _, _ = detector.detectAndDecode(img)
                    if data: return data
                    
                    # 2. Grayscale + Preprocessing methods
                    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                    
                    # A. Pure Gray
                    data, _, _ = detector.detectAndDecode(gray)
                    if data: return data
                    
                    # B. Standard Threshold
                    _, thresh = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
                    data, _, _ = detector.detectAndDecode(thresh)
                    if data: return data

                    # C. Adaptive Threshold (Good for uneven lighting/shadows)
                    adaptive = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
                    data, _, _ = detector.detectAndDecode(adaptive)
                    if data: return data

            return None
        except Exception as e:
            print(f"QR Error: {e}")
            return None

    def create_outlook_mail(self, outlook, attachment_path, subject_text):
        try:
            mail = outlook.CreateItem(0) 
            mail.Subject = subject_text
            mail.To = "sistematizacion@sparkgy.com"
            mail.Body = f"Adjunto el documento escaneado: {os.path.basename(attachment_path)}\n\nCodificado: {subject_text}"
            mail.Attachments.Add(attachment_path)
            mail.Send() # CHANGED FROM SAVE TO SEND
            return True
        except Exception as e:
            self.update_log_safe(f"  > Outlook Error: {e}")
            return False

    def send_via_emailjs(self, attachment_path, subject_text):
        # Determine paths carefully for config
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        config_path = os.path.join(base_path, "config.json")
        
        if not os.path.exists(config_path):
             self.update_log_safe(f"  > ERROR: 'config.json' not found at {config_path}. Cannot use EmailJS.")
             return False
             
        try:
            with open(config_path, "r") as f:
                config = json.load(f)
        except:
             self.update_log_safe("  > ERROR: Invalid 'config.json'.")
             return False

        service_id = config.get("service_id")
        template_id = config.get("template_id")
        user_id = config.get("user_id")
        
        if not all([service_id, template_id, user_id]):
            self.update_log_safe("  > ERROR: Missing keys in 'config.json'.")
            return False
            
        # Read file and encode
        try:
            with open(attachment_path, "rb") as f:
                encoded_string = base64.b64encode(f.read()).decode('utf-8')
        except Exception as e:
            self.update_log_safe(f"  > ERROR reading file: {e}")
            return False
            
        url = "https://api.emailjs.com/api/v1.0/email/send"
        payload = {
            "service_id": service_id,
            "template_id": template_id,
            "user_id": user_id,
            "template_params": {
                "qr_text": subject_text,
                "file_name": os.path.basename(attachment_path),
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
        # Determine mode: Outlook or EmailJS
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        config_path = os.path.join(base_path, "config.json")
        use_emailjs = os.path.exists(config_path)
        outlook = None
        
        if not use_emailjs:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception as e:
                self.update_log_safe(f"CRITICAL ERROR: Could not connect to Outlook and no config.json found.")
                self.reset_ui_after_process()
                return
        else:
            self.update_log_safe("Using EmailJS configuration...")

        total = len(self.file_queue)
        
        for i, file_path in enumerate(self.file_queue):
            if self.stop_event.is_set():
                self.update_log_safe(">>> BATCH STOPPED BY USER.")
                break

            filename = os.path.basename(file_path)
            self.update_log_safe(f"[{i+1}/{total}] Processing: {filename}...")
            
            try:
                qr_text = self.extract_qr_from_pdf(file_path)
                
                if qr_text:
                    self.update_log_safe(f"  > QR Found: {qr_text}")
                    
                    success = False
                    if use_emailjs:
                        success = self.send_via_emailjs(file_path, qr_text)
                    else:
                        success = self.create_outlook_mail(outlook, file_path, qr_text)
                    
                    if success:
                        self.update_log_safe(f"  > Email Sent.")
                    else:
                        self.update_log_safe(f"  > FAILED to send email.")
                else:
                    self.update_log_safe(f"  > WARNING: No QR code found.")

            except Exception as e:
                self.update_log_safe(f"  > ERROR: {str(e)}")
            
            self.update_progress_safe((i + 1) / total)
        
        self.update_log_safe("--- Batch Process Completed ---")
        self.reset_ui_after_process()

if __name__ == "__main__":
    app = AutoArchiverApp()
    app.mainloop()
