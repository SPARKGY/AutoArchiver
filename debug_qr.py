import os
import fitz
import cv2
import numpy as np

def try_detect(img, description):
    detector = cv2.QRCodeDetector()
    try:
        data, _, _ = detector.detectAndDecode(img)
        if data: return data
    except: pass
    return None

def test_file(pdf_path):
    print(f"Testing: {os.path.basename(pdf_path)}")
    try:
        doc = fitz.open(pdf_path)
        # Try different DPIs
        for dpi in [150, 200, 300, 400]:
            page = doc.load_page(0)
            pix = page.get_pixmap(dpi=dpi)
            
            img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            if pix.n == 3: img = cv2.cvtColor(img_data, cv2.COLOR_RGB2BGR)
            elif pix.n == 4: img = cv2.cvtColor(img_data, cv2.COLOR_RGBA2BGR)
            else: img = cv2.cvtColor(img_data, cv2.COLOR_GRAY2BGR)
            
            # 1. Raw
            if res := try_detect(img, f"Raw {dpi}dpi"): 
                print(f"SUCCESS: Raw {dpi}dpi")

            # 2. Gray
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            if res := try_detect(gray, f"Gray {dpi}dpi"): 
                print(f"SUCCESS: Gray {dpi}dpi")

            # 3. Gaussian Blur + Otsu
            blur = cv2.GaussianBlur(gray, (5,5), 0)
            _, otsu = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            if res := try_detect(otsu, f"Otsu {dpi}dpi"): 
                print(f"SUCCESS: Otsu {dpi}dpi"); return res

            # 4. Adaptive Threshold
            adaptive = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
            if res := try_detect(adaptive, f"Adaptive {dpi}dpi"): 
                print(f"SUCCESS: Adaptive {dpi}dpi"); return res

            # 5. Contrast Limited Adaptive Histogram Equalization (CLAHE)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
            cl1 = clahe.apply(gray)
            if res := try_detect(cl1, f"CLAHE {dpi}dpi"): 
                print(f"SUCCESS: CLAHE {dpi}dpi"); return res
            
            # 6. Sharpening kernel
            kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
            sharp = cv2.filter2D(gray, -1, kernel)
            if res := try_detect(sharp, f"Sharp {dpi}dpi"): 
                print(f"SUCCESS: Sharp {dpi}dpi"); return res

    except Exception as e:
        print(f"Error: {e}")
    
    print("FAILED all methods.")
    return None

import glob
folder = r"c:\Users\sae_v\OneDrive - SPARK ENERGY SOLUTIONS\SPARKTRACK\Autoarchiver\AutoArchiver\Prueba"
files = glob.glob(os.path.join(folder, "*.pdf"))
if files:
    # Test on the first file first, as it failed before
    test_file(files[0])
else:
    print("No files found.")
