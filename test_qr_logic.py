import os
import fitz
import cv2
import numpy as np

def extract_qr_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        pages_to_check = min(5, len(doc))
        detector = cv2.QRCodeDetector()

        for dpi in [150, 300]:
            for page_num in range(pages_to_check):
                page = doc.load_page(page_num)
                pix = page.get_pixmap(dpi=dpi) 
                
                img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                
                if pix.n == 3: img = cv2.cvtColor(img_data, cv2.COLOR_RGB2BGR)
                elif pix.n == 4: img = cv2.cvtColor(img_data, cv2.COLOR_RGBA2BGR)
                else: img = cv2.cvtColor(img_data, cv2.COLOR_GRAY2BGR)
                
                # 1. Standard
                data, _, _ = detector.detectAndDecode(img)
                if data: return data, f"Standard {dpi}dpi"
                
                # 2. Preprocessing
                gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                
                data, _, _ = detector.detectAndDecode(gray)
                if data: return data, f"Gray {dpi}dpi"
                
                _, thresh = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
                data, _, _ = detector.detectAndDecode(thresh)
                if data: return data, f"Otsu {dpi}dpi"

                adaptive = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
                data, _, _ = detector.detectAndDecode(adaptive)
                if data: return data, f"Adaptive {dpi}dpi"

        return None, None
    except Exception as e:
        return None, str(e)

def test_folder(folder_path):
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    print(f"Testing {len(files)} files in {folder_path}...\n")
    
    success_count = 0
    failed_files = []
    for f in files:
        full_path = os.path.join(folder_path, f)
        qr, method = extract_qr_from_pdf(full_path)
        if qr:
            print(f"[OK] {f[:30]}... -> {qr[:20]}... (Method: {method})")
            success_count += 1
        else:
            print(f"[FAIL] {f[:30]}...")
            failed_files.append(f)

    print(f"\nSummary: {success_count}/{len(files)} files detected.")
    if failed_files:
        print("Failed Files:")
        for ff in failed_files:
            print(f" - {ff}")

if __name__ == "__main__":
    target_folder = r"c:\Users\sae_v\OneDrive - SPARK ENERGY SOLUTIONS\SPARKTRACK\Autoarchiver\AutoArchiver\Prueba"
    test_folder(target_folder)
