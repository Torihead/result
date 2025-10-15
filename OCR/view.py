from pdf2image import convert_from_path
from PIL import Image
import pytesseract

# Tesseractのパス（Windowsの場合）
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\TesseractOCR\tesseract.exe"

# PDF → 画像変換
images = convert_from_path(r"C:\Users\USER06\Desktop\Auto_code\OCR\製造表.pdf", dpi=300)

# OCR処理
for i, image in enumerate(images):
    text = pytesseract.image_to_string(image, lang='jpn')
    with open(f'page_{i+1}.txt', 'w', encoding='utf-8') as f:
        f.write(text)
    print(f"--- Page {i+1} ---\n{text}\n")