import pdfplumber
import pytesseract
from PIL import Image

# 設定 Tesseract 執行檔的路徑
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # 將 PDF 頁面轉換成 PIL 影像格式
            pil_image = page.to_image(resolution=300).original
            pil_image = pil_image.convert("L")
            # 使用 Tesseract 提取文字
            page_text = pytesseract.image_to_string(pil_image,lang='chi_tra')
            text += page_text
    return text

# 使用範例
pdf_path = r"C:\Users\lin78\OneDrive\文件\永續report\110永續報告書(中文)\&fileName=t100sa11_1110_110.pdf"
extracted_text = extract_text_from_pdf(pdf_path)
print(extracted_text)