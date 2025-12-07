import pdfplumber
import re

file_path = r'c:\Users\HP\Desktop\호텔\부킹매출자료\부킹1.pdf'

try:
    with pdfplumber.open(file_path) as pdf:
        # Check the last page for Total Amount
        last_page = pdf.pages[-1]
        text = last_page.extract_text()
        print("--- Last Page Text ---")
        print(text)
        
        # Regex to find Total Amount
        # Pattern: 대금 총액 ₩69,519,200 or similar
        match = re.search(r'대금 총액\s+₩([\d,]+)', text)
        if match:
            print(f"Found Total Amount: {match.group(1)}")
        else:
            print("Total Amount not found via regex.")

except Exception as e:
    print(f"Error reading PDF: {e}")
