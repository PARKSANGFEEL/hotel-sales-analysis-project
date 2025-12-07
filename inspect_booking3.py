import pdfplumber

file_path = r'c:\Users\HP\Desktop\호텔\부킹매출자료\부킹3.pdf'

try:
    with pdfplumber.open(file_path) as pdf:
        print(f"Total pages: {len(pdf.pages)}")
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            print(f"--- Page {i+1} ---")
            print(text)
            print("------------------")
except Exception as e:
    print(f"Error reading PDF: {e}")
