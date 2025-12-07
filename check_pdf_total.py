import pdfplumber
import re

pdf_path = r'c:\Users\HP\Desktop\호텔\부킹매출자료\부킹3.pdf'

def inspect_pdf_total():
    print(f"Inspecting {pdf_path}...")
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"
        
        print("--- Full Text Search for Totals ---")
        # Look for keywords like "Total", "합계", "청구 금액", "지급 금액"
        # Also look for the amount 69,519,200
        
        if "69,519,200" in full_text or "69519200" in full_text.replace(",",""):
            print("Found explicit amount 69,519,200 in text.")
        else:
            print("Amount 69,519,200 NOT found in text.")
            
        # Print lines containing numbers that might be totals
        lines = full_text.split('\n')
        for line in lines:
            if "합계" in line or "Total" in line or "지급" in line or "Amount" in line:
                print(f"Relevant Line: {line}")

if __name__ == "__main__":
    inspect_pdf_total()
