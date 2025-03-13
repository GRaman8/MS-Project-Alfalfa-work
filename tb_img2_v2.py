import os
import sys
from pdf2image import convert_from_path
from img2table.document import PDF, Image
from img2table.ocr import TesseractOCR

def parse_page_numbers(page_input, total_pages):
    """Parses a user input string like '1,3,5-7' into a list of integers."""
    selected_pages = set()
    
    for part in page_input.split(","):
        part = part.strip()
        if "-" in part:
            try:
                start, end = map(int, part.split("-"))
                if start <= end and 1 <= start <= total_pages and 1 <= end <= total_pages:
                    selected_pages.update(range(start, end + 1))
            except ValueError:
                print(f"Invalid range: {part}")
        else:
            try:
                page = int(part)
                if 1 <= page <= total_pages:
                    selected_pages.add(page)
            except ValueError:
                print(f"Invalid page number: {part}")

    return sorted(selected_pages)

def main():
    file_path = input("Enter the full path to your document (PDF or Image): ").strip()

    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        sys.exit(1)

    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == ".pdf":
        # Convert PDF to images to get total page count
        images = convert_from_path(file_path)
        total_pages = len(images)
        print(f"Total pages in PDF: {total_pages}")

        # Ask user for specific pages to process
        page_input = input(f"Enter page numbers to process (1-{total_pages}, e.g., '1,3,5-7'): ").strip()
        selected_pages = parse_page_numbers(page_input, total_pages)

        if not selected_pages:
            print("No valid pages selected. Exiting...")
            sys.exit(1)

        # Load the PDF with the selected pages for processing
        doc = PDF(file_path, pages=selected_pages, pdf_text_extraction=True)
    
    elif ext in [".png", ".jpg", ".jpeg", ".bmp", ".tiff"]:
        doc = Image(file_path, detect_rotation=False)
    else:
        print("Unsupported file type! Please provide a PDF or an image (png, jpg, jpeg, bmp, tiff).")
        sys.exit(1)

    borderless_tables = input("Does the document contain borderless tables? (Y/N): ").strip().lower() == "y"
    implicit_rows = input("Does the document contain implicit rows? (Y/N): ").strip().lower() == "y"
    implicit_columns = input("Does the document contain implicit columns? (Y/N): ").strip().lower() == "y"

    ocr = TesseractOCR(n_threads=1, lang="eng")

    output_excel = input("Enter the desired output Excel file name (e.g., output.xlsx): ").strip()
    if not output_excel.endswith(".xlsx"):
        output_excel += ".xlsx"

    try:
        doc.to_xlsx(dest=output_excel,
                    ocr=ocr,
                    implicit_rows=implicit_rows,
                    implicit_columns=implicit_columns,
                    borderless_tables=borderless_tables,
                    min_confidence=50)
        print(f"✅ Success: Excel file created at {output_excel}")
    except Exception as e:
        print("❌ An error occurred during table extraction:", e)
        sys.exit(1)

if __name__ == "__main__":
    main()
