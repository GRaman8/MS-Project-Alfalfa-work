#This version of the code can take multiple pdf/images as input and if they have portrait tables in them. This code can extract those tables.

import os
import sys
# import subprocess
from pdf2image import convert_from_path
from img2table.document import PDF, Image
from img2table.ocr import TesseractOCR
from PIL import Image as PILImage  # for image rotation

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

#def process_pdf(file_path, ocr, implicit_rows, implicit_columns, borderless_tables):
    """Processes a PDF file with the option to rotate portrait pages."""
    # Convert PDF pages to images so we can both count pages and later rotate if needed.
    images = convert_from_path(file_path)
    total_pages = len(images)
    print(f"Total pages in {file_path}: {total_pages}")
    
    # Ask user which pages to process from the PDF
    page_input = input(f"Enter page numbers to process for {file_path} (1-{total_pages}, e.g., '1,3,5-7'): ").strip()
    selected_pages = parse_page_numbers(page_input, total_pages)
    if not selected_pages:
        print(f"No valid pages selected for {file_path}. Skipping...")
        return

    # Ask if any of the selected pages contain portrait tables.
    portrait_choice = input("Do any of these pages contain portrait tables? (Y/N): ").strip().lower() == "y"
    portrait_pages = []
    if portrait_choice:
        portrait_input = input("Enter page numbers (from the selected pages) that are portrait (e.g., '2,5'): ").strip()
        portrait_pages = parse_page_numbers(portrait_input, total_pages)
        # Only keep pages that were selected for processing.
        portrait_pages = [p for p in portrait_pages if p in selected_pages]

    # Determine landscape pages as those selected but not marked as portrait.
    landscape_pages = [p for p in selected_pages if p not in portrait_pages]

    # Process landscape pages using PDF extraction (if any)
    if landscape_pages:
        doc_landscape = PDF(file_path, pages=landscape_pages, pdf_text_extraction=True)
        output_excel_landscape = os.path.splitext(os.path.basename(file_path))[0] + "_landscape.xlsx"
        try:
            doc_landscape.to_xlsx(dest=output_excel_landscape,
                                  ocr=ocr,
                                  implicit_rows=implicit_rows,
                                  implicit_columns=implicit_columns,
                                  borderless_tables=borderless_tables,
                                  min_confidence=50)
            print(f"✅ Success: Excel file created at {output_excel_landscape} for landscape pages.")
        except Exception as e:
            print(f"❌ An error occurred while processing landscape pages for {file_path}:", e)
    
    if portrait_pages:
        for p in portrait_pages:
            original_image = images[p -1]
            rotated_image = original_image.rotate(-90, expand =True)

            temp_path = f"temp_page_{p}.png"
            rotated_image.save(temp_path)
            doc_img = Image(temp_path, detect_rotation=False)
            output_excel_portrait = os.path.splitext(os.path.basename(file_path))[0] + f"_portrait_page_{p}.xlsx"

            try:
                doc_img.to_xlsx(dest = output_excel_portrait,
                                ocr = ocr,
                                implicit_rows = implicit_rows,
                                implicit_columns = implicit_columns,
                                borderless_tables = borderless_tables,
                                min_confidence=50)
                print(f"✅ Success: Excel file created at {output_excel_portrait} for portrait page {p}.")
            except Exception as e:
                print(f"❌ An error occurred while processing portrait page {p} for {file_path}:", e)

            os.remove(temp_path)    

def process_pdf(file_path, ocr, implicit_rows, implicit_columns, borderless_tables):
    """Processes a PDF file with user-specified portrait and landscape pages."""
    # Convert PDF pages to images to count pages
    images = convert_from_path(file_path)
    total_pages = len(images)
    print(f"Total pages in {file_path}: {total_pages}")

    # Ask user which pages to process
    page_input = input(f"Enter page numbers to process for {file_path} (1-{total_pages}, e.g., 1->0 , 2->1): ").strip()
    selected_pages = parse_page_numbers(page_input, total_pages)
    if not selected_pages:
        print(f"No valid pages selected for {file_path}. Skipping...")
        return

    # Ask which pages have portrait tables
    portrait_choice = input("Do any of these pages contain portrait tables? (Y/N): ").strip().lower() == "y"
    portrait_pages = []
    if portrait_choice:
        portrait_input = input("Enter page numbers (from the selected pages) that are portrait (e.g., 1->0 , 2->1): ").strip()
        portrait_pages = parse_page_numbers(portrait_input, total_pages)
        portrait_pages = [p for p in portrait_pages if p in selected_pages]

    # ✅ Ask the user which pages are landscape
    landscape_choice = input("Do any of these pages contain landscape tables? (Y/N): ").strip().lower() == "y"
    landscape_pages = []
    if landscape_choice:
        landscape_input = input("Enter page numbers (from the selected pages) that are landscape (e.g., 1->0 , 2->1): ").strip()
        landscape_pages = parse_page_numbers(landscape_input, total_pages)
        landscape_pages = [p-1 for p in landscape_pages if p in selected_pages]

    # Process landscape pages using PDF extraction (if any)
    if landscape_pages:
        doc_landscape = PDF(file_path, pages=sorted(landscape_pages), pdf_text_extraction=True)
        output_excel_landscape = os.path.splitext(os.path.basename(file_path))[0] + "_landscape.xlsx"
        try:
            doc_landscape.to_xlsx(dest=output_excel_landscape,
                                  ocr=ocr,
                                  implicit_rows=implicit_rows,
                                  implicit_columns=implicit_columns,
                                  borderless_tables=borderless_tables,
                                  min_confidence=50)
            print(f"✅ Success: Excel file created at {output_excel_landscape} for landscape pages.")
        except Exception as e:
            print(f"❌ An error occurred while processing landscape pages for {file_path}:", e)

    # Process portrait pages separately
    if portrait_pages:
        for p in portrait_pages:
            original_image = images[p - 1]
            rotated_image = original_image.rotate(-90, expand=True)

            temp_path = f"temp_page_{p}.png"
            rotated_image.save(temp_path)
            doc_img = Image(temp_path, detect_rotation=False)
            output_excel_portrait = os.path.splitext(os.path.basename(file_path))[0] + f"_portrait_page_{p}.xlsx"

            try:
                doc_img.to_xlsx(dest=output_excel_portrait,
                                ocr=ocr,
                                implicit_rows=implicit_rows,
                                implicit_columns=implicit_columns,
                                borderless_tables=borderless_tables,
                                min_confidence=50)
                print(f"✅ Success: Excel file created at {output_excel_portrait} for portrait page {p}.")
            except Exception as e:
                print(f"❌ An error occurred while processing portrait page {p} for {file_path}:", e)

            os.remove(temp_path)    


def process_image(file_path, ocr, implicit_rows, implicit_columns, borderless_tables):
    """Processes an image file (non-PDF)."""
    try:
        doc = Image(file_path, detect_rotation=False)
        output_excel = os.path.splitext(os.path.basename(file_path))[0] + ".xlsx"
        doc.to_xlsx(dest=output_excel,
                    ocr=ocr,
                    implicit_rows=implicit_rows,
                    implicit_columns=implicit_columns,
                    borderless_tables=borderless_tables,
                    min_confidence=50)
        print(f"✅ Success: Excel file created at {output_excel}")
    except Exception as e:
        print(f"❌ An error occurred while processing {file_path}:", e)

def process_file(file_path, ocr, implicit_rows, implicit_columns, borderless_tables):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        process_pdf(file_path, ocr, implicit_rows, implicit_columns, borderless_tables)
    elif ext in [".png", ".jpg", ".jpeg", ".bmp", ".tiff"]:
        process_image(file_path, ocr, implicit_rows, implicit_columns, borderless_tables)
    else:
        print(f"Unsupported file type: {file_path}. Skipping...")

def main():
    file_paths = input("Enter the paths to your PDF/Image files, separated by commas: ").strip().split(",")
    file_paths = [fp.strip() for fp in file_paths if fp.strip()]
    
    if not file_paths:
        print("No valid files provided. Exiting...")
        sys.exit(1)
    
    borderless_tables = input("Do the documents contain borderless tables? (Y/N): ").strip().lower() == "y"
    implicit_rows = input("Do the documents contain implicit rows? (Y/N): ").strip().lower() == "y"
    implicit_columns = input("Do the documents contain implicit columns? (Y/N): ").strip().lower() == "y"
    
    ocr = TesseractOCR(n_threads=1, lang="eng")
    
    for file_path in file_paths:
        process_file(file_path, ocr, implicit_rows, implicit_columns, borderless_tables)

if __name__ == "__main__":
    main()
