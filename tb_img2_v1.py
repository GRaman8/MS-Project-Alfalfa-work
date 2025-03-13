import os 
from img2table.document import PDF, Image
from img2table.ocr import TesseractOCR
import sys
import cv2
print(hasattr(cv2.ximgproc, 'niBlackThreshold'))


def main():
    file_path = input("Enter the full path to your document (PDF or Image): ").strip()

    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        sys.exit(1)

    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        doc = PDF(file_path, pdf_text_extraction = True)
    elif ext in [".png", ".jpg", ".jpeg", ".bmp", ".tiff"]:
        doc = Image(file_path, detect_rotation=False)
    else:
        print("Unsupported file type! Please provide a PDF or an image (png, jpg, jpeg, bmp, tiff.)")
        sys.exit(1)

    borderless_input = input("does the document contain borderless tables? (Y/N):").strip().lower()
    borderless_tables = True if borderless_input == "y" else False

    implicit_rows_input = input("Does the document contain implicit rows? (Y/N):").strip().lower()
    implicit_rows = True if implicit_rows_input == "y" else False

    implicit_columns_input= input("Does the document contain implicit columns? (Y/N):").strip().lower()
    implicit_columns = True if implicit_columns_input == "y" else False

    ocr = TesseractOCR(n_threads=1, lang ="eng")

    output_excel = input("Enter the desired output Excel file name (e.g, output.xlsx):").strip()
    if not output_excel.endswith(".xlsx"):
        output_excel += ".xlsx"

    try:
        doc.to_xlsx(dest = output_excel,
                    ocr=ocr,
                    implicit_rows= implicit_rows,
                    implicit_columns=implicit_columns,
                    borderless_tables= borderless_tables,
                    min_confidence=50)
        print(f"Success: Excel file created at {output_excel}")
    except Exception as e:
        print("An error occurred during table extraction:", e)
        sys.exit(1)

if __name__ == "__main__":
    main()        