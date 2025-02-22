import os
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pdfplumber
from typing import Optional
from pdf2docx import Converter


class PDFConverter:
    def __init__(self):
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s: %(message)s",
            handlers=[
                logging.FileHandler("pdf_conversion.log"),
                logging.StreamHandler(),
            ],
        )
        self.logger = logging.getLogger(__name__)
        self.root = tk.Tk()
        self.root.withdraw()

    def extract_table_from_page(self, page) -> Optional[list]:
        try:
            table = page.extract_table()
            return table if table and len(table) > 0 else None
        except Exception as e:
            self.logger.warning(f"Could not extract table from page: {e}")
            return None

    def convert_pdf_to_excel(self, pdf_path: str, excel_path: str) -> bool:
        try:
            if not os.path.exists(pdf_path):
                self.logger.error(f"PDF file not found: {pdf_path}")
                return False

            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                print(f"Processing {os.path.basename(pdf_path)} ({total_pages} pages)")

                all_tables = []
                for page_number, page in enumerate(pdf.pages, start=1):
                    table = self.extract_table_from_page(page)

                    if table:
                        if not all_tables:
                            all_tables = table
                        else:
                            all_tables.extend(table[1:])

                    progress = (page_number / total_pages) * 100
                    print(
                        f"\rProcessing page {page_number}/{total_pages} ({progress:.2f}%)",
                        end="",
                        flush=True,
                    )

                print()

                if not all_tables or len(all_tables) <= 1:
                    self.logger.warning("No valid table data found in PDF")
                    return False

                df = pd.DataFrame(
                    all_tables[1:],
                    columns=all_tables[0] if len(all_tables[0]) > 1 else None,
                )

                for col in df.columns:
                    try:
                        df[col] = pd.to_numeric(df[col], errors="ignore")
                    except ValueError as e:
                        self.logger.warning(f"Column '{col}' could not be converted: {e}")

                os.makedirs(os.path.dirname(excel_path), exist_ok=True)
                df.to_excel(excel_path, index=False)
                self.logger.info(f"Successfully converted to: {excel_path}")
                return True

        except Exception as e:
            self.logger.error(f"Conversion failed: {e}")
            return False

    def convert_pdf_to_word(self, pdf_path: str, word_path: str) -> bool:
        try:
            if not os.path.exists(pdf_path):
                self.logger.error(f"PDF file not found: {pdf_path}")
                return False

            cv = Converter(pdf_path)
            cv.convert(word_path, start=0, end=None)
            cv.close()
            self.logger.info(f"Successfully converted to: {word_path}")
            return True
        except Exception as e:
            self.logger.error(f"Word conversion failed: {e}")
            return False

    def select_input_and_output(self):
        conversion_type = None
        while conversion_type not in ["excel", "word"]:
            conversion_type = input("Convert PDF to Excel or Word? (excel/word): ").strip().lower()

        selection = None
        while selection not in ["file", "folder"]:
            selection = input("Convert a single file or a folder? (file/folder): ").strip().lower()

        if selection == "file":
            input_path = filedialog.askopenfilename(
                title="Select a PDF File",
                filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            )
        else:
            input_path = filedialog.askdirectory(title="Select a Folder Containing PDFs")

        if not input_path:
            self.logger.info("No input selected. Exiting.")
            return

        output_folder = filedialog.askdirectory(title="Select Output Folder")
        if not output_folder:
            self.logger.info("No output folder selected. Exiting.")
            return

        self.run_conversion(input_path, output_folder, conversion_type)

    def run_conversion(self, input_path, output_folder, conversion_type):
        if os.path.isdir(input_path):
            for root_dir, _, files in os.walk(input_path):
                for file_name in files:
                    if file_name.lower().endswith(".pdf"):
                        pdf_path = os.path.join(root_dir, file_name)
                        relative_path = os.path.relpath(pdf_path, input_path)
                        output_path = os.path.join(
                            output_folder,
                            os.path.splitext(relative_path)[0] + (".xlsx" if conversion_type == "excel" else ".docx"),
                        )
                        print(f"Converting {file_name}...")
                        success = (
                            self.convert_pdf_to_excel(pdf_path, output_path)
                            if conversion_type == "excel"
                            else self.convert_pdf_to_word(pdf_path, output_path)
                        )
                        if success:
                            print(f"Successfully converted {file_name}!")
                        else:
                            print(f"Failed to convert {file_name}.")
        else:
            file_name = os.path.basename(input_path)
            output_path = os.path.join(output_folder, os.path.splitext(file_name)[0] + (".xlsx" if conversion_type == "excel" else ".docx"))
            print(f"Converting {file_name}...")
            success = (
                self.convert_pdf_to_excel(input_path, output_path)
                if conversion_type == "excel"
                else self.convert_pdf_to_word(input_path, output_path)
            )
            if success:
                print(f"Successfully converted {file_name}!")
            else:
                print(f"Failed to convert {file_name}.")

        print(f"All files processed. Output stored in: {output_folder}")


def main():
    converter = PDFConverter()
    converter.select_input_and_output()


if __name__ == "__main__":
    main()

