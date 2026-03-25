import tkinter as tk
from tkinter import filedialog
from parsers.cosco_parser import parse_cosco_pdf
from excel.excel_writer_test import update_excel_rates_test
from openpyxl import load_workbook

def main():
    pdf_path = "data/COSCO1.pdf"

    rates = parse_cosco_pdf(pdf_path)

    print("==== PARSED RESULT ====")

    for r in rates:
        print(r)

    print(f"\nTotal rows: {len(rates)}")

    excel_path = "data/Cheatsheet.xlsx"

    update_excel_rates_test(excel_path, rates)

    print("Excel updated successfully!")

    

    wb = load_workbook(excel_path)
    sheet = wb.active

    print("\n=== CHECK EXCEL ===")
    print("H5:", sheet["H5"].value)
    print("H6:", sheet["H6"].value)
    print("H26:", sheet["H26"].value)
    print("H27:", sheet["H27"].value)
    print("AA5:", sheet["AA5"].value)
    print("AA6:", sheet["AA6"].value)

    wb.close()

if __name__ == "__main__":
    main()