import tkinter as tk
from tkinter import filedialog
from parsers.cosco_parser import parse_cosco_pdf, parse_wharfage, parse_ets
from excel.excel_writer_test import update_excel_rates_test
from openpyxl import load_workbook

def main():
    pdf_path = "data/COSCO3.pdf"

    rates = parse_cosco_pdf(pdf_path)

    print("==== PARSED RESULT ====")

    for r in rates:
        print(r)

    print(f"\nTotal rows: {len(rates)}")

    whf_rates = parse_wharfage(pdf_path)

    print("\n==== WHF PARSE RESULT ====")
    for city, rate_dict in whf_rates.items():
        print(f"\n{city}")
        print(f"  20   : {rate_dict['20']}")
        print(f"  40   : {rate_dict['40']}")
        print(f"  40HC : {rate_dict['40HC']}")

    
    ets_value = parse_ets(pdf_path)

    print("\n==== ETS GC ====")
    print(ets_value)

    excel_path = "data/Cheatsheet2.xlsx"

    update_excel_rates_test(excel_path, rates, whf_rates)

    print("Excel updated successfully!")

    

    wb = load_workbook(excel_path)
    sheet = wb.active

    print("\n=== CHECK EXCEL ===")
    print("H5:", sheet["H5"].value)
    print("H6:", sheet["H6"].value)
    print("H26:", sheet["H26"].value)
    print("H27:", sheet["H27"].value)

    wb.close()

if __name__ == "__main__":
    main()