import xlwings as xw

def update_excel_rates(excel_path, rates):

    wb = xw.Book(excel_path)
    sheet = wb.sheets.active

    # --- Step 1: copy OLD rates ---
    for row in range(5, 98):
        sheet[f"AA{row}"] = sheet[f"H{row}"].value
        sheet[f"AB{row}"] = sheet[f"I{row}"].value
        sheet[f"AC{row}"] = sheet[f"J{row}"].value


    # --- Step 2: group ---
    bremerhaven_rates = [r for r in rates if r["POL"] == "BREMERHAVEN"]
    hamburg_rates = [r for r in rates if r["POL"] == "HAMBURG"]

    # --- Step 3: write BREMERHAVEN ---
    row = 5
    for r in bremerhaven_rates:
        sheet.range(f"H{row}").value = r["20"]
        sheet.range(f"I{row}").value = r["40"]
        sheet.range(f"J{row}").value = r["40HC"]
        row += 1

    # --- Step 4: write HAMBURG ---
    row = 26
    for r in hamburg_rates:
        sheet.range(f"H{row}").value = r["20"]
        sheet.range(f"I{row}").value = r["40"]
        sheet.range(f"J{row}").value = r["40HC"]
        row += 1

    wb.save()
    wb.close()