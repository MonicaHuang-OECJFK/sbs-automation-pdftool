import pdfplumber
import re

def normalize_rate(rate_str):
    if not rate_str:
        return None

    # 移除 $
    rate_str = rate_str.replace("$", "")

    # 移除空格
    rate_str = rate_str.replace(" ", "")

    # 處理歐洲格式 1.650,00 → 1650.00
    rate_str = rate_str.replace(".", "").replace(",", ".")

    return int(float(rate_str))


def parse_cosco_pdf(pdf_path):

    rates = []

    with pdfplumber.open(pdf_path) as pdf:

        for page in pdf.pages:

            tables = page.extract_tables()

            if not tables:
                continue

            for table in tables:

                header = table[0]
                if not header:
                    continue

                header_text = " ".join(str(h) for h in header if h)

                if "Port of Discharge" not in header_text:
                    continue

                current_pol = None

                for row in table[1:]:

                    if not row:
                        continue

                    country = row[0]
                    pol_cell = row[1]
                    pod = row[2]
                    term = row[3]
                    rate20 = row[4]
                    rate40 = row[5]

                    # skip invalid row
                    if term != "CY/CY":
                        continue

                    # handle merged cell（POL 可能只出現一次）
                    if pol_cell:
                        current_pol = pol_cell

                    if not current_pol or not pod:
                        continue

                    # ✅ 👉 關鍵：展開多 POL
                    pol_list = [
                        p.strip().upper()
                        for p in str(current_pol).split(",")
                        if p.strip()
                    ]

                    pod = pod.strip().upper()

                    rate20 = normalize_rate(rate20)
                    rate40 = normalize_rate(rate40)

                    # ✅ 👉 一個 row → 多個 POL
                    for pol in pol_list:
                        rates.append({
                            "POL": pol,
                            "POD": pod,
                            "20": rate20,
                            "40": rate40,
                            "40HC": rate40
                        })

    return rates


def parse_wharfage(pdf_path):

    whf_rates = {}

    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"

    pattern = r"WHF.*?(Houston|New Orleans|Miami).*?(20|40).*?(USD\s*[\d.,]+)"

    matches = re.findall(pattern, text, re.IGNORECASE)
    
    print("\n[DEBUG] WHF MATCHES:")
    for m in matches:
        print(m)

    for city, size, price in matches:
        city = city.upper()
        price = normalize_rate_whf(price)

        if city not in whf_rates:
            whf_rates[city] = {"20": None, "40": None, "40HC": None}

        if size == "20":
            whf_rates[city]["20"] = price
        else:
            whf_rates[city]["40"] = price
            whf_rates[city]["40HC"] = price

    return whf_rates

def normalize_rate_whf(rate_str):
    if not rate_str:
        return None

    rate_str = rate_str.replace("USD", "")
    rate_str = rate_str.replace("$", "")
    rate_str = rate_str.replace(" ", "")

    # 歐洲格式轉換
    rate_str = rate_str.replace(".", "").replace(",", ".")

    return float(rate_str)

    
def parse_ets(pdf_path):

    import pdfplumber
    import re

    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"

    # 👉 先抓 surcharge 區塊
    block_match = re.search(
        r"Surcharges.*?(ETS\s*RF)",
        text,
        re.IGNORECASE | re.DOTALL
    )

    if block_match:
        block = block_match.group(0)

        print("\n[DEBUG SURCHARGE BLOCK]")
        print(block[:500])

        match = re.search(
            r"ETS\s*GC.*?EUR\s*([\d,]+)",
            block,
            re.IGNORECASE
        )

        if match:
            value = int(match.group(1).replace(",", ""))
            print("[DEBUG ETS GC FOUND]:", value)
            return value

    print("[DEBUG ETS GC NOT FOUND]")
    return None