from __future__ import annotations
import argparse
import re
from pathlib import Path
from datetime import datetime, date
from typing import List, Dict
from openpyxl import load_workbook
from openpyxl.styles import numbers, Alignment
from openpyxl.utils import get_column_letter

def read_pdf_text(pdf_path: Path) -> str:
    try:
        import pdfplumber
        text = ""
        with pdfplumber.open(str(pdf_path)) as pdf:
            for p in pdf.pages:
                text += (p.extract_text(x_tolerance=1, y_tolerance=1) or "") + "\n"
        if text.strip():
            return text
    except Exception:
        pass
    try:
        from PyPDF2 import PdfReader
        rd = PdfReader(str(pdf_path))
        return "\n".join([(pg.extract_text() or "") for pg in rd.pages])
    except Exception:
        return ""

def parse_chase_transactions_full(text: str) -> List[Dict]:
    sect_deposits     = re.compile(r"^\s*DEPOSITS\s+AND\s+ADDITIONS\b", re.I)
    sect_withdrawals  = re.compile(r"^\s*ELECTRONIC\s+WITHDRAWALS\b", re.I)
    sect_fees         = re.compile(r"^\s*FEES\b", re.I)
    sect_daily_bal    = re.compile(r"DAILY\s+ENDING\s+BALANCE", re.I)
    line_total        = re.compile(r"^\s*TOTAL\b", re.I)

    date_re    = r"(0[1-9]|1[0-2])/(0[1-9]|[12]\d|3[01])(?:/(20\d{2}))?"
    amount_re  = r"[\(\-]?\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})[\)]?"
    tx_header  = re.compile(rf"^\s*{date_re}(?:\s+{date_re})?\s+(.*?)\s+({amount_re})\s*$")

    def norm_amount(s: str) -> float | None:
        s1 = s.replace("$", "").replace(",", "").strip()
        neg = False
        if s1.startswith("(") and s1.endswith(")"):
            neg = True
            s1 = s1[1:-1]
        if s1.startswith("-"):
            neg = True
            s1 = s1[1:]
        try:
            v = float(s1)
        except ValueError:
            return None
        return -v if neg else v

    transactions: List[Dict] = []
    section = None
    current = None

    for raw in text.splitlines():
        ln = raw.strip()

        if sect_daily_bal.search(ln):
            section = None; current=None; continue
        if sect_deposits.search(ln):
            section = "deposit"; current=None; continue
        if sect_withdrawals.search(ln):
            section = "withdrawal"; current=None; continue
        if sect_fees.search(ln):
            section = "fee"; current=None; continue

        if not section:
            continue
        if line_total.search(ln):
            current = None; continue

        m = tx_header.match(ln)
        if m:
            if current is not None:
                current["Description"] = "\n".join(current["desc_lines"]).strip()
                transactions.append(current)

            g = m.groups()
            d1m, d1d, d1y = g[0], g[1], g[2]
            d2m, d2d, d2y = g[3], g[4], g[5]
            desc = g[6]
            amt_str = g[7]

            year  = int((d2y or d1y) or "2024")
            month = int((d2m or d1m))
            day   = int((d2d or d1d))
            try:
                tx_date = datetime(year, month, day).date()
            except ValueError:
                current = None; continue

            amt = norm_amount(amt_str)
            if amt is None:
                current = None; continue

            current = {"Date": tx_date, "Amount": amt, "Section": section, "desc_lines": [desc]}
            continue

        if current is not None:
            current["desc_lines"].append(ln)

    if current is not None:
        current["Description"] = "\n".join(current["desc_lines"]).strip()
        transactions.append(current)

    seen = set()
    uniq = []
    for t in transactions:
        key = (t["Date"].isoformat(), round(t["Amount"],2), t["Section"], t.get("Description",""))
        if key in seen:
            continue
        seen.add(key)
        uniq.append(t)
    return uniq

def find_headers(ws):
    for r in range(1, 100):
        vals = [
            (str(ws.cell(row=r, column=c).value).strip().lower()
             if ws.cell(row=r, column=c).value is not None else "")
            for c in range(1, 41)
        ]
        date_col = next((i+1 for i,v in enumerate(vals) if v in ("date","transaction date","posting date")), None)
        check_col = next((i+1 for i,v in enumerate(vals) if v in ("check #","check#","check no","check number")), None)
        amt_col = next((i+1 for i,v in enumerate(vals) if v in ("amount","amt","amount (usd)","debit/credit")), None)
        if date_col and check_col and amt_col:
            return r, date_col, check_col, amt_col
    return 1, 1, 2, 4

def main():
    parser = argparse.ArgumentParser(description="Book-keeping: build/update register (STRICT inputs).")
    parser.add_argument("--directory", required=True, help="Folder named YYYY-MM containing 'Corp Registers_.xlsx' and a '*2590*.pdf'")
    args = parser.parse_args()

    target = Path(args.directory)
    if not target.exists() or not target.is_dir():
        raise FileNotFoundError(f"Directory not found: {target}")

    # Parse folder name YYYY-MM
    month_tag = target.name
    m = re.fullmatch(r"(20\d{2})-(0[1-9]|1[0-2])", month_tag)
    if not m:
        raise ValueError(f"Folder name must be YYYY-MM (got: {month_tag})")
    year = int(m.group(1)); month = int(m.group(2))
    yyyymm = f"{year}{month:02d}"
    first_day = datetime(year, month, 1)

    # STRICT Excel template
    template = target / "Corp Registers_.xlsx"
    if not template.exists():
        raise FileNotFoundError("Excel template 'Corp Registers_.xlsx' not found in the target folder.")

    # STRICT PDF selection
    candidates = sorted(target.glob("*2590*.pdf"))
    if not candidates:
        raise FileNotFoundError("No PDF containing '2590' found in the target folder.")
    pdf_path = candidates[0]

    # Copy template to new name
    dest_excel = target / f"Corp Registers_{yyyymm}.xlsx"
    if template.resolve() != dest_excel.resolve():
        dest_excel.write_bytes(template.read_bytes())

    # Load workbook
    wb = load_workbook(dest_excel)
    if "Check Register-Corp" not in wb.sheetnames:
        raise RuntimeError("Sheet 'Check Register-Corp' not found in template.")
    ws = wb["Check Register-Corp"]

    # Update B9
    ws["B9"].value = first_day

    header_row, date_col, check_col, amt_col = find_headers(ws)
    expense_col = 4
    deposit_col = 5

    if header_row < ws.max_row:
        ws.delete_rows(header_row+1, ws.max_row - header_row)

    # Parse PDF & order
    text = read_pdf_text(pdf_path)
    transactions = parse_chase_transactions_full(text)
    ordered_tx = []
    for sec in ("deposit","withdrawal","fee"):
        ordered_tx.extend([t for t in transactions if t["Section"] == sec])

    # Write rows
    for r_i, t in enumerate(ordered_tx, start=header_row+1):
        ws.cell(row=r_i, column=date_col).value = t["Date"]
        ws.cell(row=r_i, column=amt_col).value = abs(float(t["Amount"]))

        full_desc = (t.get("Description") or "").replace("\r", " ").replace("\n", " ").replace("\t", " ")
        full_desc = re.sub(r"\s+", " ", full_desc).strip()
        ws.cell(row=r_i, column=check_col).value = full_desc
        ws.cell(row=r_i, column=check_col).alignment = Alignment(wrap_text=True)

        dlow = full_desc.lower()
        if t["Section"] == "deposit":
            if "transfer" in dlow and "0639" in dlow:
                ws.cell(row=r_i, column=deposit_col).value = "Promissory Note to bond holder Ting Wang"
            elif "e*trade" in dlow:
                ws.cell(row=r_i, column=deposit_col).value = "Transfer money from E*Trade Brokerage Account"
            elif "gainsystems" in dlow:
                ws.cell(row=r_i, column=deposit_col).value = "Income by Consulting with GAINSystems"
            elif "allegis group" in dlow:
                ws.cell(row=r_i, column=deposit_col).value = "Income by Consulting with JP Morgan Chase"
            else:
                ws.cell(row=r_i, column=deposit_col).value = "Income"
        else:
            value = "Payment"
            if "e*trade" in dlow:
                value = "Transfer money to E*Trade Brokerage Account"
            elif "transfer" in dlow and "0639" in dlow:
                value = "Return money to bond holder Ting Wang"
            ws.cell(row=r_i, column=expense_col).value = value

    # Format
    for rr in range(header_row+1, ws.max_row+1):
        dcell = ws.cell(row=rr, column=date_col)
        if isinstance(dcell.value, (datetime, date)):
            dcell.number_format = "MM/dd/yyyy"
        acell = ws.cell(row=rr, column=amt_col)
        if isinstance(acell.value, (int,float)):
            acell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

    # Hide extra sheets
    for name in wb.sheetnames:
        lname = name.strip().lower()
        if lname in ("information to provide to ta", "typical exp in trading business"):
            wb[name].sheet_state = "hidden"

    # Autofit
    for c in range(1, ws.max_column + 1):
        max_len = 0
        col_letter = get_column_letter(c)
        for rr in range(1, ws.max_row + 1):
            v = ws.cell(row=rr, column=c).value
            if v is None:
                continue
            s = str(v)
            if len(s) > max_len:
                max_len = len(s)
        ws.column_dimensions[col_letter].width = min(max_len + 2, 80)

    wb.save(dest_excel)
    print(f"Workbook updated: {dest_excel}")

if __name__ == "__main__":
    main()