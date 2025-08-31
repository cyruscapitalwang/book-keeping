from __future__ import annotations
import argparse
import re
from pathlib import Path
from datetime import datetime, date
from typing import List, Dict, Tuple
from math import isclose

from openpyxl import load_workbook
from openpyxl.styles import numbers, Alignment
from openpyxl.utils import get_column_letter


# ==============================
# PDF text extraction
# ==============================
def read_pdf_text(pdf_path: Path) -> str:
    """
    Extract text from PDF using pdfplumber if available, else PyPDF2.
    """
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


# ==============================
# Helpers / cleaners
# ==============================
_AMOUNT_RE = r"[\(\-]?\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})[\)]?"

def _norm_amount(s: str) -> float | None:
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

def strip_after_amount(line: str) -> str:
    """
    If a line contains an amount and then extra trailing text (e.g., '... 10,000.00 Page 2 of 6 ...'),
    return only up to (and including) the FIRST amount. Otherwise return the line unchanged.
    """
    m = re.search(_AMOUNT_RE, line)
    if not m:
        return line
    return line[: m.end()].rstrip()

def clean_description(desc: str) -> str:
    """
    Remove common PDF artifacts (page footers, long numeric blobs), collapse whitespace.
    """
    desc = re.sub(r"\bPage\s+\d+\s+of\s+\d+\b", "", desc, flags=re.I)
    desc = re.sub(r"\d{6,}", "", desc)   # very long number garbage
    desc = re.sub(r"\s+", " ", desc)
    return desc.strip()


# ==============================
# Parser with PDF totals + reconciliation
# ==============================
def parse_chase_transactions_full(
    text: str, verbose: bool = False
) -> Tuple[List[Dict], Dict[str, float], Dict[str, List[str]]]:
    """
    Parse statement text into transactions with full multi-line descriptions,
    collect PDF section totals (Deposits/Withdrawals/Fees), and capture
    unparsed lines per section for reconciliation.

    Supports:
      - strict headers (date[s] at start, amount at end)
      - two-line headers (amount on next line)
      - loose headers (date anywhere, amount at end)
      - continuation lines (to keep Trace# etc.)
      - NO deduplication (distinct lines kept)
    """
    sect_deposits     = re.compile(r"^\s*DEPOSITS\s+AND\s+ADDITIONS\b", re.I)
    sect_withdrawals  = re.compile(r"^\s*ELECTRONIC\s+WITHDRAWALS\b", re.I)
    sect_fees         = re.compile(r"^\s*FEES\b", re.I)
    sect_daily_bal    = re.compile(r"DAILY\s+ENDING\s+BALANCE", re.I)
    line_total        = re.compile(r"^\s*TOTAL\b", re.I)

    date_re    = r"(0[1-9]|1[0-2])/(0[1-9]|[12]\d|3[01])(?:/(20\d{2}))?"
    amount_re  = _AMOUNT_RE

    # Header types
    tx_header_strict = re.compile(
        rf"^\s*{date_re}(?:\s+{date_re})?\s+(.*?)\s+({amount_re})\s*$"
    )  # date(s) at start, amount at end
    tx_header_no_amt = re.compile(
        rf"^\s*{date_re}(?:\s+{date_re})?\s+(?!.*{amount_re}\s*$)(.+)$"
    )  # date(s) at start, NO amount (2-line)
    tx_header_loose  = re.compile(rf".*?{date_re}.*?({amount_re})\s*$")  # date anywhere, amount at end

    trailing_amt = re.compile(rf"({amount_re})\s*$")

    transactions: List[Dict] = []
    pdf_totals = {"deposit": 0.0, "withdrawal": 0.0, "fee": 0.0}
    unparsed: Dict[str, List[str]] = {"deposit": [], "withdrawal": [], "fee": []}

    section = None
    current: Dict | None = None           # accumulating {Date, Amount, Section, desc_lines:[...]}
    pending_header: Dict | None = None    # for two-line header (amount on next line)

    def finalize_current():
        nonlocal current
        if current is not None:
            current["Description"] = "\n".join(current["desc_lines"]).strip()
            transactions.append(current)
            current = None

    for raw in text.splitlines():
        ln = raw.strip()

        # Stop at Daily Ending Balance
        if sect_daily_bal.search(ln):
            finalize_current()
            section = None
            pending_header = None
            continue

        # Section headers
        if sect_deposits.search(ln):
            finalize_current()
            section = "deposit"; pending_header = None; continue
        if sect_withdrawals.search(ln):
            finalize_current()
            section = "withdrawal"; pending_header = None; continue
        if sect_fees.search(ln):
            finalize_current()
            section = "fee"; pending_header = None; continue

        if not section:
            continue

        # Section TOTAL: record PDF total and close any open tx
        if line_total.search(ln):
            m_amt = trailing_amt.search(ln)
            if m_amt:
                val = _norm_amount(m_amt.group(1))
                if val is not None:
                    pdf_totals[section] = abs(val)
            finalize_current()
            pending_header = None
            continue

        # --- Two-line header: we already have date/desc, expect amount on next line ---
        if pending_header is not None:
            ln_clipped = strip_after_amount(ln)
            m_amt_first = re.search(amount_re, ln_clipped)  # FIRST amount
            if m_amt_first:
                amt = _norm_amount(m_amt_first.group(0))
                if amt is not None:
                    current = {
                        "Date": pending_header["date"],
                        "Amount": amt,
                        "Section": pending_header["section"],
                        "desc_lines": [pending_header["desc"], ln_clipped],
                    }
                    pending_header = None
                    # keep accumulating continuation lines in subsequent iterations
                    continue
            # not an amount line — treat as continuation of description
            pending_header["desc"] += " " + ln_clipped
            continue

        # Strict header: start a new current (do NOT finalize yet)
        m = tx_header_strict.match(ln)
        if m:
            finalize_current()
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
                unparsed[section].append(ln); continue
            amt = _norm_amount(amt_str)
            if amt is None:
                unparsed[section].append(ln); continue
            current = {"Date": tx_date, "Amount": amt, "Section": section, "desc_lines": [desc]}
            continue

        # Date at start but no amount → two-line header begins
        m_no = tx_header_no_amt.match(ln)
        if m_no:
            d1m, d1d, d1y = m_no.group(1), m_no.group(2), m_no.group(3)
            d2m, d2d, d2y = m_no.group(4), m_no.group(5), m_no.group(6)
            desc_only = m_no.group(7)
            year  = int((d2y or d1y) or "2024")
            month = int((d2m or d1m))
            day   = int((d2d or d1d))
            try:
                tx_date = datetime(year, month, day).date()
            except ValueError:
                unparsed[section].append(ln); continue
            finalize_current()
            pending_header = {"date": tx_date, "desc": desc_only, "section": section}
            continue

        # Loose header (date anywhere, amount at end): start a new current
        m2 = tx_header_loose.match(ln)
        if m2:
            ln_clipped = strip_after_amount(ln)
            m_amt_first = re.search(amount_re, ln_clipped)
            dmatch = re.search(date_re, ln_clipped)
            if not (m_amt_first and dmatch):
                unparsed[section].append(ln); continue
            amt = _norm_amount(m_amt_first.group(0))
            if amt is None:
                unparsed[section].append(ln); continue
            mm, dd, yy = int(dmatch.group(1)), int(dmatch.group(2)), dmatch.group(3)
            year = int(yy or "2024")
            try:
                tx_date = datetime(year, mm, dd).date()
            except ValueError:
                unparsed[section].append(ln); continue
            finalize_current()
            current = {"Date": tx_date, "Amount": amt, "Section": section, "desc_lines": [ln_clipped]}
            continue

        # Continuation line for an open current?
        if current is not None:
            current["desc_lines"].append(strip_after_amount(ln))
            continue

        # Otherwise, we couldn't place this line
        unparsed[section].append(ln)

    # finalize at EOF
    finalize_current()
    if pending_header is not None:
        # never found the amount line
        unparsed[pending_header["section"]].append(pending_header["desc"])

    # NOTE: no dedup — keep both same-date/same-amount items if they appear separately
    return transactions, pdf_totals, unparsed


# ==============================
# Excel helpers
# ==============================
def find_headers(ws):
    """
    Locate header row and key columns: Date, Check #, Amount.
    """
    for r in range(1, 100):
        vals = [
            (str(ws.cell(row=r, column=c).value).strip().lower()
             if ws.cell(row=r, column=c).value is not None else "")
            for c in range(1, 41)
        ]
        date_col = next((i + 1 for i, v in enumerate(vals) if v in ("date", "transaction date", "posting date")), None)
        check_col = next((i + 1 for i, v in enumerate(vals) if v in ("check #", "check#", "check no", "check number")), None)
        amt_col = next((i + 1 for i, v in enumerate(vals) if v in ("amount", "amt", "amount (usd)", "debit/credit")), None)
        if date_col and check_col and amt_col:
            return r, date_col, check_col, amt_col
    return 1, 1, 2, 4


def compute_section_sums(transactions: List[Dict]) -> Dict[str, float]:
    sums = {"deposit": 0.0, "withdrawal": 0.0, "fee": 0.0}
    for t in transactions:
        sec = t["Section"]
        amt = abs(float(t["Amount"]))
        sums[sec] += amt
    for k in list(sums.keys()):
        sums[k] = round(sums[k], 2)
    return sums


# ==============================
# Main
# ==============================
def main():
    parser = argparse.ArgumentParser(
        description="Book-keeping: build/update register (STRICT inputs). Template is read from project root."
    )
    parser.add_argument(
        "--directory",
        required=True,
        help="Target folder named YYYY-MM containing a '*2590*.pdf'. Output Excel will be created here.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Simulate; do not write Excel")
    parser.add_argument("--verbose", action="store_true", help="Print detailed steps")
    parser.add_argument("--force", action="store_true", help="Write Excel even if section totals mismatch")
    args = parser.parse_args()

    target = Path(args.directory)
    if not target.exists() or not target.is_dir():
        raise FileNotFoundError(f"Directory not found: {target}")

    # Parse folder name YYYY-MM
    month_tag = target.name
    m = re.fullmatch(r"(20\d{2})-(0[1-9]|1[0-2])", month_tag)
    if not m:
        raise ValueError(f"Folder name must be YYYY-MM (got: {month_tag})")
    year = int(m.group(1))
    month = int(m.group(2))
    yyyymm = f"{year}{month:02d}"
    first_day = datetime(year, month, 1)

    # Template ALWAYS from project root (…/src/book_keeping/cli.py -> project_root)
    project_root = Path(__file__).resolve().parent.parent.parent
    template = project_root / "Corp Registers_.xlsx"
    if not template.exists():
        raise FileNotFoundError(
            f"Excel template not found at project root: {template}\n"
            "Place 'Corp Registers_.xlsx' next to your pyproject.toml."
        )

    # Strict PDF selection in the target folder
    candidates = sorted(target.glob("*2590*.pdf"))
    if not candidates:
        raise FileNotFoundError("No PDF containing '2590' found in the target folder.")
    pdf_path = candidates[0]

    if args.verbose:
        print(f"[INFO] Project root: {project_root}")
        print(f"[INFO] Month tag: {month_tag} -> {yyyymm}")
        print(f"[INFO] Template (project root): {template.name}")
        print(f"[INFO] PDF (target folder): {pdf_path.name}")

    # Parse the PDF
    text = read_pdf_text(pdf_path)
    transactions, pdf_totals, unparsed = parse_chase_transactions_full(text, verbose=args.verbose)

    # Preserve original order: deposit -> withdrawal -> fee
    ordered_tx: List[Dict] = []
    for sec in ("deposit", "withdrawal", "fee"):
        ordered_tx.extend([t for t in transactions if t["Section"] == sec])

    # Compute totals from parsed data
    sums_from_data = compute_section_sums(ordered_tx)

    # Print PDF totals and computed totals
    print("PDF Section Totals:")
    print(f"  Deposits   : ${pdf_totals['deposit']:.2f}")
    print(f"  Withdrawals: ${pdf_totals['withdrawal']:.2f}")
    print(f"  Fees       : ${pdf_totals['fee']:.2f}")

    print("Computed Totals (from data to be written):")
    print(f"  Deposits   : ${sums_from_data['deposit']:.2f}")
    print(f"  Withdrawals: ${sums_from_data['withdrawal']:.2f}")
    print(f"  Fees       : ${sums_from_data['fee']:.2f}")

    # Validate totals match to the cent; print reconciliation on mismatch
    mismatches = []
    for sec in ("deposit", "withdrawal", "fee"):
        if not isclose(pdf_totals[sec], sums_from_data[sec], abs_tol=0.01):
            mismatches.append(sec)

    if mismatches:
        for sec in mismatches:
            print(f"\n=== RECONCILIATION NEEDED: {sec.upper()} ===")
            print(f"PDF total: ${pdf_totals[sec]:.2f} | Parsed total: ${sums_from_data[sec]:.2f}")

            # Show a small sample of parsed rows for this section
            parsed_sec = [t for t in ordered_tx if t["Section"] == sec]
            shown = parsed_sec[:10]
            print(f"Parsed {sec} sample ({len(shown)} shown of {len(parsed_sec)}):")
            for i, t in enumerate(shown, 1):
                first_line = (t.get("Description") or "").splitlines()[0]
                print(f"  [{i:02d}] {t['Date']}  ${abs(float(t['Amount'])):,.2f}  {first_line[:120]}")

            # Show unparsed lines we saw in this section
            if unparsed.get(sec):
                print(f"\nUnparsed lines in {sec} section ({len(unparsed[sec])}):")
                for i, ln in enumerate(unparsed[sec], 1):
                    print(f"  - {ln}")

        if not args.force:
            raise AssertionError(
                "Section totals mismatch between PDF and parsed data. "
                "See reconciliation above. Use --force to proceed anyway."
            )
        else:
            print("[WARN] Totals mismatch, but --force specified; continuing.")

    if args.dry_run:
        print("[DRY-RUN] Totals check passed (or forced). No files written.")
        return

    # Proceed to write Excel in the target folder
    dest_excel = target / f"Corp Registers_{yyyymm}.xlsx"
    if args.verbose:
        print(f"[INFO] Copying template -> {dest_excel.name}")
    if template.resolve() != dest_excel.resolve():
        dest_excel.write_bytes(template.read_bytes())

    wb = load_workbook(dest_excel)
    if "Check Register-Corp" not in wb.sheetnames:
        raise RuntimeError("Sheet 'Check Register-Corp' not found in template.")
    ws = wb["Check Register-Corp"]

    # Update B9 to first day of month
    ws["B9"].value = first_day

    # Identify columns
    header_row, date_col, check_col, amt_col = find_headers(ws)
    expense_col = 4  # D
    deposit_col = 5  # E

    # Clear existing rows below header
    if header_row < ws.max_row:
        ws.delete_rows(header_row + 1, ws.max_row - header_row)

    # Write rows
    for r_i, t in enumerate(ordered_tx, start=header_row + 1):
        ws.cell(row=r_i, column=date_col).value = t["Date"]
        ws.cell(row=r_i, column=amt_col).value = abs(float(t["Amount"]))

        # Full paragraph → single line with wrapping + cleanup
        full_desc = (t.get("Description") or "").replace("\r", " ").replace("\n", " ").replace("\t", " ")
        full_desc = clean_description(full_desc)
        ws.cell(row=r_i, column=check_col).value = full_desc
        ws.cell(row=r_i, column=check_col).alignment = Alignment(wrap_text=True)

        # Deposit/Expense classification rules (specific -> general)
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
            elif "quinnox" in dlow:
                ws.cell(row=r_i, column=deposit_col).value = "Income by Consulting with US Bank"
            else:
                ws.cell(row=r_i, column=deposit_col).value = "Income"

        else:
            # Withdrawal / Fee (payments)
            value = "Payment"
            if "e*trade" in dlow:
                value = "Transfer money to E*Trade Brokerage Account"
            elif "transfer" in dlow and "0639" in dlow:
                value = "Return money to bond holder Ting Wang"
            elif "gusto" in dlow:
                value = "Payroll professional service fee"
            elif (("u.s. bank" in dlow) or ("us bank" in dlow)) and (("lse pmts" in dlow) or ("lease" in dlow)):
                value = "Car lease payment"
            elif ("tesla" in dlow) or ("telsa" in dlow):  # handle common misspelling
                value = "Car wireless subscription payment"
            ws.cell(row=r_i, column=expense_col).value = value

    # Format date & amount columns
    for rr in range(header_row + 1, ws.max_row + 1):
        dcell = ws.cell(row=rr, column=date_col)
        if isinstance(dcell.value, (datetime, date)):
            dcell.number_format = "MM/dd/yyyy"
        acell = ws.cell(row=rr, column=amt_col)
        if isinstance(acell.value, (int, float)):
            acell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

    # Hide extra sheets if present
    for name in wb.sheetnames:
        lname = name.strip().lower()
        if lname in ("information to provide to ta", "typical exp in trading business"):
            wb[name].sheet_state = "hidden"

    # Autofit columns (approximate)
    for c in range(1, ws.max_column + 1):
        max_len = 0
        col_letter = get_column_letter(c)
        for rr in range(1, ws.max_row + 1):
            v = ws.cell(row=rr, column=c).value
            if v is None:
                continue
            s = str(v)
            max_len = max(max_len, len(s))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 80)

    wb.save(dest_excel)
    print(f"Workbook updated: {dest_excel}")


if __name__ == "__main__":
    main()
