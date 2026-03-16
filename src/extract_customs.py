"""
extract_customs.py
------------------
Hybrid extraction:
- pdfplumber x/y field map  → all C.VALU SUMMA + EX.PR numeric fields
- pdfplumber text regex     → port code, SB no, SB date, INV fields
- Claude vision             → Part II C.VAL DTLS + item table

Requirements:
    pip install pdfplumber pymupdf anthropic openpyxl
"""

import sys, re, json, base64
import pdfplumber
import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── 1. PART I EXTRACTION ─────────────────────────────────────────────────────
# The C.VALU SUMMA and D.EX.PR sections have a fixed layout across all
# Indian Customs EDI PDFs. Headers sit at y≈276 and y≈294; values sit
# exactly 9px below at y≈285 and y≈303. x ranges are constant.
# We simply read whichever numeric word falls within each field's x/y box.

_FIELD_MAP = [
    # (field_name,      x_min, x_max, val_y)
    # x ranges derived from header midpoints with generous padding.
    # Matching uses word midpoint (x0+x1)/2, so wide numbers still match.
    ("fob_value",          72,   143,   285),
    ("freight",           143,   194,   285),
    ("insurance",         194,   238,   285),   # first half of merged "3.INSURANC4.DISCOU"
    ("discount",          238,   286,   285),   # second half of merged header
    ("com",               286,   352,   285),
    ("dbk_claim",         352,   433,   285),
    ("igst_amt",          433,   510,   285),
    ("cess_amt",          510,   580,   285),
    ("deductions",         69,   155,   303),
    ("pc",                155,   248,   303),
    ("duty",              248,   285,   303),
    ("cess",              285,   352,   303),
    ("igst_value",        350,   420,   303),
    ("rodtep_amt",        420,   505,   303),
    ("rosctl_amt",        505,   580,   303),
]

def extract_part1(pdf_path: str) -> dict:
    out = {k: "" for k in [
        "port_code","sb_no","sb_date",
        "fob_value","freight","insurance","discount","com",
        "deductions","pc","duty","cess",
        "dbk_claim","igst_amt","cess_amt","igst_value","rodtep_amt","rosctl_amt",
        "inv_no","inv_amt","currency",
    ]}

    with pdfplumber.open(pdf_path) as pdf:
        words = pdf.pages[0].extract_words()
        text  = pdf.pages[0].extract_text() or ""

    # Numeric fields: match using word MIDPOINT so wide numbers don't fall outside range
    for field, x_min, x_max, val_y in _FIELD_MAP:
        hits = [
            w for w in words
            if abs(w["top"] - val_y) < 5                        # on the value row
            and x_min <= (w["x0"] + w["x1"]) / 2 < x_max       # midpoint in column
            and re.match(r'^[\d.,]+$', w["text"])               # is a number
        ]
        out[field] = hits[0]["text"] if hits else "0"

    # Text fields: regex on the extracted text lines
    for line in text.split("\n"):
        # "INDIAN CUSTOMS EDI SYSTEM INSGF6 9774136 04-MAY-23"
        m = re.search(r'(INSGF\d)\s+(\d{6,8})\s+(\d{2}-[A-Z]{3}-\d{2,4})', line)
        if m and not out["sb_no"]:
            out["port_code"] = m.group(1)
            out["sb_no"]     = m.group(2)
            out["sb_date"]   = m.group(3)
        # "1 3E23A007 43936.9 USD"
        m2 = re.search(r'^\d+\s+([A-Z0-9]+)\s+([\d.]+)\s+(USD|INR|EUR|GBP)', line.strip())
        if m2 and not out["inv_no"]:
            out["inv_no"]   = m2.group(1)
            out["inv_amt"]  = m2.group(2)
            out["currency"] = m2.group(3)

    return out

# ── 2. PDF PAGES → IMAGES (for Claude vision) ────────────────────────────────

def pdf_pages_to_images(pdf_path: str, page_indices: list, dpi: int = 150) -> list:
    try:
        import fitz
    except ImportError:
        import pymupdf as fitz
    doc = fitz.open(pdf_path)
    mat = fitz.Matrix(dpi/72, dpi/72)
    imgs = []
    for i in page_indices:
        if i < len(doc):
            pix = doc[i].get_pixmap(matrix=mat)
            imgs.append(pix.tobytes("png"))
    doc.close()
    return imgs


def find_part2_end(pdf_path: str) -> int:
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            if i > 0 and "PART - III" in (page.extract_text() or ""):
                return i
        return len(pdf.pages)


# ── 3. CLAUDE VISION (Part II: val_dtls + items) ─────────────────────────────

PART2_PROMPT = """You are reading Part II invoice detail pages of an Indian Customs EDI System shipping bill.

Extract TWO things:

1. C.VAL DTLS row (near top of first page, left side, labeled C.VAL DTLS or C.VAL/DTLS):
   The row has HEADER labels on one line and DATA VALUES on the line below.
   Column order strictly left to right:
   Invoice Value | FOB Value | Freight | Insurance | Discount | Commission | Deduct | P/C | Exchange Rate

   The currency for Invoice Value and FOB Value appears just below their number (e.g. "USD").
   Exchange Rate looks like "1 USD INR 81.4".

   IMPORTANT: Values are numbers only. Do not output header label text as values.

2. ALL items from the D.ITEM DETAILS table across ALL pages shown.
   Columns: Item No, HS Code, Description, Quantity, Rate, Value(F/C).
   Clean up wrapped/split description text into one string per item.
   quantity, rate, value_fc must be numbers.
   Include every item — do not stop early.

Return this exact JSON (no markdown, no explanation):
{
  "val_dtls": {
    "invoice_value": "",
    "invoice_currency": "",
    "fob_value": "",
    "fob_currency": "",
    "freight": "",
    "insurance": "",
    "discount": "",
    "commission": "",
    "deduct": "",
    "pc": "",
    "exchange_rate": ""
  },
  "items": [
    {
      "item_no": 1,
      "hs_code": "73089090",
      "description": "full clean description",
      "quantity": 750,
      "rate": 1.546,
      "value_fc": 1159.5
    }
  ]
}"""


def get_api_key() -> str:
    import os
    key = os.environ.get("ANTHROPIC_API_KEY","")
    if key: return key
    key_file = os.path.join(
        os.path.dirname(sys.executable) if getattr(sys,"frozen",False)
        else os.path.dirname(os.path.abspath(__file__)),
        "key.txt"
    )
    if os.path.exists(key_file):
        with open(key_file) as f: return f.read().strip()
    import tkinter as tk
    from tkinter import simpledialog, messagebox
    root = tk.Tk(); root.withdraw()
    key = simpledialog.askstring(
        "API Key Required",
        "Enter your Anthropic API key.\n\nTip: put it in key.txt next to this app.",
        show="*"
    )
    if not key:
        messagebox.showerror("Error","No API key provided. Exiting."); sys.exit(1)
    if messagebox.askyesno("Save key?","Save to key.txt so you're not asked again?"):
        with open(key_file,"w") as f: f.write(key)
    return key


def call_claude_vision(prompt: str, image_bytes_list: list) -> str:
    client  = anthropic.Anthropic(api_key=get_api_key())
    content = []
    for img in image_bytes_list:
        content.append({"type":"image","source":{
            "type":"base64","media_type":"image/png",
            "data": base64.standard_b64encode(img).decode()
        }})
    content.append({"type":"text","text":prompt})
    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        messages=[{"role":"user","content":content}]
    )
    return resp.content[0].text.strip()


def parse_json(raw: str):
    raw = re.sub(r"^```[a-z]*\n?","",raw.strip()).rstrip("`").strip()
    return json.loads(raw)


# ── 4. MAIN EXTRACTION ────────────────────────────────────────────────────────

def extract_from_pdf(pdf_path: str) -> tuple:
    """Returns (summary, val_dtls, items)."""
    # Part I: fully positional — no AI
    summary = extract_part1(pdf_path)

    # Part II: Claude vision for val_dtls + items
    part2_end  = find_part2_end(pdf_path)
    part2_idxs = list(range(1, min(part2_end, 5)))  # pages 1-4 max
    part2_imgs = pdf_pages_to_images(pdf_path, part2_idxs)

    raw      = call_claude_vision(PART2_PROMPT, part2_imgs)
    part2    = parse_json(raw)
    val_dtls = part2.get("val_dtls", {})
    items    = part2.get("items", [])

    return summary, val_dtls, items

# ── 3. EXCEL OUTPUT ───────────────────────────────────────────────────────────

def write_to_sheet(ws, summary: dict, val_dtls: dict, items: list[dict], start_row: int) -> int:
    """
    Write one PDF's data into ws starting at start_row.
    Returns the next available row after all data is written.

    Layout per PDF (relative to start_row):
      start_row + 0 : Part I + Val Dtls headers  (cols A–AN)
      start_row + 1 : Part I + Val Dtls values    (cols A–AN)
      start_row + 2 : Item table headers          (cols AP–AU)
      start_row + 3+: Item rows
    """
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    def cell_style(cell, value, bold=False, bg=None, font_color="000000",
                   align=None, nfmt=None):
        cell.value = value
        cell.font  = Font(name="Arial", bold=bold, size=10, color=font_color)
        if bg:
            cell.fill = PatternFill("solid", start_color=bg)
        if align:
            cell.alignment = align
        if nfmt:
            cell.number_format = nfmt

    # ── Header row ────────────────────────────────────────────────────────────
    hr = start_row
    ws.row_dimensions[hr].height = 13.9

    header_cols = [
        (1,  "Port Code",        "DEEAF1"),
        (2,  "SB No",            None),
        (3,  "SB Date",          "DEEAF1"),
        (6,  "FOB Value",        "DEEAF1"),
        (7,  "Freight",          None),
        (8,  "Insurance",        "DEEAF1"),
        (9,  "Discount",         None),
        (10, "COM",              "DEEAF1"),
        (11, "Deductions",       None),
        (12, "P/C",              "DEEAF1"),
        (13, "Duty",             None),
        (14, "Cess",             "DEEAF1"),
        (17, "DBK Claim",        "DEEAF1"),
        (18, "IGST AMT",         None),
        (19, "Cess AMT",         "DEEAF1"),
        (20, "IGST Value",       None),
        (21, "RODTEP AMT",       "DEEAF1"),
        (22, "ROSCTL AMT",       None),
        (25, "INV No",           "DEEAF1"),
        (26, "INV AMT",          None),
        (27, "Currency",         "DEEAF1"),
        (30, "Invoice Value",    "DEEAF1"),
        (31, "Invoice Currency", None),
        (32, "FOB Value",        "DEEAF1"),
        (33, "FOB Currency",     None),
        (34, "Freight",          "DEEAF1"),
        (35, "Insurance",        None),
        (36, "Discount",         "DEEAF1"),
        (37, "Commission",       None),
        (38, "Deduct",           "DEEAF1"),
        (39, "P/C",              None),
        (40, "Exchange Rate",    "DEEAF1"),
    ]
    for col_idx, label, bg in header_cols:
        cell_style(ws.cell(row=hr, column=col_idx), label, bold=True, bg=bg)

    # ── Value row ─────────────────────────────────────────────────────────────
    vr = start_row + 1
    value_cols = [
        (1,  summary.get("port_code")),
        (2,  summary.get("sb_no")),
        (3,  summary.get("sb_date")),
        (6,  summary.get("fob_value")),
        (7,  summary.get("freight")),
        (8,  summary.get("insurance")),
        (9,  summary.get("discount")),
        (10, summary.get("com")),
        (11, summary.get("deductions")),
        (12, summary.get("pc")),
        (13, summary.get("duty")),
        (14, summary.get("cess")),
        (17, summary.get("dbk_claim")),
        (18, summary.get("igst_amt")),
        (19, summary.get("cess_amt")),
        (20, summary.get("igst_value")),
        (21, summary.get("rodtep_amt")),
        (22, summary.get("rosctl_amt")),
        (25, summary.get("inv_no")),
        (26, summary.get("inv_amt")),
        (27, summary.get("currency")),
        (30, val_dtls.get("invoice_value")),
        (31, val_dtls.get("invoice_currency")),
        (32, val_dtls.get("fob_value")),
        (33, val_dtls.get("fob_currency")),
        (34, val_dtls.get("freight")),
        (35, val_dtls.get("insurance")),
        (36, val_dtls.get("discount")),
        (37, val_dtls.get("commission")),
        (38, val_dtls.get("deduct")),
        (39, val_dtls.get("pc")),
        (40, val_dtls.get("exchange_rate")),
    ]
    for col_idx, value in value_cols:
        cell_style(ws.cell(row=vr, column=col_idx), value, bg="FFFF00")

    # ── Item headers row ──────────────────────────────────────────────────────
    ihr = start_row + 2
    ws.row_dimensions[ihr].height = 26.25

    cell_style(ws.cell(row=ihr, column=42), "SB No", bold=True)
    for i, hdr in enumerate(["HS Code", "Description", "Quantity", "Rate (USD)", "Value F/C (USD)"]):
        cell_style(ws.cell(row=ihr, column=43 + i), hdr,
                   bold=True, bg="2E75B6", font_color="FFFFFF", align=center)

    # ── Item rows ─────────────────────────────────────────────────────────────
    sb_no      = summary.get("sb_no", "")
    item_nfmts = [None, None, "#,##0", "#,##0.000", "#,##0.00"]
    item_aligns = [center, left, center, center, center]

    for idx, item in enumerate(items):
        r      = start_row + 3 + idx
        fill_bg = "DEEAF1" if idx % 2 == 0 else None

        cell_style(ws.cell(row=r, column=42), sb_no, bg="FFFF00")

        values = [
            item.get("hs_code"),
            item.get("description"),
            item.get("quantity"),
            item.get("rate"),
            item.get("value_fc"),
        ]
        for i, (val, aln, nfmt) in enumerate(zip(values, item_aligns, item_nfmts)):
            cell_style(ws.cell(row=r, column=43 + i), val,
                       bg=fill_bg, align=aln, nfmt=nfmt)

    # Return next available row (leave 1 blank gap between PDFs)
    return start_row + 3 + len(items) + 1


def build_workbook(all_results: list[dict], skipped: list[tuple],
                   failures: list[tuple], output_path: str):
    """
    Create a single workbook with:
      Sheet 1 "Customs Extract" — all valid PDFs appended vertically
      Sheet 2 "Errors"          — skipped duplicates + extraction failures
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Customs Extract"

    # Column widths (shared by all PDF blocks)
    col_widths = {
        "A": 8.9375,  "B": 7.25,   "C": 9.25,
        "F": 9.625,   "G": 6.5625, "H": 8.8125, "I": 8.0625,
        "J": 4.5625,  "K": 10.0625,"L": 3.5625, "M": 4.5,    "N": 4.75,
        "Q": 9.75,    "R": 8.875,  "S": 9.8125,
        "T": 9.8125,  "U": 12.0625,"V": 12.0,
        "Y": 8.5625,  "Z": 7.6875, "AA": 8.25,
        "AD": 11.6875,"AE": 14.75, "AF": 9.375, "AG": 12.4375,
        "AH": 6.5625, "AI": 8.8125,"AJ": 8.0625,"AK": 11.25,
        "AL": 6.5,    "AM": 3.5625,"AN": 13.0,
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    current_row = 1
    for result in all_results:
        current_row = write_to_sheet(
            ws,
            result["summary"],
            result["val_dtls"],
            result["items"],
            start_row=current_row
        )

    # ── Errors sheet (always created) ────────────────────────────────────────
    we = wb.create_sheet("Errors")
    we.column_dimensions["A"].width = 40
    we.column_dimensions["B"].width = 18
    we.column_dimensions["C"].width = 50

    hdr_fill  = PatternFill("solid", start_color="C00000")
    hdr_font  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    row_fill  = PatternFill("solid", start_color="FFE0E0")
    nrm_font  = Font(name="Arial", size=10)
    center    = Alignment(horizontal="center", vertical="center")
    left      = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    for col, label in enumerate(["File", "Type", "Reason"], 1):
        cell = we.cell(row=1, column=col, value=label)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.alignment = center
    we.row_dimensions[1].height = 18

    row = 2
    for filename, reason in skipped:
        vals = [filename, "Duplicate", reason]
        for col, val in enumerate(vals, 1):
            cell = we.cell(row=row, column=col, value=val)
            cell.font      = nrm_font
            cell.fill      = row_fill
            cell.alignment = left
        we.row_dimensions[row].height = 16
        row += 1

    for filename, reason in failures:
        vals = [filename, "Error", reason]
        for col, val in enumerate(vals, 1):
            cell = we.cell(row=row, column=col, value=val)
            cell.font      = nrm_font
            cell.alignment = left
        we.row_dimensions[row].height = 16
        row += 1

    wb.save(output_path)
    print(f"✅  Saved: {output_path}")


# ── 4. MAIN ───────────────────────────────────────────────────────────────────

def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    # ── File picker ───────────────────────────────────────────────────────────
    root = tk.Tk()
    root.withdraw()

    pdf_paths = filedialog.askopenfilenames(
        title="Select Customs PDF files to process",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )
    if not pdf_paths:
        messagebox.showinfo("Cancelled", "No files selected. Exiting.")
        sys.exit(0)

    output_dir = filedialog.askdirectory(title="Select folder to save Excel output")
    if not output_dir:
        messagebox.showinfo("Cancelled", "No output folder selected. Exiting.")
        sys.exit(0)

    # ── Progress window ───────────────────────────────────────────────────────
    progress_win = tk.Tk()
    progress_win.title("Customs PDF Extractor")
    progress_win.geometry("500x300")
    progress_win.resizable(False, False)

    tk.Label(progress_win, text="Processing PDFs...", font=("Arial", 12, "bold")).pack(pady=10)
    progress_bar = ttk.Progressbar(progress_win, length=460, mode="determinate",
                                   maximum=len(pdf_paths))
    progress_bar.pack(pady=5, padx=20)
    status_label = tk.Label(progress_win, text="", font=("Arial", 9), wraplength=460)
    status_label.pack(pady=5)
    log_box = tk.Text(progress_win, height=10, font=("Courier", 8), state="disabled")
    log_box.pack(pady=5, padx=20, fill="both", expand=True)

    def log(msg):
        log_box.config(state="normal")
        log_box.insert("end", msg + "\n")
        log_box.see("end")
        log_box.config(state="disabled")
        progress_win.update()

    # ── Process each PDF, collect results ────────────────────────────────────
    import os
    from datetime import datetime

    all_results = []
    failures    = []
    skipped     = []  # list of (filename, reason)

    seen_sb_nos  = set()
    seen_inv_nos = set()

    for i, pdf_path in enumerate(pdf_paths):
        filename = os.path.basename(pdf_path)
        status_label.config(text=f"Processing: {filename}")
        log(f"\n[{i+1}/{len(pdf_paths)}] {filename}")
        progress_win.update()

        try:
            log("  → Rendering PDF + extracting Part I with Claude vision...")
            summary, val_dtls, items = extract_from_pdf(pdf_path)

            # ── Duplicate check ───────────────────────────────────────────
            sb_no  = str(summary.get("sb_no")  or "").strip()
            inv_no = str(summary.get("inv_no") or "").strip()

            dup_reason = None
            if sb_no and sb_no in seen_sb_nos:
                dup_reason = f"duplicate SB No ({sb_no})"
            elif inv_no and inv_no in seen_inv_nos:
                dup_reason = f"duplicate Invoice No ({inv_no})"

            if dup_reason:
                log(f"  ⚠️  SKIPPED — {dup_reason}")
                skipped.append((filename, dup_reason))
                progress_bar["value"] = i + 1
                progress_win.update()
                continue

            # Not a duplicate — register and proceed
            if sb_no:  seen_sb_nos.add(sb_no)
            if inv_no: seen_inv_nos.add(inv_no)

            all_results.append({"summary": summary, "val_dtls": val_dtls, "items": items})
            log(f"  ✅ Extracted {len(items)} items")

        except Exception as e:
            log(f"  ❌ ERROR: {e}")
            failures.append((filename, str(e)))

        progress_bar["value"] = i + 1
        progress_win.update()

    # ── Write single combined Excel ───────────────────────────────────────────
    if all_results or skipped or failures:
        status_label.config(text="Building Excel...")
        progress_win.update()

        timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_name = f"customs_extracted_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_name)

        log(f"\n→ Writing {len(all_results)} PDF(s) to {output_name} ...")
        try:
            build_workbook(all_results, skipped, failures, output_path)
            log(f"✅ Saved: {output_path}")
        except Exception as e:
            log(f"❌ Failed to save Excel: {e}")

    # ── Done summary ──────────────────────────────────────────────────────────
    status_label.config(text="Done!")

    summary_msg = (
        f"✅ {len(all_results)} PDF(s) extracted\n"
        f"⚠️  {len(skipped)} skipped (duplicates)\n"
        f"❌ {len(failures)} failed\n\n"
        f"Saved to:\n{output_dir}"
    )
    if skipped or failures:
        summary_msg += "\n\nSee the 'Errors' sheet in the Excel file for details."

    log("\n" + "=" * 50)
    log(summary_msg)
    messagebox.showinfo("Complete", summary_msg)
    progress_win.protocol("WM_DELETE_WINDOW", lambda: (progress_win.destroy(), sys.exit(0)))
    progress_win.mainloop()
    sys.exit(0)


if __name__ == "__main__":
    main()