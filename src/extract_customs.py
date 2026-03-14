"""
extract_customs.py
------------------
Extracts highlighted fields from Indian Customs EDI System PDFs.

- Part I  (page 1):  summary fields — Port Code, SB No, SB Date, FOB Value,
                     Freight, Insurance, Discount, COM, Deductions, P/C, Duty,
                     Cess, DBK Claim, IGST AMT, Cess AMT, IGST Value,
                     RODTEP AMT, ROSCTL AMT, INV NO, INV AMT, Currency
- Part II (pages 2–N before Part III): item table — Item No, HS Code,
                     Description, Quantity, UQC, Rate, Value (F/C)

Usage:
    python extract_customs.py <input.pdf> [output.xlsx]

Requirements:
    pip install pdfplumber anthropic openpyxl
"""

import sys
import re
import json
import pdfplumber
import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── 1. PDF TEXT EXTRACTION ────────────────────────────────────────────────────

def get_part_texts(pdf_path: str) -> tuple[str, str]:
    """
    Returns (part1_text, part2_text).
    Part I  = page 1 only.
    Part II = pages 2..N where N is the last page before 'PART - III' appears.
    """
    with pdfplumber.open(pdf_path) as pdf:
        part1_text = pdf.pages[0].extract_text() or ""

        part2_pages = []
        for page in pdf.pages[1:]:
            text = page.extract_text() or ""
            if "PART - III" in text:
                # Stop before Part III (page may contain both Part II rows
                # at top and Part III below — grab only the Part II portion)
                lines = text.split("\n")
                part2_lines = []
                for line in lines:
                    if "PART - III" in line:
                        break
                    part2_lines.append(line)
                if part2_lines:
                    part2_pages.append("\n".join(part2_lines))
                break
            part2_pages.append(text)

        part2_text = "\n".join(part2_pages)

    return part1_text, part2_text


# ── 2. AI EXTRACTION ──────────────────────────────────────────────────────────

PART1_PROMPT = """You are a data extraction assistant for Indian Customs EDI shipping bills.

Extract ONLY the following fields from the Part I shipping bill summary text below.

IMPORTANT FIELD MAPPING RULES — the C.VALU SUMMA section has this exact column order:
  "1.FOB VALUE | 2.FREIGHT | 3.INSURANC | 4.DISCOU | 5.COM"
  "6.DEDUCTIONS | 7.P/C | 8.DUTY | 9.CESS"

The values appear on the line(s) BELOW the column headers in the SAME left-to-right order.
For example if the text reads:
  "1.FOB VALUE  2.FREIGHT  3.INSURANC  4.DISCOU  5.COM"
  "3575464.88   0          999         0          0"
Then: fob_value=3575464.88, freight=0, insurance=999, discount=0, com=0

Map each number strictly to its column position — do NOT guess from the value itself.

Return a JSON object with exactly these keys (use null if not found):

{{
  "port_code": "",
  "sb_no": "",
  "sb_date": "",
  "fob_value": "",
  "freight": "",
  "insurance": "",
  "discount": "",
  "com": "",
  "deductions": "",
  "pc": "",
  "duty": "",
  "cess": "",
  "dbk_claim": "",
  "igst_amt": "",
  "cess_amt": "",
  "igst_value": "",
  "rodtep_amt": "",
  "rosctl_amt": "",
  "inv_no": "",
  "inv_amt": "",
  "currency": ""
}}

Return ONLY the JSON object, no explanation, no markdown.

--- PART I TEXT ---
{text}
"""

PART2_ITEMS_PROMPT = """You are a data extraction assistant for Indian Customs EDI shipping bills.

From the Part II invoice details text below, extract ONLY the item table (D. ITEM DETAILS).
Each item has columns: Item No, HS Code, Description, Quantity, Rate, Value(F/C).
Do NOT extract UQC.

Return a JSON array where each element is:
{{
  "item_no": 1,
  "hs_code": "73089090",
  "description": "BAR B ARM 2\" X 1-5/8\" (BLACK)",
  "quantity": 750,
  "rate": 1.546,
  "value_fc": 1159.5
}}

Rules:
- Combine split description lines into one clean string (descriptions often wrap to next line)
- Strip any OCR noise characters (lone letters like "O", "X", "E" on their own line are not part of description)
- quantity, rate, value_fc should be numbers (not strings)
- item_no should be an integer
- Include ALL items found across all pages
- Return ONLY the JSON array, no explanation, no markdown.

--- PART II TEXT ---
{text}
"""

# Column x-boundaries for C.VAL DTLS positional extraction
# Based on the fixed form layout of Indian Customs EDI PDFs
_VAL_DTLS_COLS = {
    "invoice_value": (75,  145),
    "fob_value":     (145, 220),
    "freight":       (220, 266),
    "insurance":     (266, 313),
    "discount":      (313, 360),
    "commission":    (360, 406),
    "deduct":        (406, 465),
    "pc":            (465, 497),
    "exchange_rate": (497, 580),
}

def extract_val_dtls_positional(pdf_path: str) -> dict:
    """
    Extracts C.VAL DTLS values using word x-coordinates rather than text order,
    avoiding the column-swap problem caused by pdfplumber's text linearisation.
    """
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[1:]:
            words = page.extract_words()

            # Locate the header anchor word "1.INVOICE"
            header_y = None
            for w in words:
                if w["text"] == "1.INVOICE":
                    header_y = w["top"]
                    break
            if header_y is None:
                continue

            # Value row sits ~13–14 px below the header row
            value_y   = header_y + 13.6
            tol       = 8

            val_words = [w for w in words if abs(w["top"] - value_y) < tol]
            val_words.sort(key=lambda w: w["x0"])

            result = {k: "0" for k in _VAL_DTLS_COLS}

            for w in val_words:
                cx = (w["x0"] + w["x1"]) / 2
                txt = w["text"]
                for col, (x_min, x_max) in _VAL_DTLS_COLS.items():
                    if x_min <= cx < x_max:
                        if col == "exchange_rate":
                            result[col] = (result[col] + " " + txt).strip() \
                                          if result[col] != "0" else txt
                        else:
                            result[col] = txt
                        break

            # Currency labels appear ~22 px below header
            curr_y    = header_y + 22
            curr_words = [w for w in words
                          if abs(w["top"] - curr_y) < 8
                          and w["text"] in ("USD", "INR", "EUR", "GBP")]
            result["invoice_currency"] = next(
                (w["text"] for w in curr_words if 75  <= (w["x0"]+w["x1"])/2 < 145), "")
            result["fob_currency"]     = next(
                (w["text"] for w in curr_words if 145 <= (w["x0"]+w["x1"])/2 < 220), "")

            return result

    return {}


def call_claude(prompt: str) -> str:
    import os
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        # Read from a key.txt file next to the exe, set once by user
        key_file = os.path.join(os.path.dirname(sys.executable)
                                if getattr(sys, "frozen", False)
                                else os.path.dirname(os.path.abspath(__file__)),
                                "key.txt")
        if os.path.exists(key_file):
            with open(key_file) as f:
                api_key = f.read().strip()
        else:
            import tkinter as tk
            from tkinter import simpledialog, messagebox
            root = tk.Tk(); root.withdraw()
            api_key = simpledialog.askstring(
                "API Key Required",
                "Enter your Anthropic API key.\n\n"
                "Tip: create a file called key.txt next to this app\n"
                "and paste your key there to avoid this prompt.",
                show="*"
            )
            if not api_key:
                messagebox.showerror("Error", "No API key provided. Exiting.")
                sys.exit(1)
            # Offer to save it
            if messagebox.askyesno("Save key?", "Save key to key.txt so you're not asked again?"):
                with open(key_file, "w") as f:
                    f.write(api_key)

    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )
    return response.content[0].text.strip()


def extract_summary(part1_text: str) -> dict:
    # Pre-process: find and annotate the valuation rows so AI maps them correctly.
    # pdfplumber loses column alignment; we know the fixed column order from the form:
    #   Row 1 headers: 1.FOB VALUE | 2.FREIGHT | 3.INSURANC | 4.DISCOU | 5.COM
    #   Row 2 values:  fob_value     freight      insurance    discount   com
    #   Row 3 headers: 6.DEDUCTIONS | 7.P/C | 8.DUTY | 9.CESS
    #   Row 4 values:  deductions     pc       duty     cess
    #
    # pdfplumber extracts FOB VALUE on its own line (e.g. "LM 3575464.88")
    # and the remaining values on one line (e.g. "0 999 0 0")
    # We annotate these lines before sending to AI.

    import re
    lines = part1_text.split('\n')
    annotated_lines = []

    i = 0
    while i < len(lines):
        line = lines[i]

        # Detect the header line for the first valuation row
        if '1.FOB VALUE' in line and '2.FREIGHT' in line:
            annotated_lines.append("=== VALUATION ROW 1 HEADERS: [1.FOB VALUE] [2.FREIGHT] [3.INSURANCE] [4.DISCOUNT] [5.COM] ===")
            annotated_lines.append(line)
            # FOB value may be on a prior line by itself (pdfplumber quirk)
            # Search backwards for it
            fob_val = None
            for back in range(max(0, i-5), i):
                m = re.search(r'(\d[\d,]+\.\d+)', lines[back])
                if m and float(m.group(1).replace(',', '')) > 1000:
                    fob_val = m.group(1)
            if fob_val:
                annotated_lines.append(f"FOB_VALUE_FOUND_ON_PRIOR_LINE: {fob_val}")

            # Next line(s) with values: "0 999 0 0" = freight, insurance, discount, com
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                nums = re.findall(r'[\d,]+(?:\.\d+)?', next_line)
                if nums:
                    labels = ['FREIGHT', 'INSURANCE', 'DISCOUNT', 'COM']
                    annotated_lines.append("VALUES IN COLUMN ORDER: " +
                        ", ".join(f"{labels[j]}={nums[j]}" for j in range(min(len(labels), len(nums)))))
            i += 1

        # Detect deductions row
        elif '6.DEDUCTIONS' in line and '8.DUTY' in line:
            annotated_lines.append("=== VALUATION ROW 2 HEADERS: [6.DEDUCTIONS] [7.P/C] [8.DUTY] [9.CESS] ===")
            annotated_lines.append(line)
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                nums = re.findall(r'[\d,]+(?:\.\d+)?', next_line)
                if nums:
                    labels = ['DEDUCTIONS', 'PC', 'DUTY', 'CESS']
                    annotated_lines.append("VALUES IN COLUMN ORDER: " +
                        ", ".join(f"{labels[j]}={nums[j]}" for j in range(min(len(labels), len(nums)))))
            i += 1
        else:
            annotated_lines.append(line)
        i += 1

    annotated_text = '\n'.join(annotated_lines)
    raw = call_claude(PART1_PROMPT.format(text=annotated_text))
    raw = re.sub(r"^```[a-z]*\n?", "", raw).rstrip("```").strip()
    return json.loads(raw)


def extract_part2(pdf_path: str, part2_text: str) -> dict:
    val_dtls = extract_val_dtls_positional(pdf_path)

    raw = call_claude(PART2_ITEMS_PROMPT.format(text=part2_text))
    raw = re.sub(r"^```[a-z]*\n?", "", raw).rstrip("```").strip()
    items = json.loads(raw)

    return {"val_dtls": val_dtls, "items": items}


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


def build_workbook(all_results: list[dict], output_path: str):
    """
    Create a single workbook with one sheet containing all PDFs appended vertically.
    all_results = list of {"summary": ..., "val_dtls": ..., "items": [...]}
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Customs Extract"

    # Column widths (set once, shared by all PDF blocks)
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
        return

    output_dir = filedialog.askdirectory(title="Select folder to save Excel output")
    if not output_dir:
        messagebox.showinfo("Cancelled", "No output folder selected. Exiting.")
        return

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
    all_results, failures = [], []

    for i, pdf_path in enumerate(pdf_paths):
        filename = os.path.basename(pdf_path)
        status_label.config(text=f"Processing: {filename}")
        log(f"\n[{i+1}/{len(pdf_paths)}] {filename}")
        progress_win.update()

        try:
            log("  → Reading PDF...")
            part1_text, part2_text = get_part_texts(pdf_path)

            log("  → Extracting Part I (Claude)...")
            summary = extract_summary(part1_text)

            log("  → Extracting Part II (Claude + positional)...")
            part2_data = extract_part2(pdf_path, part2_text)
            val_dtls = part2_data.get("val_dtls", {})
            items    = part2_data.get("items", [])

            all_results.append({"summary": summary, "val_dtls": val_dtls, "items": items})
            log(f"  ✅ Extracted {len(items)} items")

        except Exception as e:
            log(f"  ❌ ERROR: {e}")
            failures.append((filename, str(e)))

        progress_bar["value"] = i + 1
        progress_win.update()

    # ── Write single combined Excel ───────────────────────────────────────────
    if all_results:
        status_label.config(text="Building Excel...")
        progress_win.update()
        output_path = os.path.join(output_dir, "customs_extracted.xlsx")
        log(f"\n→ Writing {len(all_results)} PDF(s) to customs_extracted.xlsx ...")
        try:
            build_workbook(all_results, output_path)
            log(f"✅ Saved: {output_path}")
        except Exception as e:
            log(f"❌ Failed to save Excel: {e}")
            failures.append(("customs_extracted.xlsx", str(e)))

    # ── Done ──────────────────────────────────────────────────────────────────
    status_label.config(text="Done!")
    summary_msg = (f"✅ {len(all_results)} PDF(s) extracted into customs_extracted.xlsx\n"
                   f"❌ {len(failures)} failed\n\nSaved to:\n{output_dir}")
    if failures:
        summary_msg += "\n\nFailed:\n" + "\n".join(f"  • {f[0]}: {f[1]}" for f in failures)

    log("\n" + "=" * 50)
    log(summary_msg)
    messagebox.showinfo("Complete", summary_msg)
    progress_win.mainloop()


if __name__ == "__main__":
    main()