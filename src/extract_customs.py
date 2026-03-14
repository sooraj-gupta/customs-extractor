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

def build_excel(summary: dict, val_dtls: dict, items: list[dict], output_path: str):
    wb = Workbook()

    # Styles
    hdr_font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    hdr_fill      = PatternFill("solid", start_color="1F4E79")
    sub_fill      = PatternFill("solid", start_color="2E75B6")
    sub_font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    alt_fill      = PatternFill("solid", start_color="DEEAF1")
    yellow_fill   = PatternFill("solid", start_color="FFFF00")
    normal_font   = Font(name="Arial", size=10)
    bold_font     = Font(name="Arial", bold=True, size=10)
    center        = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left          = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    right_align   = Alignment(horizontal="right",  vertical="center")

    thin = Side(style="thin", color="B8CCE4")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def apply(cell, font=None, fill=None, alignment=None, border=None, number_format=None):
        if font:          cell.font          = font
        if fill:          cell.fill          = fill
        if alignment:     cell.alignment     = alignment
        if border:        cell.border        = border
        if number_format: cell.number_format = number_format

    def write_kv_sheet(ws, title, sections):
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 34
        ws.column_dimensions["C"].width = 24
        ws.column_dimensions["D"].width = 6
        ws.merge_cells("A1:D1")
        ws["A1"] = title
        apply(ws["A1"], font=Font(name="Arial", bold=True, size=12, color="FFFFFF"),
              fill=hdr_fill, alignment=center)
        ws.row_dimensions[1].height = 28
        row = 2
        for section_title, fields in sections:
            ws.merge_cells(f"A{row}:D{row}")
            ws[f"A{row}"] = section_title
            apply(ws[f"A{row}"], font=sub_font, fill=sub_fill, alignment=left)
            ws.row_dimensions[row].height = 18
            row += 1
            for i, (label, value) in enumerate(fields):
                fill = alt_fill if i % 2 == 0 else None
                ws[f"B{row}"] = label
                ws[f"C{row}"] = value
                apply(ws[f"B{row}"], font=bold_font, fill=fill,  alignment=left, border=border)
                apply(ws[f"C{row}"], font=normal_font, fill=yellow_fill, alignment=left, border=border)
                ws.row_dimensions[row].height = 16
                row += 1
            row += 1

    # ── Sheet 1: Part I Summary ───────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Part I – Summary"
    write_kv_sheet(ws1, "INDIAN CUSTOMS EDI SYSTEM – SHIPPING BILL SUMMARY (PART I)", [
        ("IDENTIFICATION", [
            ("Port Code",  summary.get("port_code")),
            ("SB No",      summary.get("sb_no")),
            ("SB Date",    summary.get("sb_date")),
        ]),
        ("VALUATION SUMMARY (C. VALU SUMMA)", [
            ("FOB Value",   summary.get("fob_value")),
            ("Freight",     summary.get("freight")),
            ("Insurance",   summary.get("insurance")),
            ("Discount",    summary.get("discount")),
            ("COM",         summary.get("com")),
            ("Deductions",  summary.get("deductions")),
            ("P/C",         summary.get("pc")),
            ("Duty",        summary.get("duty")),
            ("Cess",        summary.get("cess")),
        ]),
        ("EXPORT PROMOTION (D. EX.PR.)", [
            ("DBK Claim",   summary.get("dbk_claim")),
            ("IGST AMT",    summary.get("igst_amt")),
            ("Cess AMT",    summary.get("cess_amt")),
            ("IGST Value",  summary.get("igst_value")),
            ("RODTEP AMT",  summary.get("rodtep_amt")),
            ("ROSCTL AMT",  summary.get("rosctl_amt")),
        ]),
        ("INVOICE (F. INVOICE SUMMARY)", [
            ("INV No",      summary.get("inv_no")),
            ("INV AMT",     summary.get("inv_amt")),
            ("Currency",    summary.get("currency")),
        ]),
    ])

    # ── Sheet 2: Part II – C.VAL DTLS ────────────────────────────────────────
    ws2 = wb.create_sheet("Part II – Val Dtls")
    write_kv_sheet(ws2, "INDIAN CUSTOMS EDI SYSTEM – INVOICE VALUATION DETAILS (PART II)", [
        ("C.VAL DTLS – INVOICE VALUATION", [
            ("Invoice Value",   val_dtls.get("invoice_value")),
            ("Invoice Currency",val_dtls.get("invoice_currency")),
            ("FOB Value",       val_dtls.get("fob_value")),
            ("FOB Currency",    val_dtls.get("fob_currency")),
            ("Freight",         val_dtls.get("freight")),
            ("Insurance",       val_dtls.get("insurance")),
            ("Discount",        val_dtls.get("discount")),
            ("Commission",      val_dtls.get("commission")),
            ("Deduct",          val_dtls.get("deduct")),
            ("P/C",             val_dtls.get("pc")),
            ("Exchange Rate",   val_dtls.get("exchange_rate")),
        ]),
    ])

    # ── Sheet 3: Part II – Items ──────────────────────────────────────────────
    ws3 = wb.create_sheet("Part II – Items")

    # Columns: Item No, HS Code, Description, Quantity, Rate, Value F/C
    col_widths = [10, 14, 55, 12, 14, 16]
    for i, w in enumerate(col_widths):
        ws3.column_dimensions[get_column_letter(i+1)].width = w

    ws3.merge_cells("A1:F1")
    ws3["A1"] = "INDIAN CUSTOMS EDI SYSTEM – INVOICE ITEM DETAILS (PART II)"
    apply(ws3["A1"], font=Font(name="Arial", bold=True, size=12, color="FFFFFF"),
          fill=hdr_fill, alignment=center)
    ws3.row_dimensions[1].height = 28

    headers = ["Item No", "HS Code", "Description", "Quantity", "Rate (USD)", "Value F/C (USD)"]
    for col, h in enumerate(headers, 1):
        cell = ws3.cell(row=2, column=col, value=h)
        apply(cell, font=sub_font, fill=sub_fill, alignment=center, border=border)
    ws3.row_dimensions[2].height = 20

    for idx, item in enumerate(items):
        r = idx + 3
        fill = alt_fill if idx % 2 == 0 else None
        values   = [item.get("item_no"), item.get("hs_code"), item.get("description"),
                    item.get("quantity"), item.get("rate"), item.get("value_fc")]
        aligns   = [center, center, left, right_align, right_align, right_align]
        num_fmts = [None, None, None, "#,##0", "#,##0.000", "#,##0.00"]
        for col, (val, aln, nfmt) in enumerate(zip(values, aligns, num_fmts), 1):
            cell = ws3.cell(row=r, column=col, value=val)
            apply(cell, font=normal_font, fill=fill, alignment=aln,
                  border=border, number_format=nfmt)
        ws3.row_dimensions[r].height = 30

    # Totals row
    total_row = len(items) + 3
    ws3.merge_cells(f"A{total_row}:C{total_row}")
    ws3[f"A{total_row}"] = "TOTAL"
    apply(ws3[f"A{total_row}"], font=bold_font, fill=sub_fill,
          alignment=Alignment(horizontal="right", vertical="center"))

    qty_cell   = ws3.cell(row=total_row, column=4, value=f"=SUM(D3:D{total_row-1})")
    value_cell = ws3.cell(row=total_row, column=6, value=f"=SUM(F3:F{total_row-1})")
    for cell, nfmt in [(qty_cell, "#,##0"), (value_cell, "#,##0.00")]:
        apply(cell, font=bold_font, fill=yellow_fill,
              alignment=right_align, border=border, number_format=nfmt)
    apply(ws3.cell(row=total_row, column=5), fill=sub_fill, border=border)

    ws3.freeze_panes = "A3"

    wb.save(output_path)
    print(f"✅  Saved: {output_path}")


# ── 4. MAIN ───────────────────────────────────────────────────────────────────

def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    # ── File picker window ────────────────────────────────────────────────────
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    pdf_paths = filedialog.askopenfilenames(
        title="Select Customs PDF files to process",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )

    if not pdf_paths:
        messagebox.showinfo("Cancelled", "No files selected. Exiting.")
        return

    # Ask where to save outputs
    output_dir = filedialog.askdirectory(title="Select folder to save Excel outputs")
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

    # ── Process each PDF ──────────────────────────────────────────────────────
    import os
    successes, failures = [], []

    for i, pdf_path in enumerate(pdf_paths):
        filename = os.path.basename(pdf_path)
        status_label.config(text=f"Processing: {filename}")
        log(f"\n[{i+1}/{len(pdf_paths)}] {filename}")
        progress_win.update()

        try:
            output_filename = os.path.splitext(filename)[0] + "_extracted.xlsx"
            output_path = os.path.join(output_dir, output_filename)

            log("  → Reading PDF...")
            part1_text, part2_text = get_part_texts(pdf_path)

            log("  → Extracting Part I (Claude)...")
            summary = extract_summary(part1_text)

            log("  → Extracting Part II (Claude + positional)...")
            part2_data = extract_part2(pdf_path, part2_text)
            val_dtls = part2_data.get("val_dtls", {})
            items    = part2_data.get("items", [])

            log(f"  → Building Excel ({len(items)} items)...")
            build_excel(summary, val_dtls, items, output_path)

            log(f"  ✅ Saved: {output_filename}")
            successes.append(filename)

        except Exception as e:
            log(f"  ❌ ERROR: {e}")
            failures.append((filename, str(e)))

        progress_bar["value"] = i + 1
        progress_win.update()

    # ── Done summary ──────────────────────────────────────────────────────────
    status_label.config(text="Done!")
    summary_msg = f"✅ {len(successes)} succeeded,  ❌ {len(failures)} failed.\n\nOutputs saved to:\n{output_dir}"
    if failures:
        summary_msg += "\n\nFailed files:\n" + "\n".join(f"  • {f[0]}: {f[1]}" for f in failures)

    log("\n" + "="*50)
    log(summary_msg)

    messagebox.showinfo("Complete", summary_msg)
    progress_win.mainloop()


if __name__ == "__main__":
    main()
