"""
extract_customs.py
------------------
Sends PDF pages directly to Claude as a document — no image conversion.

Requirements:
    pip install anthropic openpyxl
"""

import sys, re, json, base64, os
import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


# ── 1. API ────────────────────────────────────────────────────────────────────

def get_api_key() -> str:
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if key:
        return key
    key_file = os.path.join(
        os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
        else os.path.dirname(os.path.abspath(__file__)),
        "key.txt"
    )
    if os.path.exists(key_file):
        with open(key_file) as f:
            return f.read().strip()
    import tkinter as tk
    from tkinter import simpledialog, messagebox
    root = tk.Tk(); root.withdraw()
    key = simpledialog.askstring(
        "API Key Required",
        "Enter your Anthropic API key.\n\nTip: put it in key.txt next to this app.",
        show="*"
    )
    if not key:
        messagebox.showerror("Error", "No API key provided. Exiting.")
        sys.exit(1)
    if messagebox.askyesno("Save key?", "Save to key.txt so you're not asked again?"):
        with open(key_file, "w") as f:
            f.write(key)
    return key


def call_claude(prompt: str, pdf_b64: str, max_tokens: int = 2048) -> str:
    client = anthropic.Anthropic(api_key=get_api_key())
    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=max_tokens,
        system=(
            "You are a data extraction assistant. "
            "Always respond with valid JSON only — no markdown, no explanation. "
            "In string values never use raw double-quote characters; "
            "replace inch marks with 'in' and & with 'and'."
        ),
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": pdf_b64
                    }
                },
                {"type": "text", "text": prompt}
            ]
        }]
    )
    return resp.content[0].text.strip()


def page_to_b64(pdf_path: str, page_idx: int) -> str:
    """Extract a single page from the PDF and return as base64."""
    import pymupdf
    src = pymupdf.open(pdf_path)
    out = pymupdf.open()
    out.insert_pdf(src, from_page=page_idx, to_page=page_idx)
    data = out.tobytes()
    src.close(); out.close()
    return base64.standard_b64encode(data).decode()


def num_pages(pdf_path: str) -> int:
    import pymupdf
    doc = pymupdf.open(pdf_path)
    n = len(doc)
    doc.close()
    return n


def parse_json(raw: str):
    raw = re.sub(r"^```[a-z]*\n?", "", raw.strip()).rstrip("`").strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        for s, e in [('{', '}'), ('[', ']')]:
            i, j = raw.find(s), raw.rfind(e)
            if i != -1 and j > i:
                try:
                    return json.loads(raw[i:j+1])
                except json.JSONDecodeError:
                    pass
        raise


# ── 2. PROMPTS ────────────────────────────────────────────────────────────────

PART1_PROMPT = """This is page 1 of an Indian Customs EDI System shipping bill.

Extract these fields and return a single JSON object.

Sections on this page:

Top-right header box:
  Port Code | SB No | SB Date

C.VALU SUMMA section (left side, mid-page):
  Row 1 headers: 1.FOB VALUE | 2.FREIGHT | 3.INSURANC | 4.DISCOU | 5.COM
  Row 1 values:  fob_value     freight      insurance    discount   com
  Row 2 headers: 6.DEDUCTIONS | 7.P/C | 8.DUTY | 9.CESS
  Row 2 values:  deductions     pc      duty     cess

D.EX.PR section (right side, same area):
  Row 1 headers: 1.DBK CLAIM | 2.IGST AMT | 3.CESS AMT
  Row 1 values:  dbk_claim     igst_amt     cess_amt
  Row 2 headers: 4.IGST VALUE | 5.RODTEP AMT | 6.ROSCTL AMT
  Row 2 values:  igst_value     rodtep_amt     rosctl_amt

F.INVOICE SUMMARY section (right side):
  inv_no, inv_amt, currency

RULES: Values are numbers only. If a value cell contains a header label use "0". Use "0" for empty numeric fields.

Return ONLY this JSON:
{
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
}"""

PART2_PROMPT = """This is page 2 of an Indian Customs EDI System shipping bill (PART II - INVOICE DETAILS).

Extract TWO things:

1. C.VAL DTLS row (near top, labeled C.VAL DTLS):
   Columns left to right: Invoice Value | FOB Value | Freight | Insurance | Discount | Commission | Deduct | P/C | Exchange Rate
   Currency labels (e.g. USD) appear below Invoice Value and FOB Value.
   Exchange Rate looks like "1 USD INR 81.4". Values must be numbers only.

2. Items from the D.ITEM DETAILS table on this page:
   Columns: Item No, HS Code, Description, Quantity, Rate, Value(F/C)
   Combine wrapped description lines. quantity/rate/value_fc must be numbers.

Also tell me if the item table is cut off and continues on the next page.

Return ONLY this JSON:
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
    {"item_no": 1, "hs_code": "73089090", "description": "clean description", "quantity": 750, "rate": 1.546, "value_fc": 1159.5}
  ],
  "has_more_items": false
}"""

PART_EXTRA_PROMPT = """This is page {page_num} of an Indian Customs EDI System shipping bill, continuing the D.ITEM DETAILS table from the previous page.

Extract ALL items shown on this page from the D.ITEM DETAILS table.
Columns: Item No, HS Code, Description, Quantity, Rate, Value(F/C)
Combine wrapped description lines. quantity/rate/value_fc must be numbers.

Also tell me if the table continues on yet another page (i.e. this page ends mid-table without a "PART - III" header appearing).

Return ONLY this JSON:
{
  "items": [
    {"item_no": 1, "hs_code": "73089090", "description": "clean description", "quantity": 750, "rate": 1.546, "value_fc": 1159.5}
  ],
  "has_more_items": false
}"""


# ── 3. EXTRACTION ─────────────────────────────────────────────────────────────

def extract_from_pdf(pdf_path: str, log=None) -> tuple:
    """Returns (summary, val_dtls, items). Processes pages one at a time."""
    def _log(msg):
        if log: log(msg)

    total = num_pages(pdf_path)

    # Page 1 — summary fields
    _log("  → Page 1: extracting summary...")
    summary = {}
    for _ in range(2):
        try:
            summary = parse_json(call_claude(PART1_PROMPT, page_to_b64(pdf_path, 0), max_tokens=1024))
            break
        except (ValueError, json.JSONDecodeError):
            pass

    # Page 2 — val_dtls + first batch of items
    val_dtls, items = {}, []
    if total < 2:
        return summary, val_dtls, items

    _log("  → Page 2: extracting val_dtls + items...")
    has_more = False
    for _ in range(2):
        try:
            result   = parse_json(call_claude(PART2_PROMPT, page_to_b64(pdf_path, 1), max_tokens=4096))
            val_dtls = result.get("val_dtls", {})
            items    = result.get("items", [])
            has_more = result.get("has_more_items", False)
            break
        except (ValueError, json.JSONDecodeError):
            pass

    # Pages 3+ — fetch only if needed and only up to page 4
    for page_idx in range(2, min(total, 4)):
        if not has_more:
            break
        _log(f"  → Page {page_idx + 1}: fetching more items...")
        prompt = PART_EXTRA_PROMPT.format(page_num=page_idx + 1)
        for _ in range(2):
            try:
                result    = parse_json(call_claude(prompt, page_to_b64(pdf_path, page_idx), max_tokens=4096))
                new_items = result.get("items", [])
                has_more  = result.get("has_more_items", False)
                items.extend(new_items)
                break
            except (ValueError, json.JSONDecodeError):
                has_more = False

    return summary, val_dtls, items


# ── 4. EXCEL OUTPUT ───────────────────────────────────────────────────────────

def write_to_sheet(ws, summary: dict, val_dtls: dict, items: list, start_row: int) -> int:
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    def cell_style(cell, value, bold=False, bg=None, font_color="000000", align=None, nfmt=None):
        cell.value = value
        cell.font  = Font(name="Arial", bold=bold, size=10, color=font_color)
        if bg:
            cell.fill = PatternFill("solid", start_color=bg)
        if align:
            cell.alignment = align
        if nfmt:
            cell.number_format = nfmt

    hr = start_row
    ws.row_dimensions[hr].height = 13.9
    for col_idx, label, bg in [
        (1,"Port Code","DEEAF1"),(2,"SB No",None),(3,"SB Date","DEEAF1"),
        (6,"FOB Value","DEEAF1"),(7,"Freight",None),(8,"Insurance","DEEAF1"),
        (9,"Discount",None),(10,"COM","DEEAF1"),(11,"Deductions",None),
        (12,"P/C","DEEAF1"),(13,"Duty",None),(14,"Cess","DEEAF1"),
        (17,"DBK Claim","DEEAF1"),(18,"IGST AMT",None),(19,"Cess AMT","DEEAF1"),
        (20,"IGST Value",None),(21,"RODTEP AMT","DEEAF1"),(22,"ROSCTL AMT",None),
        (25,"INV No","DEEAF1"),(26,"INV AMT",None),(27,"Currency","DEEAF1"),
        (30,"Invoice Value","DEEAF1"),(31,"Invoice Currency",None),(32,"FOB Value","DEEAF1"),
        (33,"FOB Currency",None),(34,"Freight","DEEAF1"),(35,"Insurance",None),
        (36,"Discount","DEEAF1"),(37,"Commission",None),(38,"Deduct","DEEAF1"),
        (39,"P/C",None),(40,"Exchange Rate","DEEAF1"),
    ]:
        cell_style(ws.cell(row=hr, column=col_idx), label, bold=True, bg=bg)

    vr = start_row + 1
    for col_idx, value in [
        (1,summary.get("port_code")),(2,summary.get("sb_no")),(3,summary.get("sb_date")),
        (6,summary.get("fob_value")),(7,summary.get("freight")),(8,summary.get("insurance")),
        (9,summary.get("discount")),(10,summary.get("com")),(11,summary.get("deductions")),
        (12,summary.get("pc")),(13,summary.get("duty")),(14,summary.get("cess")),
        (17,summary.get("dbk_claim")),(18,summary.get("igst_amt")),(19,summary.get("cess_amt")),
        (20,summary.get("igst_value")),(21,summary.get("rodtep_amt")),(22,summary.get("rosctl_amt")),
        (25,summary.get("inv_no")),(26,summary.get("inv_amt")),(27,summary.get("currency")),
        (30,val_dtls.get("invoice_value")),(31,val_dtls.get("invoice_currency")),
        (32,val_dtls.get("fob_value")),(33,val_dtls.get("fob_currency")),
        (34,val_dtls.get("freight")),(35,val_dtls.get("insurance")),
        (36,val_dtls.get("discount")),(37,val_dtls.get("commission")),
        (38,val_dtls.get("deduct")),(39,val_dtls.get("pc")),(40,val_dtls.get("exchange_rate")),
    ]:
        cell_style(ws.cell(row=vr, column=col_idx), value, bg="FFFF00")

    ihr = start_row + 2
    ws.row_dimensions[ihr].height = 26.25
    cell_style(ws.cell(row=ihr, column=42), "SB No", bold=True)
    for i, hdr in enumerate(["HS Code","Description","Quantity","Rate (USD)","Value F/C (USD)"]):
        cell_style(ws.cell(row=ihr, column=43+i), hdr, bold=True, bg="2E75B6", font_color="FFFFFF", align=center)

    sb_no = summary.get("sb_no", "")
    for idx, item in enumerate(items):
        r = start_row + 3 + idx
        bg = "DEEAF1" if idx % 2 == 0 else None
        cell_style(ws.cell(row=r, column=42), sb_no, bg="FFFF00")
        for i, (val, aln, nfmt) in enumerate(zip(
            [item.get("hs_code"), item.get("description"), item.get("quantity"), item.get("rate"), item.get("value_fc")],
            [center, left, center, center, center],
            [None, None, "#,##0", "#,##0.000", "#,##0.00"]
        )):
            cell_style(ws.cell(row=r, column=43+i), val, bg=bg, align=aln, nfmt=nfmt)

    return start_row + 3 + len(items) + 1


def build_workbook(all_results, skipped, failures, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Customs Extract"
    for col, width in {
        "A":8.9375,"B":7.25,"C":9.25,"F":9.625,"G":6.5625,"H":8.8125,"I":8.0625,
        "J":4.5625,"K":10.0625,"L":3.5625,"M":4.5,"N":4.75,"Q":9.75,"R":8.875,
        "S":9.8125,"T":9.8125,"U":12.0625,"V":12.0,"Y":8.5625,"Z":7.6875,"AA":8.25,
        "AD":11.6875,"AE":14.75,"AF":9.375,"AG":12.4375,"AH":6.5625,"AI":8.8125,
        "AJ":8.0625,"AK":11.25,"AL":6.5,"AM":3.5625,"AN":13.0,
    }.items():
        ws.column_dimensions[col].width = width

    current_row = 1
    for r in all_results:
        current_row = write_to_sheet(ws, r["summary"], r["val_dtls"], r["items"], current_row)

    we = wb.create_sheet("Errors")
    we.column_dimensions["A"].width = 40
    we.column_dimensions["B"].width = 18
    we.column_dimensions["C"].width = 50
    hdr_fill = PatternFill("solid", start_color="C00000")
    hdr_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    row_fill = PatternFill("solid", start_color="FFE0E0")
    nrm_font = Font(name="Arial", size=10)
    ctr = Alignment(horizontal="center", vertical="center")
    lft = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for col, label in enumerate(["File","Type","Reason"], 1):
        c = we.cell(row=1, column=col, value=label)
        c.font = hdr_font; c.fill = hdr_fill; c.alignment = ctr
    we.row_dimensions[1].height = 18
    row = 2
    for filename, reason in skipped:
        for col, val in enumerate([filename,"Duplicate",reason], 1):
            c = we.cell(row=row, column=col, value=val)
            c.font = nrm_font; c.fill = row_fill; c.alignment = lft
        we.row_dimensions[row].height = 16; row += 1
    for filename, reason in failures:
        for col, val in enumerate([filename,"Error",reason], 1):
            c = we.cell(row=row, column=col, value=val)
            c.font = nrm_font; c.alignment = lft
        we.row_dimensions[row].height = 16; row += 1

    wb.save(output_path)


# ── 5. MAIN ───────────────────────────────────────────────────────────────────

def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from datetime import datetime
    import threading, queue

    win = tk.Tk()
    win.withdraw()

    pdf_paths = filedialog.askopenfilenames(
        title="Select Customs PDF files to process",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )
    if not pdf_paths:
        win.destroy(); sys.exit(0)

    output_dir = filedialog.askdirectory(title="Select folder to save Excel output")
    if not output_dir:
        win.destroy(); sys.exit(0)

    win.title("Customs PDF Extractor")
    win.geometry("500x320")
    win.resizable(False, False)
    win.deiconify()

    tk.Label(win, text="Processing PDFs...", font=("Arial", 12, "bold")).pack(pady=10)
    progress_bar = ttk.Progressbar(win, length=460, mode="determinate", maximum=len(pdf_paths))
    progress_bar.pack(pady=5, padx=20)
    status_label = tk.Label(win, text="", font=("Arial", 9), wraplength=460)
    status_label.pack(pady=2)
    log_box = tk.Text(win, height=12, font=("Courier", 8), state="disabled")
    log_box.pack(pady=5, padx=20, fill="both", expand=True)

    msg_queue = queue.Queue()

    def poll_queue():
        try:
            while True:
                msg = msg_queue.get_nowait()
                if msg["type"] == "log":
                    log_box.config(state="normal")
                    log_box.insert("end", msg["text"] + "\n")
                    log_box.see("end")
                    log_box.config(state="disabled")
                elif msg["type"] == "status":
                    status_label.config(text=msg["text"])
                elif msg["type"] == "progress":
                    progress_bar["value"] = msg["value"]
                elif msg["type"] == "done":
                    status_label.config(text="Done!")
                    messagebox.showinfo("Complete", msg["text"])
                    win.destroy(); sys.exit(0)
        except queue.Empty:
            pass
        win.after(100, poll_queue)

    def worker():
        def log(text): msg_queue.put({"type": "log", "text": text})

        all_results, failures, skipped = [], [], []
        seen_sb_nos, seen_inv_nos = set(), set()

        for i, pdf_path in enumerate(pdf_paths):
            filename = os.path.basename(pdf_path)
            msg_queue.put({"type": "status", "text": f"Processing: {filename}"})
            log(f"\n[{i+1}/{len(pdf_paths)}] {filename}")
            try:
                summary, val_dtls, items = extract_from_pdf(pdf_path, log=log)

                sb_no  = str(summary.get("sb_no")  or "").strip()
                inv_no = str(summary.get("inv_no") or "").strip()
                dup = None
                if sb_no  and sb_no  in seen_sb_nos:  dup = f"duplicate SB No ({sb_no})"
                elif inv_no and inv_no in seen_inv_nos: dup = f"duplicate Invoice No ({inv_no})"
                if dup:
                    log(f"  ⚠️  SKIPPED — {dup}")
                    skipped.append((filename, dup))
                    msg_queue.put({"type": "progress", "value": i+1}); continue

                if sb_no:  seen_sb_nos.add(sb_no)
                if inv_no: seen_inv_nos.add(inv_no)
                all_results.append({"summary": summary, "val_dtls": val_dtls, "items": items})
                log(f"  ✅ Extracted {len(items)} items")
            except Exception as e:
                log(f"  ❌ ERROR: {e}")
                failures.append((filename, str(e)))
            msg_queue.put({"type": "progress", "value": i+1})

        msg_queue.put({"type": "status", "text": "Building Excel..."})
        timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"customs_extracted_{timestamp}.xlsx")
        log(f"\n→ Writing {len(all_results)} PDF(s) ...")
        try:
            build_workbook(all_results, skipped, failures, output_path)
            log(f"✅ Saved: {output_path}")
        except Exception as e:
            log(f"❌ Failed to save Excel: {e}")

        summary_msg = (
            f"✅ {len(all_results)} PDF(s) extracted\n"
            f"⚠️  {len(skipped)} skipped (duplicates)\n"
            f"❌ {len(failures)} failed\n\nSaved to:\n{output_dir}"
        )
        if skipped or failures:
            summary_msg += "\n\nSee the 'Errors' sheet for details."
        log("\n" + "=" * 50)
        log(summary_msg)
        msg_queue.put({"type": "done", "text": summary_msg})

    threading.Thread(target=worker, daemon=True).start()
    win.after(100, poll_queue)
    win.protocol("WM_DELETE_WINDOW", lambda: (win.destroy(), sys.exit(0)))
    win.mainloop()


if __name__ == "__main__":
    main()