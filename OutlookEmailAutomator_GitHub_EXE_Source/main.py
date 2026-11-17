"""
Outlook Email Automator - Windows Desktop Application
=====================================================
Features:
- Persistent Memory: Automatically saves Routing (To/CC), Subject & Body templates,
  Font settings, Table styles, Signature, and Custom Logo to local disk
  (%USERPROFILE%\\.outlook_email_automator\\profile.json).
- Merged-Cell & Multi-Block Excel Parser: Reads .xlsx files using openpyxl, unrolls
  vertical/horizontal merged cells, skips repeated block headers/totals, and captures
  footer notes (e.g., "Did not Select any RS PB").
- Direct Outlook Integration: Opens drafts directly in Microsoft Outlook on Windows
  (via pywin32) OR exports .eml / batch .zip files.
"""

import os
import re
import json
import copy
import base64
import zipfile
import datetime
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, colorchooser, scrolledtext

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False


APP_DIR = os.path.join(os.path.expanduser("~"), ".outlook_email_automator")
PROFILE_PATH = os.path.join(APP_DIR, "profile.json")

DEFAULT_PROFILE = {
    "font_family": "Calibri",
    "body_font_size": "11",
    "table_font_size": "10.5",
    "merge_visual_cells": True,
    "include_footer_notes": True,
    "include_signature": True,
    "subject_template": "RE: Shelving Status of Ventrax HFA Inhaler {shortDate} of {region}",
    "body_template": (
        "Dear Vai,\n\n"
        "Please find below the MPE-wise shelving status of Ventrax® HFA Inhaler (VHI) as of <b>{longDate}</b>"
    ),
    "signature_html": (
        '<p style="margin: 0 0 8px 0; color: #1f4e79; font-weight: bold;">Thanks and regards,</p>'
        '<p style="margin: 0; color: #595959;">'
        '<span style="color: #2e75b6; font-weight: bold;">S.M. Masrakul Islam Medha</span> | '
        'Brand Executive, Strategic Brand Management | Beximco Pharmaceuticals PLC. | '
        '19 Dhanmondi R/A, Road #7, Dhaka 1205 | Bangladesh</p>'
        '<p style="margin: 3px 0 0 0; color: #595959;">'
        '| <strong>Cell</strong> +88 <a href="tel:+8801760021915" style="color: #954f72;">01760021915</a> '
        '| <strong>Fax</strong> +880 2 58613888 '
        '| <a href="mailto:mke@bpl.net" style="color: #0563c1;">mke@bpl.net</a> '
        '| <a href="http://www.beximcopharma.com" style="color: #0563c1;">www.beximcopharma.com</a></p>'
    ),
    "custom_logo_base64": "",
    "last_excel_path": "",
    "routing_data": {
        "R4": {
            "to": "PSL-(NAL)-Md. Nur-A- Alam",
            "cc": "SBM-(ALM)-Mahabub Alam.; PSL-(KNH)-Kazi Nurul Hoque; PSL-(MDR)-Masud Rana khan; PSL-(ANL)-Ahsanul Kabir;"
        },
        "A2": {
            "to": "PSL-(MZO)-Md. Zakir Hossain",
            "cc": "SBM-(ALM)-Mahabub Alam.; PSL-(KNH)-Kazi Nurul Hoque; PSL-(MDR)-Masud Rana khan; PSL-(ANL)-Ahsanul Kabir;"
        },
        "N4": {
            "to": "PSL-(KMO)-Kazi Md. Mostafa",
            "cc": "SBM-(RML)-Noor-Ul Azam Khan Romel; SBM-(ALM)-Mahabub Alam.; PSL-(KNH)-Kazi Nurul Hoque;"
        }
    }
}

SAMPLE_TABLE_DATA = {
    "R4": [
        {"ASE": "R410", "T.Code": "R411", "Shelving Till 15": 1, "Shelving Till 19": 3, "Shelving Till 25": 14, "Shelving Till 28": 14, "New Shelving": 0, "Remarks": ""},
        {"ASE": "R410", "T.Code": "R412", "Shelving Till 15": 1, "Shelving Till 19": 7, "Shelving Till 25": 7, "Shelving Till 28": 9, "New Shelving": 2, "Remarks": "Increased but still below 10. Need more shelving."},
        {"ASE": "R410", "T.Code": "R413", "Shelving Till 15": 3, "Shelving Till 19": 4, "Shelving Till 25": 10, "Shelving Till 28": 10, "New Shelving": 0, "Remarks": ""},
        {"ASE": "R420", "T.Code": "R421", "Shelving Till 15": 4, "Shelving Till 19": 4, "Shelving Till 25": 12, "Shelving Till 28": 12, "New Shelving": 0, "Remarks": ""},
        {"ASE": "R420", "T.Code": "R422", "Shelving Till 15": 3, "Shelving Till 19": 3, "Shelving Till 25": 10, "Shelving Till 28": 10, "New Shelving": 0, "Remarks": ""},
        {"ASE": "R420", "T.Code": "R423", "Shelving Till 15": 4, "Shelving Till 19": 4, "Shelving Till 25": 5, "Shelving Till 28": 13, "New Shelving": 8, "Remarks": ""},
        {"ASE": "R430", "T.Code": "R431", "Shelving Till 15": 8, "Shelving Till 19": 14, "Shelving Till 25": 18, "Shelving Till 28": 18, "New Shelving": 0, "Remarks": ""},
        {"ASE": "R430", "T.Code": "R432", "Shelving Till 15": 2, "Shelving Till 19": 4, "Shelving Till 25": 6, "Shelving Till 28": 8, "New Shelving": 2, "Remarks": "Increased but still below 10. Need more shelving."}
    ],
    "A2": [
        {"ASE": "A210", "T.Code": "A211", "Shelving Till 15": 5, "Shelving Till 19": 6, "Shelving Till 25": 6, "Shelving Till 28": 8, "New Shelving": 2, "Remarks": "Increased but still below 10. Need more shelving."},
        {"ASE": "A210", "T.Code": "A212", "Shelving Till 15": 9, "Shelving Till 19": 9, "Shelving Till 25": 13, "Shelving Till 28": 18, "New Shelving": 5, "Remarks": ""},
        {"ASE": "A240", "T.Code": "A243", "Shelving Till 15": 6, "Shelving Till 19": 6, "Shelving Till 25": 6, "Shelving Till 28": 6, "New Shelving": 0, "Remarks": "Has taken 0 steps. RSE & ASE vai, immediate Attention needed", "_isRedRow": True}
    ],
    "N4": [
        {"ASE": "N410", "T.Code": "N411", "Shelving Till 15": 12, "Shelving Till 19": 18, "Shelving Till 25": 20, "Shelving Till 28": 22, "New Shelving": 2, "Remarks": ""},
        {"ASE": "N420", "T.Code": "N421", "Shelving Till 15": 4, "Shelving Till 19": 10, "Shelving Till 25": 14, "Shelving Till 28": 12, "New Shelving": -2, "Remarks": ""}
    ]
}


def load_profile():
    os.makedirs(APP_DIR, exist_ok=True)
    if os.path.exists(PROFILE_PATH):
        try:
            with open(PROFILE_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            merged = copy.deepcopy(DEFAULT_PROFILE)
            merged.update(saved)
            return merged
        except Exception:
            pass
    return copy.deepcopy(DEFAULT_PROFILE)


def save_profile(profile_data):
    os.makedirs(APP_DIR, exist_ok=True)
    with open(PROFILE_PATH, "w", encoding="utf-8") as f:
        json.dump(profile_data, f, indent=2, ensure_ascii=False)


def parse_excel_workbook(file_path, sheet_name=None):
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl is not installed.")

    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet_names = wb.sheetnames
    active_sheet_name = sheet_name if (sheet_name and sheet_name in sheet_names) else sheet_names[0]
    ws = wb[active_sheet_name]

    max_r = ws.max_row or 0
    max_c = ws.max_column or 0
    if max_r == 0 or max_c == 0:
        return {}, {}, sheet_names, active_sheet_name, False, {}

    grid = []
    for r in range(1, max_r + 1):
        row_vals = []
        for c in range(1, max_c + 1):
            val = ws.cell(row=r, column=c).value
            if val is None:
                row_vals.append("")
            elif isinstance(val, str):
                row_vals.append(val.strip())
            else:
                row_vals.append(val)
        grid.append(row_vals)

    has_vertical_merges = False
    for merged_range in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        if max_row > min_row:
            has_vertical_merges = True
        top_left_val = grid[min_row - 1][min_col - 1]
        for r in range(min_row - 1, max_row):
            for c in range(min_col - 1, max_col):
                if r == (min_row - 1) and c > (min_col - 1) and max_row == min_row:
                    continue
                grid[r][c] = top_left_val

    header_row_idx = -1
    for idx, row_vals in enumerate(grid):
        lowered = [str(v).lower().strip() for v in row_vals]
        if "rse" in lowered or "region" in lowered:
            header_row_idx = idx
            break

    if header_row_idx == -1:
        raise ValueError("Could not find a header row containing 'RSE' or 'Region'.")

    raw_headers = grid[header_row_idx]
    valid_col_indices = []
    headers = []
    for idx, h in enumerate(raw_headers):
        clean_h = str(h).strip()
        if clean_h != "":
            valid_col_indices.append(idx)
            headers.append(clean_h)

    rse_idx = next((i for i, h in enumerate(headers) if h.lower() in ("rse", "region")), -1)
    ase_idx = next((i for i, h in enumerate(headers) if h.lower() == "ase"), -1)
    mpe_idx = next((i for i, h in enumerate(headers) if h.lower() in ("mpe", "t.code", "territory")), -1)
    spec_idx = next((i for i, h in enumerate(headers) if "specialty" in h.lower()), -1)

    parsed_tables = {}
    parsed_notes = {}
    extracted_routing = {}
    last_region = ""
    last_ase = ""
    last_spec = ""

    for r in range(header_row_idx + 1, len(grid)):
        row_cells = [grid[r][c] for c in valid_col_indices]
        if all(str(v).strip() == "" for v in row_cells):
            continue

        first_str = str(row_cells[0]).strip()
        first_lower = first_str.lower()

        if first_lower in ("rse", "region"):
            last_region = ""
            last_ase = ""
            last_spec = ""
            continue

        if first_lower == "total" or first_lower.endswith(" total"):
            continue

        if ("did not select" in first_lower or first_lower.startswith("note") or
                (rse_idx == 0 and len(first_str) > 12 and last_region)):
            note_val = next((str(v).strip() for v in row_cells[1:] if str(v).strip() != ""), "")
            if last_region:
                parsed_notes.setdefault(last_region, []).append({
                    "label": first_str,
                    "value": note_val
                })
            continue

        region_val = str(row_cells[rse_idx]).strip() if rse_idx != -1 else ""
        mpe_val = str(row_cells[mpe_idx]).strip() if mpe_idx != -1 else ""

        if not region_val and mpe_val and re.match(r"^[A-Za-z]\d{2,}", mpe_val):
            prefix = mpe_val[:2].upper()
            region_val = last_region if (last_region and prefix == last_region.upper()) else prefix

        if not region_val:
            continue

        if region_val != last_region:
            last_ase = ""
            last_spec = ""
        last_region = region_val

        row_dict = {}
        for i, h in enumerate(headers):
            val = row_cells[i]
            if i == ase_idx:
                if str(val).strip() != "":
                    last_ase = str(val).strip()
                elif last_ase and mpe_val and mpe_val.upper().startswith(last_ase[:3].upper()):
                    val = last_ase
            if i == spec_idx:
                if str(val).strip() != "":
                    last_spec = str(val).strip()
                elif last_spec:
                    val = last_spec
            row_dict[h] = val

        to_val = row_dict.get("To") or row_dict.get("TO")
        if to_val:
            extracted_routing.setdefault(region_val, {})["to"] = str(to_val).strip()

        cc_parts = []
        if row_dict.get("CC: ZSI"):
            cc_parts.append(str(row_dict["CC: ZSI"]).strip())
        if row_dict.get("Fixed CC"):
            cc_parts.append(str(row_dict["Fixed CC"]).strip())
        if cc_parts:
            extracted_routing.setdefault(region_val, {})["cc"] = "; ".join(cc_parts)
        elif row_dict.get("CC") or row_dict.get("Cc"):
            extracted_routing.setdefault(region_val, {})["cc"] = str(row_dict.get("CC") or row_dict.get("Cc")).strip()

        remarks_lower = str(row_dict.get("Remarks", "")).lower()
        row_dict["_isRedRow"] = ("0 steps" in remarks_lower) or ("immediate attention" in remarks_lower)

        parsed_tables.setdefault(region_val, []).append(row_dict)

    return parsed_tables, parsed_notes, sheet_names, active_sheet_name, has_vertical_merges, extracted_routing


def calculate_rowspans(rows, group_cols):
    spans = {}
    for col in group_cols:
        col_spans = [1] * len(rows)
        start_idx = 0
        for i in range(1, len(rows) + 1):
            prev_val = str(rows[start_idx].get(col, "")).strip()
            curr_val = str(rows[i].get(col, "")).strip() if i < len(rows) else None
            if curr_val is not None and curr_val != "" and curr_val == prev_val:
                col_spans[start_idx] += 1
                col_spans[i] = 0
            else:
                start_idx = i
        spans[col] = col_spans
    return spans


def build_table_html(region, rows, notes, font_family, table_font_size, merge_cells=True, include_notes=True):
    if not rows:
        return "<p style='color:#dc2626; font-style:italic;'>No data available for this region.</p>"

    exclude = {"to", "cc", "_isredrow", "cc: zsi", "fixed cc"}
    columns = [k for k in rows[0].keys() if k.lower().strip() not in exclude and not k.startswith("_")]

    group_cols = [c for c in columns if re.match(r"^(rse|ase|specialty)$", c.strip(), re.IGNORECASE)]
    rowspans = calculate_rowspans(rows, group_cols) if merge_cells else {}

    flagged_mpes = set()
    for n in notes:
        if "did not select" in n.get("label", "").lower():
            for code in re.split(r"[,;\s]+", n.get("value", "")):
                if code.strip():
                    flagged_mpes.add(code.strip().upper())

    totals = {col: 0.0 for col in columns}
    is_numeric = {}
    for col in columns:
        lower = col.lower()
        if any(tag in lower for tag in ("code", "ase", "rse", "mpe", "specialty", "remarks")):
            is_numeric[col] = False
            continue
        has_num = False
        all_valid = True
        for r in rows:
            v = r.get(col, "")
            if v == "" or v is None:
                continue
            try:
                float(v)
                has_num = True
            except (ValueError, TypeError):
                all_valid = False
                break
        is_numeric[col] = has_num and all_valid

    table_style = (
        f"border-collapse: collapse; width: auto; border: 1px solid #000000; "
        f"font-family: '{font_family}', Calibri, Arial, sans-serif; font-size: {table_font_size}pt;"
    )
    th_style = (
        "background-color: #e6005c; color: #ffffff; border: 1px solid #000000; "
        "padding: 4px 8px; text-align: center; font-weight: bold;"
    )

    html = [f'<table style="{table_style}"><thead><tr>']
    for col in columns:
        html.append(f'<th style="{th_style}">{col}</th>')
    html.append("</tr></thead><tbody>")

    for r_idx, row in enumerate(rows):
        is_red_row = bool(row.get("_isRedRow"))
        html.append("<tr>")
        for col in columns:
            if merge_cells and col in rowspans and rowspans[col][r_idx] == 0:
                continue
            span_attr = ""
            if merge_cells and col in rowspans and rowspans[col][r_idx] > 1:
                span_attr = f' rowspan="{rowspans[col][r_idx]}"'

            val = row.get(col, "")
            if val is None:
                val = ""
            lower = col.lower()

            bg = "#ff0000" if is_red_row else "#fdf6e3"
            fg = "#ffffff" if is_red_row else "#000000"
            fw = "bold" if is_red_row else "normal"
            align = "left" if lower == "remarks" else "center"
            pad = "4px 10px" if lower == "remarks" else "4px 8px"

            if not is_red_row:
                if "new shelving" in lower or lower == "new":
                    try:
                        num = float(val)
                        fg = "#00b050" if num > 0 else "#ff0000"
                        fw = "bold"
                    except (ValueError, TypeError):
                        pass
                if lower == "mpe" and str(val).strip().upper() in flagged_mpes:
                    fg = "#ff0000"
                    fw = "bold"

            if is_numeric[col] and val != "":
                try:
                    totals[col] += float(val)
                except (ValueError, TypeError):
                    pass

            td_style = (
                f"border: 1px solid #000000; padding: {pad}; text-align: {align}; "
                f"background-color: {bg}; color: {fg}; font-weight: {fw}; vertical-align: middle;"
            )
            html.append(f'<td{span_attr} style="{td_style}">{val}</td>')
        html.append("</tr>")

    html.append("<tr>")
    for idx, col in enumerate(columns):
        tot_style = (
            "border: 1px solid #000000; padding: 4px 8px; text-align: center; "
            "background-color: #ffff00; color: #000000; font-weight: bold;"
        )
        if idx == 0:
            html.append(f'<td style="{tot_style}">{region} Total</td>')
        elif is_numeric[col]:
            tot_val = int(totals[col]) if totals[col].is_integer() else round(totals[col], 2)
            html.append(f'<td style="{tot_style}">{tot_val}</td>')
        else:
            html.append(f'<td style="{tot_style}"></td>')
    html.append("</tr>")

    if include_notes and notes and len(columns) >= 2:
        label_span = min(3, len(columns) - 1)
        val_span = len(columns) - label_span
        for note in notes:
            lbl_style = (
                "border: 1px solid #000000; padding: 4px 8px; text-align: center; "
                "background-color: #ffff00; color: #000000; font-weight: bold;"
            )
            val_style = (
                "border: 1px solid #000000; padding: 4px 8px; text-align: center; "
                "background-color: #ffffff; color: #1f4e79; font-weight: bold;"
            )
            html.append(
                f'<tr><td colspan="{label_span}" style="{lbl_style}">{note["label"]}</td>'
                f'<td colspan="{val_span}" style="{val_style}">{note["value"]}</td></tr>'
            )

    html.append("</tbody></table>")
    return "".join(html)


def build_full_email_html(profile, region, short_date, long_date, table_rows, footer_notes):
    font_family = profile.get("font_family", "Calibri")
    body_size = profile.get("body_font_size", "11")
    table_size = profile.get("table_font_size", "10.5")

    body_raw = profile.get("body_template", "")
    body_html = (
        body_raw.replace("{region}", region)
        .replace("{shortDate}", short_date)
        .replace("{longDate}", long_date)
        .replace("\n", "<br>")
    )

    table_html = build_table_html(
        region=region,
        rows=table_rows,
        notes=footer_notes,
        font_family=font_family,
        table_font_size=table_size,
        merge_cells=profile.get("merge_visual_cells", True),
        include_notes=profile.get("include_footer_notes", True),
    )

    sig_html = ""
    if profile.get("include_signature", True):
        sig_body = profile.get("signature_html", "")
        logo_b64 = profile.get("custom_logo_base64", "")
        logo_tag = ""
        if logo_b64:
            logo_tag = f'<div style="margin-top:8px;"><img src="{logo_b64}" alt="Logo" style="max-height:42px; width:auto;"></div>'
        sig_html = (
            f'<br><br><div style="font-family: \'{font_family}\', Calibri, Arial, sans-serif; '
            f'font-size: {body_size}pt; color: #595959; line-height: 1.45;">{sig_body}{logo_tag}</div>'
        )

    return (
        f'<!DOCTYPE html><html><head><meta charset="utf-8"></head>'
        f'<body style="font-family: \'{font_family}\', Calibri, Arial, sans-serif; font-size: {body_size}pt; color: #000000;">'
        f'<div>{body_html}<br><br></div>'
        f'<div>{table_html}</div>'
        f'{sig_html}'
        f'</body></html>'
    )


class OutlookAutomatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook Email Automator — Windows Desktop Edition (Auto-Saving Memory)")
        self.root.geometry("1220x780")
        self.root.minsize(1024, 680)
        self.root.configure(bg="#f1f5f9")

        self.profile = load_profile()
        self.table_data = copy.deepcopy(SAMPLE_TABLE_DATA)
        self.footer_notes = {}
        self.current_excel_path = self.profile.get("last_excel_path", "")
        self.sheet_names = []

        self._setup_styles()
        self._build_ui()
        self._init_dates()

        if self.current_excel_path and os.path.exists(self.current_excel_path):
            self._load_excel_file(self.current_excel_path, silent=True)

        self.refresh_region_dropdown()
        self.update_preview()

    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background="#f1f5f9")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("TLabel", background="#ffffff", foreground="#1e293b", font=("Segoe UI", 9))

    def _build_ui(self):
        top_bar = tk.Frame(self.root, bg="#0f172a", height=58, padx=16)
        top_bar.pack(fill=tk.X, side=tk.TOP)
        top_bar.pack_propagate(False)

        title_frame = tk.Frame(top_bar, bg="#0f172a")
        title_frame.pack(side=tk.LEFT, pady=8)
        tk.Label(title_frame, text="✉ Outlook Email Automator", bg="#0f172a", fg="#ffffff",
                 font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(title_frame, text=f"Memory Active • Saved automatically to {PROFILE_PATH}",
                 bg="#0f172a", fg="#38bdf8", font=("Segoe UI", 8)).pack(anchor="w")

        btn_bar = tk.Frame(top_bar, bg="#0f172a")
        btn_bar.pack(side=tk.RIGHT, pady=10)

        tk.Button(btn_bar, text="💾 Save Memory Now", command=self.manual_save_profile,
                  bg="#334155", fg="#ffffff", font=("Segoe UI", 9, "bold"), relief=tk.FLAT,
                  padx=10, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=4)

        tk.Button(btn_bar, text="📤 Open Draft in Outlook", command=self.open_in_outlook_single,
                  bg="#4f46e5", fg="#ffffff", font=("Segoe UI", 9, "bold"), relief=tk.FLAT,
                  padx=12, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=4)

        tk.Button(btn_bar, text="⚡ Batch Create All Outlook Drafts", command=self.open_in_outlook_batch,
                  bg="#059669", fg="#ffffff", font=("Segoe UI", 9, "bold"), relief=tk.FLAT,
                  padx=12, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=4)

        tk.Button(btn_bar, text="📦 Export .EML / .ZIP", command=self.export_eml_or_zip,
                  bg="#0284c7", fg="#ffffff", font=("Segoe UI", 9, "bold"), relief=tk.FLAT,
                  padx=10, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=4)

        self.status_var = tk.StringVar(value="Ready. All edits are remembered automatically.")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bg="#e2e8f0", fg="#1e293b",
                              font=("Segoe UI", 9), anchor="w", padx=14, pady=4)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, bg="#cbd5e1", sashwidth=5)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left_container = tk.Frame(paned, bg="#ffffff", bd=1, relief=tk.SOLID)
        paned.add(left_container, width=500, minsize=430)

        notebook = ttk.Notebook(left_container)
        notebook.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        tab_data = ttk.Frame(notebook, style="Card.TFrame")
        tab_template = ttk.Frame(notebook, style="Card.TFrame")
        tab_routing = ttk.Frame(notebook, style="Card.TFrame")
        tab_signature = ttk.Frame(notebook, style="Card.TFrame")

        notebook.add(tab_data, text=" 1. Data & Typography ")
        notebook.add(tab_template, text=" 2. Subject & Body ")
        notebook.add(tab_routing, text=" 3. Routing Memory ")
        notebook.add(tab_signature, text=" 4. Signature & Logo ")

        self._build_tab_data(tab_data)
        self._build_tab_template(tab_template)
        self._build_tab_routing(tab_routing)
        self._build_tab_signature(tab_signature)

        right_container = tk.Frame(paned, bg="#ffffff", bd=1, relief=tk.SOLID)
        paned.add(right_container, minsize=520)
        self._build_preview_panel(right_container)

    def _build_tab_data(self, parent):
        pad = {"padx": 12, "pady": 6}

        tk.Label(parent, text="EXCEL DATA SOURCE (MERGED CELLS SUPPORTED)", bg="#ffffff",
                 fg="#4f46e5", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(10, 4))

        file_row = tk.Frame(parent, bg="#ffffff")
        file_row.pack(fill=tk.X, **pad)
        tk.Button(file_row, text="📂 Select Excel Table (.xlsx)", command=self.browse_excel,
                  bg="#eef2ff", fg="#3730a3", font=("Segoe UI", 9, "bold"), relief=tk.GROOVE,
                  padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT)

        self.excel_label_var = tk.StringVar(
            value=os.path.basename(self.current_excel_path) if self.current_excel_path else "Using built-in sample data"
        )
        tk.Label(file_row, textvariable=self.excel_label_var, bg="#ffffff", fg="#475569",
                 font=("Segoe UI", 8, "italic")).pack(side=tk.LEFT, padx=10)

        sheet_row = tk.Frame(parent, bg="#ffffff")
        sheet_row.pack(fill=tk.X, **pad)
        tk.Label(sheet_row, text="Active Worksheet:", bg="#ffffff", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(sheet_row, textvariable=self.sheet_var, state="readonly", width=28)
        self.sheet_combo.pack(side=tk.LEFT, padx=8)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)

        ttk.Separator(parent, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=12, pady=8)

        tk.Label(parent, text="REGION & DATE PARAMETERS", bg="#ffffff",
                 fg="#4f46e5", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(2, 4))

        param_grid = tk.Frame(parent, bg="#ffffff")
        param_grid.pack(fill=tk.X, **pad)

        tk.Label(param_grid, text="Region (RSE):", bg="#ffffff", font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w")
        self.region_var = tk.StringVar(value="R4")
        self.region_combo = ttk.Combobox(param_grid, textvariable=self.region_var, state="readonly", width=12)
        self.region_combo.grid(row=1, column=0, padx=(0, 10), pady=2, sticky="w")
        self.region_combo.bind("<<ComboboxSelected>>", lambda e: self.on_region_changed())

        tk.Label(param_grid, text="Short Date ({shortDate}):", bg="#ffffff", font=("Segoe UI", 8, "bold")).grid(row=0, column=1, sticky="w")
        self.short_date_var = tk.StringVar()
        e_short = ttk.Entry(param_grid, textvariable=self.short_date_var, width=16)
        e_short.grid(row=1, column=1, padx=(0, 10), pady=2, sticky="w")
        e_short.bind("<KeyRelease>", lambda e: self.update_preview())

        tk.Label(param_grid, text="Long Date ({longDate}):", bg="#ffffff", font=("Segoe UI", 8, "bold")).grid(row=0, column=2, sticky="w")
        self.long_date_var = tk.StringVar()
        e_long = ttk.Entry(param_grid, textvariable=self.long_date_var, width=22)
        e_long.grid(row=1, column=2, pady=2, sticky="w")
        e_long.bind("<KeyRelease>", lambda e: self.update_preview())

        chk_frame = tk.Frame(parent, bg="#ffffff")
        chk_frame.pack(fill=tk.X, **pad)

        self.merge_cells_var = tk.BooleanVar(value=self.profile.get("merge_visual_cells", True))
        tk.Checkbutton(chk_frame, text="Merge repeated group cells (RSE / ASE / Specialty)",
                       variable=self.merge_cells_var, bg="#ffffff", font=("Segoe UI", 9),
                       command=self.on_settings_changed).pack(anchor="w")

        self.include_notes_var = tk.BooleanVar(value=self.profile.get("include_footer_notes", True))
        tk.Checkbutton(chk_frame, text="Include bottom summary note rows (e.g., 'Did not Select any RS PB')",
                       variable=self.include_notes_var, bg="#ffffff", font=("Segoe UI", 9),
                       command=self.on_settings_changed).pack(anchor="w")

        ttk.Separator(parent, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=12, pady=8)

        tk.Label(parent, text="GLOBAL EMAIL & TABLE TYPOGRAPHY (SAVED IN MEMORY)", bg="#ffffff",
                 fg="#4f46e5", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(2, 4))

        typo_grid = tk.Frame(parent, bg="#ffffff")
        typo_grid.pack(fill=tk.X, **pad)

        tk.Label(typo_grid, text="Email Font Family:", bg="#ffffff", font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w")
        self.font_family_var = tk.StringVar(value=self.profile.get("font_family", "Calibri"))
        cb_font = ttk.Combobox(
            typo_grid, textvariable=self.font_family_var, state="readonly", width=18,
            values=["Calibri", "Segoe UI", "Arial", "Aptos", "Verdana", "Tahoma", "Times New Roman"]
        )
        cb_font.grid(row=1, column=0, padx=(0, 10), pady=2, sticky="w")
        cb_font.bind("<<ComboboxSelected>>", lambda e: self.on_settings_changed())

        tk.Label(typo_grid, text="Body & Sig Size (pt):", bg="#ffffff", font=("Segoe UI", 8, "bold")).grid(row=0, column=1, sticky="w")
        self.body_size_var = tk.StringVar(value=self.profile.get("body_font_size", "11"))
        cb_bsize = ttk.Combobox(
            typo_grid, textvariable=self.body_size_var, state="readonly", width=12,
            values=["9", "10", "10.5", "11", "12", "13", "14"]
        )
        cb_bsize.grid(row=1, column=1, padx=(0, 10), pady=2, sticky="w")
        cb_bsize.bind("<<ComboboxSelected>>", lambda e: self.on_settings_changed())

        tk.Label(typo_grid, text="Table Font Size (pt):", bg="#ffffff", font=("Segoe UI", 8, "bold")).grid(row=0, column=2, sticky="w")
        self.table_size_var = tk.StringVar(value=self.profile.get("table_font_size", "10.5"))
        cb_tsize = ttk.Combobox(
            typo_grid, textvariable=self.table_size_var, state="readonly", width=12,
            values=["8.5", "9", "9.5", "10", "10.5", "11", "12"]
        )
        cb_tsize.grid(row=1, column=2, pady=2, sticky="w")
        cb_tsize.bind("<<ComboboxSelected>>", lambda e: self.on_settings_changed())

    def _build_tab_template(self, parent):
        tk.Label(parent, text="SUBJECT TEMPLATE", bg="#ffffff",
                 fg="#4f46e5", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(10, 2))

        self.subject_var = tk.StringVar(value=self.profile.get("subject_template", ""))
        subj_entry = ttk.Entry(parent, textvariable=self.subject_var, font=("Segoe UI", 9))
        subj_entry.pack(fill=tk.X, padx=12, pady=(0, 8))
        subj_entry.bind("<KeyRelease>", lambda e: self.on_template_changed())

        var_bar = tk.Frame(parent, bg="#f8fafc", bd=1, relief=tk.GROOVE, padx=6, pady=4)
        var_bar.pack(fill=tk.X, padx=12, pady=(0, 8))
        tk.Label(var_bar, text="Insert Variable:", bg="#f8fafc", fg="#475569",
                 font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 6))
        for var_tag in ("{region}", "{shortDate}", "{longDate}"):
            tk.Button(
                var_bar, text=var_tag, command=lambda t=var_tag: self.insert_body_snippet(t),
                bg="#eef2ff", fg="#4338ca", font=("Consolas", 8, "bold"), relief=tk.FLAT,
                padx=8, pady=2, cursor="hand2"
            ).pack(side=tk.LEFT, padx=3)

        tk.Label(parent, text="OPENING BODY TEXT (SUPPORTS HTML TAGS & TOOLBAR FORMATTING)", bg="#ffffff",
                 fg="#4f46e5", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(2, 2))

        fmt_bar = tk.Frame(parent, bg="#e2e8f0", padx=4, pady=3)
        fmt_bar.pack(fill=tk.X, padx=12)

        tk.Button(fmt_bar, text="B", font=("Segoe UI", 9, "bold"), width=3, relief=tk.FLAT, bg="#ffffff",
                  command=lambda: self.wrap_body_selection("<b>", "</b>")).pack(side=tk.LEFT, padx=2)
        tk.Button(fmt_bar, text="I", font=("Segoe UI", 9, "italic"), width=3, relief=tk.FLAT, bg="#ffffff",
                  command=lambda: self.wrap_body_selection("<i>", "</i>")).pack(side=tk.LEFT, padx=2)
        tk.Button(fmt_bar, text="U", font=("Segoe UI", 9, "underline"), width=3, relief=tk.FLAT, bg="#ffffff",
                  command=lambda: self.wrap_body_selection("<u>", "</u>")).pack(side=tk.LEFT, padx=2)
        tk.Button(fmt_bar, text="A+ (Larger)", font=("Segoe UI", 8, "bold"), relief=tk.FLAT, bg="#ffffff", padx=6,
                  command=lambda: self.wrap_body_selection('<span style="font-size:13pt;">', "</span>")).pack(side=tk.LEFT, padx=2)
        tk.Button(fmt_bar, text="A- (Smaller)", font=("Segoe UI", 8), relief=tk.FLAT, bg="#ffffff", padx=6,
                  command=lambda: self.wrap_body_selection('<span style="font-size:9.5pt;">', "</span>")).pack(side=tk.LEFT, padx=2)
        tk.Button(fmt_bar, text="🎨 Text Color", font=("Segoe UI", 8, "bold"), relief=tk.FLAT, bg="#ffffff", padx=6,
                  command=self.pick_body_text_color).pack(side=tk.LEFT, padx=2)
        tk.Button(fmt_bar, text="🖍 Highlight", font=("Segoe UI", 8, "bold"), relief=tk.FLAT, bg="#fef08a", padx=6,
                  command=lambda: self.wrap_body_selection('<span style="background-color:#ffff00;">', "</span>")).pack(side=tk.LEFT, padx=2)

        self.body_text = scrolledtext.ScrolledText(parent, height=10, font=("Consolas", 9), wrap=tk.WORD)
        self.body_text.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))
        self.body_text.insert("1.0", self.profile.get("body_template", ""))
        self.body_text.bind("<KeyRelease>", lambda e: self.on_template_changed())

    def _build_tab_routing(self, parent):
        tk.Label(parent, text="STORED ROUTING MEMORY (REMEMBERED PERMANENTLY)", bg="#ffffff",
                 fg="#4f46e5", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(10, 4))

        top_actions = tk.Frame(parent, bg="#ffffff")
        top_actions.pack(fill=tk.X, padx=12, pady=(0, 6))
        tk.Button(top_actions, text="📥 Import Routing Excel (.xlsx)", command=self.import_routing_excel,
                  bg="#eef2ff", fg="#3730a3", font=("Segoe UI", 8, "bold"), relief=tk.GROOVE,
                  padx=8, pady=4, cursor="hand2").pack(side=tk.LEFT)
        tk.Button(top_actions, text="➕ Add / Update Region", command=self.add_or_update_routing_row,
                  bg="#ecfdf5", fg="#065f46", font=("Segoe UI", 8, "bold"), relief=tk.GROOVE,
                  padx=8, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=6)
        tk.Button(top_actions, text="🗑 Delete Selected", command=self.delete_routing_row,
                  bg="#fef2f2", fg="#991b1b", font=("Segoe UI", 8, "bold"), relief=tk.GROOVE,
                  padx=8, pady=4, cursor="hand2").pack(side=tk.LEFT)

        cols = ("RSE", "To", "Cc")
        self.routing_tree = ttk.Treeview(parent, columns=cols, show="headings", height=9)
        self.routing_tree.heading("RSE", text="Region (RSE)")
        self.routing_tree.heading("To", text="To Recipient")
        self.routing_tree.heading("Cc", text="Cc Recipients")
        self.routing_tree.column("RSE", width=80, anchor="center")
        self.routing_tree.column("To", width=160)
        self.routing_tree.column("Cc", width=220)
        self.routing_tree.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)
        self.routing_tree.bind("<<TreeviewSelect>>", self.on_routing_tree_select)

        edit_box = tk.Frame(parent, bg="#f8fafc", bd=1, relief=tk.GROOVE, padx=8, pady=8)
        edit_box.pack(fill=tk.X, padx=12, pady=(4, 12))

        tk.Label(edit_box, text="RSE:", bg="#f8fafc", font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w")
        self.rt_rse_var = tk.StringVar()
        ttk.Entry(edit_box, textvariable=self.rt_rse_var, width=10).grid(row=0, column=1, padx=4, pady=2, sticky="w")

        tk.Label(edit_box, text="To:", bg="#f8fafc", font=("Segoe UI", 8, "bold")).grid(row=0, column=2, sticky="w")
        self.rt_to_var = tk.StringVar()
        ttk.Entry(edit_box, textvariable=self.rt_to_var, width=32).grid(row=0, column=3, padx=4, pady=2, sticky="ew")

        tk.Label(edit_box, text="Cc:", bg="#f8fafc", font=("Segoe UI", 8, "bold")).grid(row=1, column=0, sticky="w")
        self.rt_cc_var = tk.StringVar()
        ttk.Entry(edit_box, textvariable=self.rt_cc_var, width=48).grid(row=1, column=1, columnspan=3, padx=4, pady=2, sticky="ew")

        self._reload_routing_tree()

    def _build_tab_signature(self, parent):
        top_sig = tk.Frame(parent, bg="#ffffff")
        top_sig.pack(fill=tk.X, padx=12, pady=(10, 4))

        self.include_sig_var = tk.BooleanVar(value=self.profile.get("include_signature", True))
        tk.Checkbutton(top_sig, text="Append Signature to All Emails", variable=self.include_sig_var,
                       bg="#ffffff", font=("Segoe UI", 9, "bold"),
                       command=self.on_settings_changed).pack(side=tk.LEFT)

        logo_row = tk.Frame(parent, bg="#f8fafc", bd=1, relief=tk.GROOVE, padx=8, pady=6)
        logo_row.pack(fill=tk.X, padx=12, pady=4)

        tk.Button(logo_row, text="🖼 Upload Brand Logo Image", command=self.upload_logo_image,
                  bg="#eef2ff", fg="#3730a3", font=("Segoe UI", 8, "bold"), relief=tk.GROOVE,
                  padx=8, pady=3, cursor="hand2").pack(side=tk.LEFT)

        tk.Button(logo_row, text="✖ Remove Logo", command=self.clear_logo_image,
                  bg="#ffffff", fg="#dc2626", font=("Segoe UI", 8), relief=tk.GROOVE,
                  padx=8, pady=3, cursor="hand2").pack(side=tk.LEFT, padx=6)

        self.logo_status_var = tk.StringVar(
            value="Logo saved in memory ✔" if self.profile.get("custom_logo_base64") else "No custom logo uploaded"
        )
        tk.Label(logo_row, textvariable=self.logo_status_var, bg="#f8fafc", fg="#475569",
                 font=("Segoe UI", 8, "italic")).pack(side=tk.LEFT, padx=6)

        tk.Label(parent, text="SIGNATURE HTML / TEXT (AUTO-SAVED TO MEMORY)", bg="#ffffff",
                 fg="#4f46e5", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(6, 2))

        self.sig_text = scrolledtext.ScrolledText(parent, height=11, font=("Consolas", 9), wrap=tk.WORD)
        self.sig_text.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))
        self.sig_text.insert("1.0", self.profile.get("signature_html", ""))
        self.sig_text.bind("<KeyRelease>", lambda e: self.on_signature_changed())

    def _build_preview_panel(self, parent):
        header = tk.Frame(parent, bg="#1e293b", height=34, padx=12)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(header, text="LIVE OUTLOOK DRAFT & TABLE PREVIEW", bg="#1e293b", fg="#f8fafc",
                 font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, pady=6)

        env = tk.Frame(parent, bg="#f8fafc", padx=12, pady=8, bd=1, relief=tk.GROOVE)
        env.pack(fill=tk.X, padx=10, pady=8)

        tk.Label(env, text="To:", bg="#f8fafc", fg="#64748b", font=("Segoe UI", 9, "bold"), width=8, anchor="w").grid(row=0, column=0, pady=2)
        self.preview_to_var = tk.StringVar()
        e_to = ttk.Entry(env, textvariable=self.preview_to_var, font=("Segoe UI", 9))
        e_to.grid(row=0, column=1, sticky="ew", pady=2)
        e_to.bind("<FocusOut>", lambda e: self._sync_envelope_to_memory())

        tk.Label(env, text="Cc:", bg="#f8fafc", fg="#64748b", font=("Segoe UI", 9, "bold"), width=8, anchor="w").grid(row=1, column=0, pady=2)
        self.preview_cc_var = tk.StringVar()
        e_cc = ttk.Entry(env, textvariable=self.preview_cc_var, font=("Segoe UI", 9))
        e_cc.grid(row=1, column=1, sticky="ew", pady=2)
        e_cc.bind("<FocusOut>", lambda e: self._sync_envelope_to_memory())

        tk.Label(env, text="Subject:", bg="#f8fafc", fg="#64748b", font=("Segoe UI", 9, "bold"), width=8, anchor="w").grid(row=2, column=0, pady=2)
        self.preview_subj_var = tk.StringVar()
        tk.Label(env, textvariable=self.preview_subj_var, bg="#f8fafc", fg="#0f172a",
                 font=("Segoe UI", 10, "bold"), anchor="w").grid(row=2, column=1, sticky="ew", pady=2)

        env.columnconfigure(1, weight=1)

        self.preview_body_label = tk.Label(
            parent, text="", bg="#ffffff", fg="#0f172a", justify=tk.LEFT, anchor="nw",
            font=("Calibri", 11), padx=14, pady=6
        )
        self.preview_body_label.pack(fill=tk.X)

        table_frame = tk.Frame(parent, bg="#ffffff", padx=12, pady=4)
        table_frame.pack(fill=tk.BOTH, expand=True)

        self.preview_tree = ttk.Treeview(table_frame, show="headings")
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.preview_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.preview_footer_var = tk.StringVar()
        tk.Label(parent, textvariable=self.preview_footer_var, bg="#fefce8", fg="#713f12",
                 font=("Segoe UI", 8, "bold"), anchor="w", padx=12, pady=5).pack(fill=tk.X, padx=12, pady=(2, 10))

    def _init_dates(self):
        today = datetime.date.today()
        self.short_date_var.set(today.strftime("%d.%m.%y"))
        self.long_date_var.set(today.strftime("%d %B %Y").lstrip("0"))

    def manual_save_profile(self):
        self._collect_and_save()
        self.status_var.set(f"✔ All settings, routing, templates, and signature saved to {PROFILE_PATH}")

    def _collect_and_save(self):
        self.profile["font_family"] = self.font_family_var.get()
        self.profile["body_font_size"] = self.body_size_var.get()
        self.profile["table_font_size"] = self.table_size_var.get()
        self.profile["merge_visual_cells"] = self.merge_cells_var.get()
        self.profile["include_footer_notes"] = self.include_notes_var.get()
        self.profile["include_signature"] = self.include_sig_var.get()
        self.profile["subject_template"] = self.subject_var.get()
        self.profile["body_template"] = self.body_text.get("1.0", tk.END).strip()
        self.profile["signature_html"] = self.sig_text.get("1.0", tk.END).strip()
        self.profile["last_excel_path"] = self.current_excel_path
        save_profile(self.profile)

    def on_settings_changed(self):
        self._collect_and_save()
        self.update_preview()

    def on_template_changed(self):
        self._collect_and_save()
        self.update_preview()

    def on_signature_changed(self):
        self._collect_and_save()
        self.update_preview()

    def insert_body_snippet(self, snippet):
        self.body_text.insert(tk.INSERT, snippet)
        self.on_template_changed()

    def wrap_body_selection(self, open_tag, close_tag):
        try:
            start = self.body_text.index(tk.SEL_FIRST)
            end = self.body_text.index(tk.SEL_LAST)
            selected = self.body_text.get(start, end)
            self.body_text.delete(start, end)
            self.body_text.insert(start, f"{open_tag}{selected}{close_tag}")
            self.on_template_changed()
        except tk.TclError:
            self.status_var.set("Select text inside Opening Body Text first to apply formatting.")

    def pick_body_text_color(self):
        color = colorchooser.askcolor(title="Choose Text Color")[1]
        if color:
            self.wrap_body_selection(f'<span style="color:{color};">', "</span>")

    def upload_logo_image(self):
        path = filedialog.askopenfilename(
            title="Select Signature Logo Image",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.gif;*.bmp")]
        )
        if not path:
            return
        ext = os.path.splitext(path)[1].lower().lstrip(".") or "png"
        if ext == "jpg":
            ext = "jpeg"
        with open(path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("ascii")
        self.profile["custom_logo_base64"] = f"data:image/{ext};base64,{encoded}"
        self.logo_status_var.set(f"Saved: {os.path.basename(path)} ✔")
        self._collect_and_save()
        self.update_preview()

    def clear_logo_image(self):
        self.profile["custom_logo_base64"] = ""
        self.logo_status_var.set("No custom logo uploaded")
        self._collect_and_save()
        self.update_preview()

    def browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel Table File",
            filetypes=[("Excel Workbooks", "*.xlsx *.xlsm")]
        )
        if path:
            self._load_excel_file(path)

    def _load_excel_file(self, path, sheet_name=None, silent=False):
        try:
            tables, notes, sheets, active_sheet, has_merges, extracted_rt = parse_excel_workbook(path, sheet_name)
            if not tables:
                if not silent:
                    self.status_var.set("⚠ No valid RSE region rows were found in that sheet.")
                return

            self.current_excel_path = path
            self.table_data = tables
            self.footer_notes = notes
            self.sheet_names = sheets
            self.excel_label_var.set(os.path.basename(path))

            self.sheet_combo["values"] = sheets
            self.sheet_var.set(active_sheet)

            if has_merges:
                self.merge_cells_var.set(True)

            if extracted_rt:
                for reg, info in extracted_rt.items():
                    self.profile["routing_data"].setdefault(reg, {})
                    if info.get("to"):
                        self.profile["routing_data"][reg]["to"] = info["to"]
                    if info.get("cc"):
                        self.profile["routing_data"][reg]["cc"] = info["cc"]
                self._reload_routing_tree()

            self._collect_and_save()
            self.refresh_region_dropdown()
            self.update_preview()
            if not silent:
                self.status_var.set(f"✔ Loaded {len(tables)} region(s) from '{active_sheet}' ({os.path.basename(path)}).")
        except Exception as e:
            traceback.print_exc()
            if not silent:
                self.status_var.set(f"❌ Excel Error: {e}")

    def on_sheet_selected(self, event=None):
        if self.current_excel_path and os.path.exists(self.current_excel_path):
            self._load_excel_file(self.current_excel_path, sheet_name=self.sheet_var.get())

    def refresh_region_dropdown(self):
        regions = list(self.table_data.keys())
        self.region_combo["values"] = regions
        if regions and self.region_var.get() not in regions:
            self.region_var.set(regions[0])
        self.on_region_changed()

    def on_region_changed(self):
        reg = self.region_var.get()
        rt = self.profile.get("routing_data", {}).get(reg, {"to": "", "cc": ""})
        self.preview_to_var.set(rt.get("to", ""))
        self.preview_cc_var.set(rt.get("cc", ""))
        self.update_preview()

    def _sync_envelope_to_memory(self):
        reg = self.region_var.get()
        if not reg:
            return
        self.profile["routing_data"].setdefault(reg, {})
        self.profile["routing_data"][reg]["to"] = self.preview_to_var.get().strip()
        self.profile["routing_data"][reg]["cc"] = self.preview_cc_var.get().strip()
        self._reload_routing_tree()
        self._collect_and_save()
        self.status_var.set(f"✔ Updated and remembered To/Cc routing for {reg}.")

    def import_routing_excel(self):
        if not HAS_OPENPYXL:
            self.status_var.set("❌ openpyxl is required.")
            return
        path = filedialog.askopenfilename(
            title="Select Routing Excel File",
            filetypes=[("Excel Workbooks", "*.xlsx *.xlsm")]
        )
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return
            headers = [str(h or "").strip() for h in rows[0]]
            rse_i = next((i for i, h in enumerate(headers) if h.lower() in ("rse", "region")), -1)
            to_i = next((i for i, h in enumerate(headers) if h.lower() == "to"), -1)
            zsi_i = next((i for i, h in enumerate(headers) if "zsi" in h.lower()), -1)
            fix_i = next((i for i, h in enumerate(headers) if "fixed" in h.lower()), -1)
            cc_i = next((i for i, h in enumerate(headers) if h.lower() == "cc"), -1)

            if rse_i == -1:
                self.status_var.set("❌ Routing file must have an 'RSE' column.")
                return

            count = 0
            for r in rows[1:]:
                reg = str(r[rse_i] or "").strip()
                if not reg:
                    continue
                entry = self.profile["routing_data"].setdefault(reg, {})
                if to_i != -1 and r[to_i]:
                    entry["to"] = str(r[to_i]).strip()
                ccs = []
                if zsi_i != -1 and r[zsi_i]:
                    ccs.append(str(r[zsi_i]).strip())
                if fix_i != -1 and r[fix_i]:
                    ccs.append(str(r[fix_i]).strip())
                if ccs:
                    entry["cc"] = "; ".join(ccs)
                elif cc_i != -1 and r[cc_i]:
                    entry["cc"] = str(r[cc_i]).strip()
                count += 1

            self._reload_routing_tree()
            self._collect_and_save()
            self.on_region_changed()
            self.status_var.set(f"✔ Imported & saved routing memory for {count} region(s).")
        except Exception as e:
            self.status_var.set(f"❌ Routing import failed: {e}")

    def _reload_routing_tree(self):
        for item in self.routing_tree.get_children():
            self.routing_tree.delete(item)
        for reg, info in sorted(self.profile.get("routing_data", {}).items()):
            self.routing_tree.insert("", tk.END, values=(reg, info.get("to", ""), info.get("cc", "")))

    def on_routing_tree_select(self, event=None):
        sel = self.routing_tree.selection()
        if not sel:
            return
        vals = self.routing_tree.item(sel[0], "values")
        self.rt_rse_var.set(vals[0])
        self.rt_to_var.set(vals[1])
        self.rt_cc_var.set(vals[2])

    def add_or_update_routing_row(self):
        reg = self.rt_rse_var.get().strip()
        if not reg:
            return
        self.profile["routing_data"][reg] = {
            "to": self.rt_to_var.get().strip(),
            "cc": self.rt_cc_var.get().strip()
        }
        self._reload_routing_tree()
        self._collect_and_save()
        if reg == self.region_var.get():
            self.on_region_changed()
        self.status_var.set(f"✔ Saved routing for {reg} to memory.")

    def delete_routing_row(self):
        sel = self.routing_tree.selection()
        if not sel:
            return
        reg = self.routing_tree.item(sel[0], "values")[0]
        self.profile["routing_data"].pop(reg, None)
        self._reload_routing_tree()
        self._collect_and_save()
        self.status_var.set(f"Deleted routing for {reg}.")

    def update_preview(self):
        region = self.region_var.get() or "N/A"
        short_d = self.short_date_var.get()
        long_d = self.long_date_var.get()

        subj = (
            self.subject_var.get()
            .replace("{region}", region)
            .replace("{shortDate}", short_d)
            .replace("{longDate}", long_d)
        )
        self.preview_subj_var.set(subj)

        raw_body = (
            self.body_text.get("1.0", tk.END).strip()
            .replace("{region}", region)
            .replace("{shortDate}", short_d)
            .replace("{longDate}", long_d)
        )
        plain_body = re.sub(r"<[^>]+>", "", raw_body)
        font_fam = self.font_family_var.get()
        try:
            f_size = int(float(self.body_size_var.get()))
        except ValueError:
            f_size = 11
        self.preview_body_label.configure(text=plain_body, font=(font_fam, f_size))

        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)

        rows = self.table_data.get(region, [])
        if not rows:
            self.preview_tree["columns"] = ()
            self.preview_footer_var.set("No rows for selected region.")
            return

        exclude = {"to", "cc", "_isredrow", "cc: zsi", "fixed cc"}
        cols = [k for k in rows[0].keys() if k.lower().strip() not in exclude and not k.startswith("_")]
        self.preview_tree["columns"] = cols

        for c in cols:
            self.preview_tree.heading(c, text=c)
            self.preview_tree.column(c, width=150 if c.lower() == "remarks" else 85, anchor="w" if c.lower() == "remarks" else "center")

        totals = {c: 0.0 for c in cols}
        is_num = {}
        for c in cols:
            low = c.lower()
            if any(t in low for t in ("code", "ase", "rse", "mpe", "specialty", "remarks")):
                is_num[c] = False
            else:
                is_num[c] = all(r.get(c, "") in ("", None) or isinstance(r.get(c), (int, float)) for r in rows)

        for r in rows:
            vals = [r.get(c, "") for c in cols]
            self.preview_tree.insert("", tk.END, values=vals)
            for c in cols:
                if is_num[c] and r.get(c, "") not in ("", None):
                    totals[c] += float(r[c])

        tot_vals = []
        for idx, c in enumerate(cols):
            if idx == 0:
                tot_vals.append(f"{region} Total")
            elif is_num[c]:
                tot_vals.append(int(totals[c]) if totals[c].is_integer() else round(totals[c], 2))
            else:
                tot_vals.append("")
        self.preview_tree.insert("", tk.END, values=tot_vals)

        notes = self.footer_notes.get(region, [])
        note_summary = " | ".join(f"{n['label']}: {n['value']}" for n in notes) if (notes and self.include_notes_var.get()) else "No footer notes"
        sig_state = "Signature ON" if self.include_sig_var.get() else "Signature OFF"
        self.preview_footer_var.set(f"Notes: {note_summary}   •   {sig_state}   •   Font: {font_fam} ({self.table_size_var.get()}pt table)")

    def _build_draft_payload(self, region):
        self._collect_and_save()
        short_d = self.short_date_var.get()
        long_d = self.long_date_var.get()

        rt = self.profile.get("routing_data", {}).get(region, {"to": "", "cc": ""})
        to_addr = self.preview_to_var.get().strip() if region == self.region_var.get() else rt.get("to", "")
        cc_addr = self.preview_cc_var.get().strip() if region == self.region_var.get() else rt.get("cc", "")

        subject = (
            self.profile.get("subject_template", "")
            .replace("{region}", region)
            .replace("{shortDate}", short_d)
            .replace("{longDate}", long_d)
        )
        html_body = build_full_email_html(
            profile=self.profile,
            region=region,
            short_date=short_d,
            long_date=long_d,
            table_rows=self.table_data.get(region, []),
            footer_notes=self.footer_notes.get(region, []),
        )
        return to_addr, cc_addr, subject, html_body

    def open_in_outlook_single(self):
        region = self.region_var.get()
        if not region or region not in self.table_data:
            self.status_var.set("⚠ No region selected.")
            return
        if not HAS_WIN32COM:
            self.status_var.set("⚠ Outlook COM unavailable. Exporting .eml instead...")
            self.export_eml_or_zip(force_single=True)
            return

        try:
            to_addr, cc_addr, subject, html_body = self._build_draft_payload(region)
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = to_addr
            mail.CC = cc_addr
            mail.Subject = subject
            mail.HTMLBody = html_body
            mail.Display(False)
            self.status_var.set(f"✔ Opened Outlook draft for {region}!")
        except Exception as e:
            self.status_var.set(f"❌ Outlook COM Error: {e}")

    def open_in_outlook_batch(self):
        if not self.table_data:
            self.status_var.set("⚠ No region data available.")
            return
        if not HAS_WIN32COM:
            self.status_var.set("⚠ Outlook COM unavailable. Exporting Batch .zip instead...")
            self.export_eml_or_zip(force_batch=True)
            return

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            count = 0
            for region in self.table_data:
                to_addr, cc_addr, subject, html_body = self._build_draft_payload(region)
                mail = outlook.CreateItem(0)
                mail.To = to_addr
                mail.CC = cc_addr
                mail.Subject = subject
                mail.HTMLBody = html_body
                mail.Display(False)
                count += 1
            self.status_var.set(f"✔ Opened {count} Outlook draft windows!")
        except Exception as e:
            self.status_var.set(f"❌ Batch Outlook Error: {e}")

    def export_eml_or_zip(self, force_single=False, force_batch=False):
        regions = list(self.table_data.keys())
        if not regions:
            return

        if len(regions) == 1 or force_single:
            reg = self.region_var.get()
            to_addr, cc_addr, subject, html_body = self._build_draft_payload(reg)
            safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", subject)[:120] + ".eml"
            save_path = filedialog.asksaveasfilename(
                title="Save Email Draft (.eml)",
                initialfile=safe_name,
                defaultextension=".eml",
                filetypes=[("Outlook Email Draft", "*.eml")]
            )
            if not save_path:
                return
            eml = f"To: {to_addr}\r\nCc: {cc_addr}\r\nSubject: {subject}\r\nX-Unsent: 1\r\nContent-Type: text/html; charset=\"utf-8\"\r\n\r\n{html_body}"
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(eml)
            self.status_var.set(f"✔ Saved {os.path.basename(save_path)}")
        else:
            short_d = self.short_date_var.get().replace(".", "_")
            save_path = filedialog.asksaveasfilename(
                title="Save Batch Emails (.zip)",
                initialfile=f"Batch_Emails_{short_d}.zip",
                defaultextension=".zip",
                filetypes=[("ZIP Archive", "*.zip")]
            )
            if not save_path:
                return
            with zipfile.ZipFile(save_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for reg in regions:
                    to_addr, cc_addr, subject, html_body = self._build_draft_payload(reg)
                    safe_name = f"{re.sub(r'[^A-Za-z0-9_-]+', '_', subject)[:110]}_{reg}.eml"
                    eml = f"To: {to_addr}\r\nCc: {cc_addr}\r\nSubject: {subject}\r\nX-Unsent: 1\r\nContent-Type: text/html; charset=\"utf-8\"\r\n\r\n{html_body}"
                    zf.writestr(safe_name, eml)
            self.status_var.set(f"✔ Exported {len(regions)} emails to {os.path.basename(save_path)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = OutlookAutomatorApp(root)
    root.mainloop()
