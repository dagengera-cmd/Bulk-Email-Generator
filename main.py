# ... existing code ...
import base64
import io
import json
import os
import re
import tempfile
import threading
import xml.etree.ElementTree as ET
from pathlib import Path
import webview

try:
    import openpyxl
    from openpyxl.styles.colors import COLOR_INDEX
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    COLOR_INDEX = ()

try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

PROFILE_DIR = Path.home() / ".outlook_email_automator"
PROFILE_FILE = PROFILE_DIR / "profile.json"
COM_LOCK = threading.Lock()


# ==============================================================================
# NATIVE EXCEL DISPLAYFORMAT ENGINE (EXACT ON-SCREEN COLORS & CONDITIONAL FORMATS)
# ==============================================================================
def _bgr_int_to_hex(color_val):
    """Convert Excel COM BGR color integer (e.g. 0x00784E1F) to #RRGGBB."""
    if color_val is None:
        return None
    try:
        ci = int(color_val)
        if ci < 0:
            return None
        r = ci & 0xFF
        g = (ci >> 8) & 0xFF
        b = (ci >> 16) & 0xFF
        return f"#{r:02X}{g:02X}{b:02X}"
    except Exception:
        return None


def _format_cell_val(raw_val):
    if raw_val is None:
        return ""
    if isinstance(raw_val, float) and raw_val.is_integer():
        return str(int(raw_val))
    return str(raw_val).strip()


def parse_excel_com_displayformat(file_path, requested_sheet=None):
    """
    Opens the Excel file silently in an isolated Excel COM instance (DispatchEx)
    and reads cell.DisplayFormat for the exact on-screen fill, font color, bold,
    italic, alignment, and 2D merges — automatically capturing ALL Conditional
    Formatting formulas, Office themes, and Excel Table styles.
    """
    if not HAS_WIN32COM:
        return None

    with COM_LOCK:
        pythoncom.CoInitialize()
        excel = None
        wb = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            try:
                excel.EnableEvents = False
            except Exception:
                pass

            wb = excel.Workbooks.Open(
                os.path.abspath(file_path),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
            )
            sheet_names = [s.Name for s in wb.Worksheets]
            if not sheet_names:
                return {"ok": False, "error": "Workbook contains no sheets."}

            active_sheet = requested_sheet if (requested_sheet in sheet_names) else sheet_names[0]
            ws = wb.Worksheets(active_sheet)
            used = ws.UsedRange
            if used is None:
                return {"ok": False, "error": f"Worksheet '{active_sheet}' is empty."}

            start_r = used.Row
            start_c = used.Column
            raw_matrix = used.Value
            if raw_matrix is None:
                return {"ok": False, "error": f"Worksheet '{active_sheet}' is empty."}
            if not isinstance(raw_matrix, tuple):
                raw_matrix = ((raw_matrix,),)
            elif raw_matrix and not isinstance(raw_matrix[0], tuple):
                raw_matrix = (raw_matrix,)

            # Trim trailing blank rows to avoid scanning phantom formatted rows
            raw_rows = list(raw_matrix)
            while raw_rows and all(v is None or str(v).strip() == "" for v in raw_rows[-1]):
                raw_rows.pop()
            if not raw_rows:
                return {"ok": False, "error": f"Worksheet '{active_sheet}' is empty."}

            # Fallback to fast openpyxl for very large sheets (>600 rows) to keep UI snappy
            if len(raw_rows) > 600:
                return None

            grid = [[_format_cell_val(v) for v in r] for r in raw_rows]
            num_rows = len(grid)
            num_cols = len(grid[0])

            # Locate header row containing RSE or Region
            header_row_idx = -1
            header_cols = []
            for r in range(num_rows):
                rse_idx = next((i for i, v in enumerate(grid[r]) if re.match(r"^(rse|region)$", v, re.I)), -1)
                if rse_idx != -1 and sum(1 for v in grid[r] if v != "") >= 2:
                    header_row_idx = r
                    header_cols = [{"name": v, "idx": idx} for idx, v in enumerate(grid[r]) if v != ""]
                    break

            if header_row_idx == -1:
                return {"ok": False, "error": "Could not find an 'RSE' or 'Region' header column in this sheet."}

            # Read DisplayFormat & MergeArea ONLY for active header + data columns
            active_col_indices = sorted({h["idx"] for h in header_cols})
            style_grid = [[None] * num_cols for _ in range(num_rows)]
            merge_master = {}
            merge_covered = set()
            seen_merge_addrs = set()

            align_map = {-4131: "left", -4108: "center", -4152: "right"}

            for r in range(header_row_idx, num_rows):
                # Skip completely empty rows
                if all(grid[r][c] == "" for c in range(num_cols)):
                    continue
                excel_r = start_r + r
                for c in active_col_indices:
                    excel_c = start_c + c
                    cell = ws.Cells(excel_r, excel_c)

                    # Detect 2D merges directly from Excel
                    if cell.MergeCells:
                        ma = cell.MergeArea
                        addr = ma.Address
                        if addr not in seen_merge_addrs:
                            seen_merge_addrs.add(addr)
                            top_r = ma.Row - start_r
                            left_c = ma.Column - start_c
                            rs = ma.Rows.Count
                            cs = ma.Columns.Count
                            if 0 <= top_r < num_rows and 0 <= left_c < num_cols:
                                top_val = grid[top_r][left_c]
                                if rs > 1 or cs > 1:
                                    merge_master[f"{top_r}:{left_c}"] = {"rs": rs, "cs": cs}
                                for rr in range(top_r, min(top_r + rs, num_rows)):
                                    for cc in range(left_c, min(left_c + cs, num_cols)):
                                        if not (rr == top_r and cc == left_c) and (rs > 1 or cs > 1):
                                            merge_covered.add(f"{rr}:{cc}")
                                        if grid[rr][cc] == "" and top_val != "":
                                            grid[rr][cc] = top_val

                    # Read exact rendered DisplayFormat (includes Conditional Formatting & Themes)
                    df = cell.DisplayFormat
                    st = {}
                    interior = df.Interior
                    if interior.ColorIndex != -4142:  # Not xlColorIndexNone
                        bg_hex = _bgr_int_to_hex(interior.Color)
                        if bg_hex:
                            st["bg"] = bg_hex

                    font = df.Font
                    fg_hex = _bgr_int_to_hex(font.Color)
                    if fg_hex:
                        st["color"] = fg_hex
                    if bool(font.Bold):
                        st["bold"] = True
                    if bool(font.Italic):
                        st["italic"] = True

                    h_align = align_map.get(df.HorizontalAlignment)
                    if h_align:
                        st["align"] = h_align

                    if st:
                        style_grid[r][c] = st

            return _build_regions_from_grid(
                grid, style_grid, merge_master, merge_covered,
                header_row_idx, header_cols, sheet_names, active_sheet,
                engine_used="Excel DisplayFormat"
            )
        except Exception:
            return None
        finally:
            try:
                if wb is not None:
                    wb.Close(False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()


# ==============================================================================
# EXCEL COLOR, THEME & CONDITIONAL FORMATTING ENGINE (OPENPYXL FALLBACK)
# ==============================================================================
# ... existing code ...
def parse_excel_bytes_openpyxl(raw_bytes, requested_sheet=None):
    """Parse an Excel workbook from bytes using openpyxl with trimmed bounds & style fidelity."""
    if not HAS_OPENPYXL:
        return {"ok": False, "error": "openpyxl is not installed."}

    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
    sheet_names = list(wb.sheetnames)
    if not sheet_names:
        return {"ok": False, "error": "Workbook contains no sheets."}

    active_sheet = requested_sheet if (requested_sheet in sheet_names) else sheet_names[0]
    ws = wb[active_sheet]
    theme_palette = extract_theme_palette(wb)
    cf_entries = build_cf_lookup(ws, theme_palette)

    max_r = min(ws.max_row or 0, 2500)
    max_c = min(ws.max_column or 0, 35)
    if max_r == 0 or max_c == 0:
        return {"ok": False, "error": f"Worksheet '{active_sheet}' is empty."}

    # Fast pass: read values first and trim trailing empty rows/columns
    raw_grid = []
    grid = []
    last_non_empty_r = -1
    last_non_empty_c = -1
    for r_idx, row_tuple in enumerate(ws.iter_rows(min_row=1, max_row=max_r, min_col=1, max_col=max_c, values_only=True)):
        row_vals = [_format_cell_val(v) for v in row_tuple]
        raw_grid.append(row_tuple)
        grid.append(row_vals)
        for c_idx, sv in enumerate(row_vals):
            if sv != "":
                last_non_empty_r = r_idx
                if c_idx > last_non_empty_c:
                    last_non_empty_c = c_idx

    if last_non_empty_r < 0 or last_non_empty_c < 0:
        return {"ok": False, "error": f"Worksheet '{active_sheet}' is empty."}

    num_r = last_non_empty_r + 1
    num_c = last_non_empty_c + 1
    grid = [row[:num_c] for row in grid[:num_r]]
    raw_grid = [row[:num_c] for row in raw_grid[:num_r]]

    # Extract styles only for the trimmed active area
    style_grid = []
    for r in range(num_r):
        row_styles = []
        row_has_data = any(grid[r])
        for c in range(num_c):
            if not row_has_data:
                row_styles.append(None)
                continue
            cell = ws.cell(row=r + 1, column=c + 1)
            row_styles.append(
                extract_cell_style_openpyxl(cell, r + 1, c + 1, raw_grid[r][c], grid[r][c], theme_palette, cf_entries)
            )
        style_grid.append(row_styles)

    # Track 2D merged cells (rowspan + colspan) and unroll values/styles
    merge_master = {}
    merge_covered = set()
    for rng in ws.merged_cells.ranges:
        top_r = rng.min_row - 1
        bot_r = min(rng.max_row - 1, num_r - 1)
        left_c = rng.min_col - 1
        right_c = min(rng.max_col - 1, num_c - 1)
        if top_r < 0 or left_c < 0 or top_r >= num_r or left_c >= num_c:
            continue
        top_val = grid[top_r][left_c]
        top_style = style_grid[top_r][left_c]
        rs = (bot_r - top_r) + 1
        cs = (right_c - left_c) + 1
        if rs > 1 or cs > 1:
            merge_master[f"{top_r}:{left_c}"] = {"rs": rs, "cs": cs}
        for rr in range(top_r, bot_r + 1):
            for cc in range(left_c, right_c + 1):
                if not (rr == top_r and cc == left_c) and (rs > 1 or cs > 1):
                    merge_covered.add(f"{rr}:{cc}")
                if grid[rr][cc] == "" and top_val != "":
                    grid[rr][cc] = top_val
                if style_grid[rr][cc] is None and top_style is not None:
                    style_grid[rr][cc] = dict(top_style)

    # Locate primary header row containing RSE or Region
    header_row_idx = -1
    header_cols = []
    for r in range(len(grid)):
        rse_idx = next((i for i, v in enumerate(grid[r]) if re.match(r"^(rse|region)$", v, re.I)), -1)
        if rse_idx != -1 and sum(1 for v in grid[r] if v != "") >= 2:
            header_row_idx = r
            header_cols = [{"name": v, "idx": idx} for idx, v in enumerate(grid[r]) if v != ""]
            break

    if header_row_idx == -1:
        return {"ok": False, "error": "Could not find an 'RSE' or 'Region' header column in this sheet."}

    return _build_regions_from_grid(
        grid, style_grid, merge_master, merge_covered,
        header_row_idx, header_cols, sheet_names, active_sheet,
        engine_used="openpyxl"
    )


def _build_regions_from_grid(
    grid, style_grid, merge_master, merge_covered,
    header_row_idx, header_cols, sheet_names, active_sheet,
    engine_used="openpyxl"
):
    """Shared region/multi-block grouping logic used by both COM DisplayFormat and openpyxl."""
    table_data = {}
    region_notes = {}
    region_headers = {}
    region_header_styles = {}
    region_explicit_merges = {}
    extracted_routing = {}

    current_region = ""
    last_group_vals = {}
    group_col_names = {"rse", "region", "ase", "specialty", "speciality", "zone", "territory"}
    routing_col_names = {"to", "cc", "cc: zsi", "fixed cc"}

    for r in range(header_row_idx + 1, len(grid)):
        row = grid[r]
        non_empty = [v for v in row if v != ""]
        if not non_empty:
            continue

        # Detect repeated block header rows in multi-block sheets
        matches_header = sum(
            1 for h in header_cols if h["idx"] < len(row) and row[h["idx"]] and row[h["idx"]].lower() == h["name"].lower()
        )
        if matches_header >= 2:
            header_cols = [{"name": v, "idx": idx} for idx, v in enumerate(row) if v != ""]
            last_group_vals = {}
            continue

        # Skip built-in Excel Total rows
        first_two = " ".join(non_empty[:2]).lower()
        if re.search(r"\btotal\b", first_two) and not any(re.match(r"^[A-Z]\d{3}$", v, re.I) for v in non_empty):
            continue

        # Capture footer notes (e.g., "Did not Select any RS PB")
        note_idx = next((i for i, v in enumerate(row) if re.search(r"did not select|not selected|note\s*:", v, re.I)), -1)
        if note_idx != -1:
            label = row[note_idx]
            detail = ", ".join(v for v in row[note_idx + 1:] if v != "")
            if current_region:
                region_notes.setdefault(current_region, []).append({"label": label, "detail": detail})
            continue

        row_obj = {}
        row_styles = {}
        row_merge_info = {}
        for h in header_cols:
            name = h["name"]
            idx = h["idx"]
            row_obj[name] = row[idx] if idx < len(row) else ""
            if idx < len(style_grid[r]) and style_grid[r][idx]:
                row_styles[name] = style_grid[r][idx]
            key = f"{r}:{idx}"
            if key in merge_master:
                row_merge_info[name] = merge_master[key]
            elif key in merge_covered:
                row_merge_info[name] = {"covered": True}

        rse_col = next((h for h in header_cols if re.match(r"^(rse|region)$", h["name"], re.I)), None)
        reg_val = row_obj.get(rse_col["name"], "") if rse_col else ""
        if not reg_val:
            code_col = next((h for h in header_cols if re.match(r"^(ase|mpe|t\.?code|territory)$", h["name"], re.I)), None)
            code_val = row_obj.get(code_col["name"], "") if code_col else ""
            m = re.match(r"^([A-Za-z]+\d)", str(code_val))
            if m:
                reg_val = m.group(1).upper()
            elif current_region:
                reg_val = current_region

        if not reg_val:
            continue

        if reg_val != current_region:
            current_region = reg_val
            last_group_vals = {}
        if rse_col:
            row_obj[rse_col["name"]] = current_region

        for h in header_cols:
            low = h["name"].lower()
            if low in group_col_names:
                if row_obj[h["name"]] != "":
                    last_group_vals[low] = row_obj[h["name"]]
                elif low in last_group_vals:
                    row_obj[h["name"]] = last_group_vals[low]

        if current_region not in table_data:
            table_data[current_region] = []
            region_headers[current_region] = [
                h["name"] for h in header_cols if h["name"].lower() not in routing_col_names
            ]
            region_header_styles[current_region] = {}
            region_explicit_merges[current_region] = []
            for h in header_cols:
                hs = style_grid[header_row_idx][h["idx"]] if h["idx"] < len(style_grid[header_row_idx]) else None
                if hs:
                    hs_copy = dict(hs)
                    if hs_copy.get("bg") and not hs_copy.get("color") and is_dark_hex(hs_copy["bg"]):
                        hs_copy["color"] = "#FFFFFF"
                    region_header_styles[current_region][h["name"]] = hs_copy

        to_val = row_obj.get("To") or row_obj.get("TO")
        if to_val:
            extracted_routing.setdefault(current_region, {})["to"] = to_val
        cc_parts = []
        if row_obj.get("CC: ZSI"):
            cc_parts.append(row_obj["CC: ZSI"])
        if row_obj.get("Fixed CC"):
            cc_parts.append(row_obj["Fixed CC"])
        if cc_parts:
            extracted_routing.setdefault(current_region, {})["cc"] = "; ".join(cc_parts)
        elif row_obj.get("CC") or row_obj.get("Cc"):
            extracted_routing.setdefault(current_region, {})["cc"] = row_obj.get("CC") or row_obj.get("Cc")

        rem_low = str(row_obj.get("Remarks", "")).lower()
        row_obj["isRedRow"] = ("0 steps" in rem_low) or ("immediate attention" in rem_low)
        row_obj["_styles"] = row_styles

        table_data[current_region].append(row_obj)
        region_explicit_merges[current_region].append(row_merge_info)

    if not table_data:
        return {"ok": False, "error": "No valid region rows found in this sheet."}

    return {
        "ok": True,
        "engineUsed": engine_used,
        "sheetNames": sheet_names,
        "activeSheet": active_sheet,
        "tableData": table_data,
        "regionNotes": region_notes,
        "regionHeaders": region_headers,
        "regionHeaderStyles": region_header_styles,
        "regionExplicitMerges": region_explicit_merges,
        "extractedRouting": extracted_routing,
    }


def parse_excel_file_smart(file_path, requested_sheet=None):
    """Try Microsoft Excel native DisplayFormat first; fall back to openpyxl seamlessly."""
    res = parse_excel_com_displayformat(file_path, requested_sheet)
    if res is not None:
        return res
    raw_bytes = Path(file_path).read_bytes()
    return parse_excel_bytes_openpyxl(raw_bytes, requested_sheet)


def parse_routing_excel_file(file_path):
    """Parse To/CC routing Excel file directly in Python via openpyxl."""
    if not HAS_OPENPYXL:
        return {"ok": False, "error": "openpyxl is not installed."}
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {"ok": False, "error": "Routing file is empty."}
    headers = [str(h or "").strip() for h in rows[0]]
    rse_i = next((i for i, h in enumerate(headers) if h.lower() in ("rse", "region")), -1)
    to_i = next((i for i, h in enumerate(headers) if h.lower() == "to"), -1)
    zsi_i = next((i for i, h in enumerate(headers) if "zsi" in h.lower()), -1)
    fix_i = next((i for i, h in enumerate(headers) if "fixed" in h.lower()), -1)
    cc_i = next((i for i, h in enumerate(headers) if h.lower() == "cc"), -1)
    if rse_i == -1:
        return {"ok": False, "error": "Routing Excel must contain an 'RSE' or 'Region' column."}

    parsed = {}
    for row in rows[1:]:
        reg = _format_cell_val(row[rse_i] if rse_i < len(row) else "")
        if not reg:
            continue
        entry = {}
        if to_i != -1 and to_i < len(row) and row[to_i]:
            entry["to"] = str(row[to_i]).strip()
        ccs = []
        if zsi_i != -1 and zsi_i < len(row) and row[zsi_i]:
            ccs.append(str(row[zsi_i]).strip())
        if fix_i != -1 and fix_i < len(row) and row[fix_i]:
            ccs.append(str(row[fix_i]).strip())
        if ccs:
            entry["cc"] = "; ".join(ccs)
        elif cc_i != -1 and cc_i < len(row) and row[cc_i]:
            entry["cc"] = str(row[cc_i]).strip()
        parsed[reg] = entry
    return {"ok": True, "routing": parsed, "count": len(parsed)}


# ==============================================================================
# EMBEDDED MODERN HTML / TAILWIND DESKTOP UI
# ==============================================================================
HTML_UI = r"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Outlook Email Automator - Desktop Edition</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js" crossorigin="anonymous"></script>
    <style>
# ... existing code ...
                    <div id="tableUploadBubble" class="p-3 bg-indigo-50/60 border border-indigo-100 rounded-lg space-y-2">
                        <div class="flex justify-between items-center">
                            <span class="text-[11px] font-bold text-indigo-950">Select Table Excel (.xlsx)</span>
                            <span class="text-[10px] text-indigo-600 font-medium">Native Excel DisplayFormat engine</span>
                        </div>
                        <div class="flex items-center gap-2">
                            <button type="button" id="btnPickExcel" onclick="pickExcelFile()" class="px-3 py-1.5 rounded-full text-[11px] font-semibold bg-indigo-600 text-white hover:bg-indigo-700 transition cursor-pointer shadow-sm">
                                Choose Excel File
                            </button>
                            <span id="excelFileLabel" class="text-[11px] text-slate-600 truncate">Using sample data</span>
                        </div>
                        <div id="sheetSelectorRow" class="hidden pt-1 flex items-center justify-between gap-2">
                            <label for="sheetSelect" class="text-[10px] font-semibold text-indigo-900 uppercase">Active Sheet:</label>
                            <select id="sheetSelect" onchange="handleSheetChange()" class="border border-indigo-200 rounded px-2 py-0.5 text-xs bg-white font-semibold text-slate-800"></select>
                        </div>
                    </div>

                    <div id="routingUploadBubble" class="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2 hidden">
                        <div class="flex justify-between items-center">
                            <span class="text-[11px] font-bold text-slate-700">Select Routing (.xlsx)</span>
                            <button type="button" onclick="downloadDemoRouting()" class="text-[10px] bg-white border border-slate-200 px-2 py-0.5 rounded-full text-indigo-600 hover:text-indigo-800 font-semibold">
                                Get Demo File
                            </button>
                        </div>
                        <div class="flex items-center gap-2">
                            <button type="button" onclick="pickRoutingFile()" class="px-3 py-1.5 rounded-full text-[11px] font-semibold bg-slate-700 text-white hover:bg-slate-800 transition cursor-pointer shadow-sm">
                                Choose Routing File
                            </button>
                            <span class="text-[10px] text-slate-500">Auto-remembered on Windows</span>
                        </div>
                    </div>
# ... existing code ...
                            <div>
                                <label for="shortDate" class="block text-[10px] font-semibold text-slate-500 uppercase mb-0.5">Short Date</label>
                                <input type="text" id="shortDate" class="w-full border border-slate-300 rounded-md px-2 py-1.5 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none" oninput="scheduleTextRender()">
                            </div>
                            <div>
                                <label for="longDate" class="block text-[10px] font-semibold text-slate-500 uppercase mb-0.5">Long Date</label>
                                <input type="text" id="longDate" class="w-full border border-slate-300 rounded-md px-2 py-1.5 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none" oninput="scheduleTextRender()">
                            </div>
# ... existing code ...
                    <div>
                        <label class="block text-[10px] font-semibold text-slate-500 uppercase mb-0.5">Subject Line</label>
                        <input type="text" id="subjectTemplate" onfocus="lastFocusedTarget = this" ondragover="allowDrop(event)" ondrop="handleDrop(event)" class="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none" value="RE: Shelving Status of Ventrax HFA Inhaler {shortDate} of {region}" oninput="scheduleTextRender(true)">
                    </div>
# ... existing code ...
                        <div id="bodyEditor" contenteditable="true" onfocus="lastFocusedTarget = this" ondragover="allowDrop(event)" ondrop="handleDrop(event)" oninput="scheduleTextRender(true)" class="w-full min-h-[85px] border border-slate-300 rounded-b-lg px-2.5 py-2 text-xs text-slate-800 bg-white focus:ring-2 focus:ring-indigo-500 focus:outline-none leading-relaxed">Dear Vai,<br><br>Please find below the MPE-wise shelving status of Ventrax HFA Inhaler (VHI) as of {longDate}</div>
                    </div>
                </div>
# ... existing code ...
                        <div id="sigPasteBox" contenteditable="true" oninput="scheduleSigRender(true)" class="w-full min-h-[90px] max-h-48 overflow-y-auto border border-slate-300 rounded-b-lg p-2.5 bg-slate-50/50 focus:bg-white focus:ring-2 focus:ring-indigo-500 focus:outline-none text-[11px] leading-relaxed">
# ... existing code ...
    <script>
        let lastFocusedTarget = null;
        let customSigLogoData = null;
        let regionNotes = {};
        let regionHeaders = {};
        let regionHeaderStyles = {};
        let regionExplicitMerges = {};
        let saveTimer = null;
        let textRenderTimer = null;
        let sigRenderTimer = null;
        let lastSavedJson = "";
        let isBusy = false;

        const DEFAULT_LOGO_B64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMgAAAAyCAYAAAC0Z5aCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAH1SURBVGhD7djbbsIwEAXQvjfpv//puzgldk2J7e7a8gVzTCCK1x7aFv8mXGZmZiY9k05E7X/v/z9cE0wKk70qjQWwZ2YvG1PZqj5r/T53m0r/jJtL7Lq9s6lX0zZtvP1W+1c3/GjX1cZ36Z6c3H3fHq3P+/4rNq2T3m9u78+J11bT80Jq207f2+zU/tA2vN4n3eZ/y0qVwG95O/b2m1695f3/v7uL2eF1+5/hV0451Xo/rZ83t/zL4u3P92f1F8F8M/N3p8r/0M1u34eX/3/Gf1Vj9+P19vN6tq6n7X/1u/v1vP/+T6Hn5t7+Xv72Lq3H2d+bTj8/7l0L/a9t/c+f31P6X//dvdv3f8n1f2n73e/gV2xP8Gk8vj4AAAAASUVORK5CYII=";
# ... existing code ...
        function scheduleSaveMemory() {
            clearTimeout(saveTimer);
            saveTimer = setTimeout(saveMemoryToDisk, 800);
        }

        async function saveMemoryToDisk() {
            const payload = {
                routingData,
                subjectTemplate: document.getElementById('subjectTemplate').value,
                bodyHtml: document.getElementById('bodyEditor').innerHTML,
                sigHtml: document.getElementById('sigPasteBox').innerHTML,
                includeSignature: document.getElementById('includeSignature').checked,
                autoDetectLogo: document.getElementById('autoDetectLogo').checked,
                customSigLogoData,
                globalFontFamily: document.getElementById('globalFontFamily').value,
                bodyFontSize: document.getElementById('bodyFontSize').value,
                tableFontSize: document.getElementById('tableFontSize').value,
                mergeGroupCells: document.getElementById('mergeGroupCells').checked,
                showTotalRow: document.getElementById('showTotalRow').checked,
                includeNoteRows: document.getElementById('includeNoteRows').checked,
                retainExcelDesign: document.getElementById('retainExcelDesign').checked
            };
            const serialized = JSON.stringify(payload);
            if (serialized === lastSavedJson) return;
            lastSavedJson = serialized;
            try {
                if (window.pywebview && window.pywebview.api) {
                    await window.pywebview.api.save_profile(serialized);
                } else {
                    localStorage.setItem('outlook_automator_profile', serialized);
                }
            } catch (e) { console.error("Memory save error:", e); }
        }
# ... existing code ...
        function onSettingChange() {
            renderEmail();
            scheduleSaveMemory();
        }

        function scheduleTextRender(shouldSave = false) {
            clearTimeout(textRenderTimer);
            textRenderTimer = setTimeout(() => {
                updateSubjectAndBodyOnly();
                if (shouldSave) scheduleSaveMemory();
            }, 100);
        }

        function scheduleSigRender(shouldSave = false) {
            clearTimeout(sigRenderTimer);
            sigRenderTimer = setTimeout(() => {
                const fontFamily = document.getElementById('globalFontFamily')?.value || "Calibri, Arial, sans-serif";
                const bodyFontSize = document.getElementById('bodyFontSize')?.value || "11pt";
                document.getElementById('signatureContainer').innerHTML = generateSignatureHTML(fontFamily, bodyFontSize);
                if (shouldSave) scheduleSaveMemory();
            }, 120);
        }
# ... existing code ...
        // --- EXCEL UPLOAD VIA DIRECT NATIVE DISK PICKER (ZERO BASE64 LAG) ---
        async function applyParsedWorkbookResult(res) {
            if (!res) return;
            if (res.cancelled) return;
            if (!res.ok) {
                showStatus(res.error || "Failed to parse Excel file.", true);
                return;
            }
            if (res.fileName) {
                document.getElementById('excelFileLabel').innerText = res.fileName;
            }
            const sheetSelect = document.getElementById('sheetSelect');
            sheetSelect.innerHTML = '';
            (res.sheetNames || []).forEach(name => {
                const opt = document.createElement('option');
                opt.value = name;
                opt.text = name;
                if (name === res.activeSheet) opt.selected = true;
                sheetSelect.appendChild(opt);
            });
            document.getElementById('sheetSelectorRow').classList.toggle('hidden', (res.sheetNames || []).length <= 1);

            tableData = res.tableData || {};
            regionNotes = res.regionNotes || {};
            regionHeaders = res.regionHeaders || {};
            regionHeaderStyles = res.regionHeaderStyles || {};
            regionExplicitMerges = res.regionExplicitMerges || {};

            if (res.extractedRouting) {
                for (const reg in res.extractedRouting) {
                    if (!routingData[reg]) routingData[reg] = {};
                    if (res.extractedRouting[reg].to) routingData[reg].to = res.extractedRouting[reg].to;
                    if (res.extractedRouting[reg].cc) routingData[reg].cc = res.extractedRouting[reg].cc;
                }
            }

            const regions = Object.keys(tableData);
            const select = document.getElementById('regionSelect');
            select.innerHTML = '';
            regions.forEach(region => {
                const option = document.createElement('option');
                option.value = region;
                option.text = region;
                select.appendChild(option);
            });

            if (regions.length > 0) {
                renderEmail();
                scheduleSaveMemory();
                document.getElementById('tableBubbleText').innerText = `Table (${regions.length} RSE)`;
                const eng = res.engineUsed ? ` via ${res.engineUsed}` : "";
                showStatus(`Loaded '${res.activeSheet}' (${regions.length} regions)${eng}!`);
            } else {
                showStatus("No valid region rows found in this sheet.", true);
            }
        }

        async function pickExcelFile() {
            if (isBusy) return;
            if (!(window.pywebview && window.pywebview.api && window.pywebview.api.pick_and_parse_excel)) {
                showStatus("Desktop file picker requires the Windows .exe app.", true);
                return;
            }
            isBusy = true;
            const btn = document.getElementById('btnPickExcel');
            btn.innerText = "Loading...";
            try {
                const res = await window.pywebview.api.pick_and_parse_excel();
                await applyParsedWorkbookResult(res);
            } catch (err) {
                showStatus("Error reading Excel file: " + err, true);
            } finally {
                btn.innerText = "Choose Excel File";
                isBusy = false;
            }
        }

        async function handleSheetChange() {
            if (isBusy) return;
            const sheetName = document.getElementById('sheetSelect').value;
            if (!(window.pywebview && window.pywebview.api && window.pywebview.api.parse_current_sheet)) return;
            isBusy = true;
            showStatus(`Loading sheet '${sheetName}'...`);
            try {
                const res = await window.pywebview.api.parse_current_sheet(sheetName);
                await applyParsedWorkbookResult(res);
            } finally {
                isBusy = false;
            }
        }

        async function pickRoutingFile() {
            if (isBusy) return;
            if (!(window.pywebview && window.pywebview.api && window.pywebview.api.pick_and_parse_routing)) return;
            isBusy = true;
            try {
                const res = await window.pywebview.api.pick_and_parse_routing();
                if (!res || res.cancelled) return;
                if (!res.ok) { showStatus(res.error || "Routing import failed.", true); return; }
                for (const reg in res.routing) {
                    if (!routingData[reg]) routingData[reg] = {};
                    if (res.routing[reg].to) routingData[reg].to = res.routing[reg].to;
                    if (res.routing[reg].cc) routingData[reg].cc = res.routing[reg].cc;
                }
                renderEmail();
                saveMemoryToDisk();
                if (res.count > 0) {
                    document.getElementById('routingBubbleText').innerText = `Routing (${res.count})`;
                    document.getElementById('routingUploadBubble').classList.add('hidden');
                    showStatus(`Saved routing (To/CC) for ${res.count} regions to Windows memory!`);
                }
            } finally {
                isBusy = false;
            }
        }

        async function downloadDemoRouting() {
            if (window.pywebview && window.pywebview.api && window.pywebview.api.save_demo_routing) {
                const res = await window.pywebview.api.save_demo_routing();
                if (res) showStatus("Saved Demo_Routing_Template.xlsx!");
            }
        }
# ... existing code ...
        function updateSubjectAndBodyOnly() {
            const region = document.getElementById('regionSelect').value || "N/A";
            const dateShort = document.getElementById('shortDate').value || "";
            const dateLong = document.getElementById('longDate').value || "";
            const fontFamily = document.getElementById('globalFontFamily')?.value || "Calibri, Arial, sans-serif";
            const bodyFontSize = document.getElementById('bodyFontSize')?.value || "11pt";

            const subjectTemplate = document.getElementById('subjectTemplate').value || "";
            const bodyTemplateHtml = document.getElementById('bodyEditor').innerHTML || "";

            const finalSubject = subjectTemplate
                .replace(/\{region\}/gi, region)
                .replace(/\{shortDate\}/gi, dateShort)
                .replace(/\{longDate\}/gi, dateLong);

            const finalBody = bodyTemplateHtml
                .replace(/\{region\}/gi, region)
                .replace(/\{shortDate\}/gi, dateShort)
                .replace(/\{longDate\}/gi, dateLong);

            document.getElementById('emailSubject').innerText = finalSubject;

            const emailBodyEl = document.getElementById('emailBodyContent');
            emailBodyEl.style.fontFamily = fontFamily;
            emailBodyEl.style.fontSize = bodyFontSize;

            const dynamicBodyEl = document.getElementById('dynamicBodyText');
            dynamicBodyEl.style.fontFamily = fontFamily;
            dynamicBodyEl.style.fontSize = bodyFontSize;
            dynamicBodyEl.innerHTML = finalBody + '<br><br>';
        }

        function renderEmail() {
            const region = document.getElementById('regionSelect').value || "N/A";
            const fontFamily = document.getElementById('globalFontFamily')?.value || "Calibri, Arial, sans-serif";
            const bodyFontSize = document.getElementById('bodyFontSize')?.value || "11pt";
            const tableFontSize = document.getElementById('tableFontSize')?.value || "10.5pt";

            updateSubjectAndBodyOnly();

            const routing = routingData[region] || { to: "", cc: "" };
            const toEl = document.getElementById('emailTo');
            const ccEl = document.getElementById('emailCc');
            if (document.activeElement !== toEl) toEl.value = routing.to || "";
            if (document.activeElement !== ccEl) ccEl.value = routing.cc || "";

            document.getElementById('tableContainer').innerHTML = generateTableHTML(region, fontFamily, tableFontSize);
            document.getElementById('signatureContainer').innerHTML = generateSignatureHTML(fontFamily, bodyFontSize);
        }
# ... existing code ...
class DesktopApi:
    def __init__(self):
        self.window = None
        self.current_excel_path = None

    def pick_and_parse_excel(self):
        if not self.window:
            return {"ok": False, "error": "Window not ready."}
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Excel Files (*.xlsx;*.xlsm;*.xls)", "All files (*.*)")
        )
        if not result:
            return {"cancelled": True}
        file_path = result[0] if isinstance(result, (list, tuple)) else result
        self.current_excel_path = file_path
        try:
            parsed = parse_excel_file_smart(file_path, None)
            if isinstance(parsed, dict):
                parsed["fileName"] = os.path.basename(file_path)
            return parsed
        except Exception as e:
            return {"ok": False, "error": f"Excel parse error: {e}"}

    def parse_current_sheet(self, sheet_name=""):
        if not self.current_excel_path or not os.path.exists(self.current_excel_path):
            return {"ok": False, "error": "Please select an Excel file first."}
        try:
            parsed = parse_excel_file_smart(self.current_excel_path, sheet_name or None)
            if isinstance(parsed, dict):
                parsed["fileName"] = os.path.basename(self.current_excel_path)
            return parsed
        except Exception as e:
            return {"ok": False, "error": f"Excel parse error: {e}"}

    def pick_and_parse_routing(self):
        if not self.window:
            return {"ok": False, "error": "Window not ready."}
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Excel Files (*.xlsx;*.xlsm)", "All files (*.*)")
        )
        if not result:
            return {"cancelled": True}
        file_path = result[0] if isinstance(result, (list, tuple)) else result
        try:
            return parse_routing_excel_file(file_path)
        except Exception as e:
            return {"ok": False, "error": f"Routing parse error: {e}"}

    def save_demo_routing(self):
        if not self.window or not HAS_OPENPYXL:
            return False
        result = self.window.create_file_dialog(
            webview.SAVE_DIALOG, save_filename="Demo_Routing_Template.xlsx"
        )
        if not result:
            return False
        target = result[0] if isinstance(result, (list, tuple)) else result
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "RoutingTemplate"
        ws.append(["RSE", "To", "CC: ZSI", "Fixed CC"])
        ws.append(["R4", "PSL-(NAL)-Md. Nur-A- Alam", "SBM-(ALM)-Mahabub Alam.", "PSL-(KNH)-Kazi Nurul Hoque; PSL-(MDR)-Masud Rana khan"])
        ws.append(["A2", "PSL-(MZO)-Md. Zakir Hossain", "SBM-(ALM)-Mahabub Alam.", "PSL-(KNH)-Kazi Nurul Hoque; PSL-(MDR)-Masud Rana khan"])
        ws.append(["N4", "PSL-(KMO)-Kazi Md. Mostafa", "SBM-(RML)-Noor-Ul Azam Khan Romel", "PSL-(KNH)-Kazi Nurul Hoque"])
        wb.save(target)
        return True

    def load_profile(self):
        try:
            if PROFILE_FILE.exists():
                return PROFILE_FILE.read_text(encoding="utf-8")
        except Exception:
            pass
        return ""

    def save_profile(self, json_str):
        try:
            PROFILE_DIR.mkdir(parents=True, exist_ok=True)
            tmp_file = PROFILE_FILE.with_suffix(".tmp")
            tmp_file.write_text(json_str, encoding="utf-8")
            os.replace(tmp_file, PROFILE_FILE)
            return True
        except Exception as e:
            return str(e)

    def save_text_file(self, default_name, text_content):
# ... existing code ...
    def open_outlook_draft(self, to_addr, cc_addr, subject, html_body):
        if not HAS_WIN32COM:
            return "pywin32 not available; use .EML export instead."
        with COM_LOCK:
            pythoncom.CoInitialize()
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)
                mail.To = to_addr
                mail.CC = cc_addr
                mail.Subject = subject
                mail.HTMLBody = html_body
                mail.Display(False)
                return "OK"
            except Exception as e:
                return f"Outlook error: {e}"
            finally:
                pythoncom.CoUninitialize()

    def batch_outlook_drafts(self, drafts_json):
        if not HAS_WIN32COM:
            return "pywin32 not available; use .ZIP export instead."
        with COM_LOCK:
            pythoncom.CoInitialize()
            try:
                drafts = json.loads(drafts_json)
                outlook = win32com.client.Dispatch("Outlook.Application")
                count = 0
                for d in drafts:
                    mail = outlook.CreateItem(0)
                    mail.To = d.get("to", "")
                    mail.CC = d.get("cc", "")
                    mail.Subject = d.get("subject", "")
                    mail.HTMLBody = d.get("html", "")
                    mail.Display(False)
                    count += 1
                return f"OK:Opened {count} email drafts in Microsoft Outlook!"
            except Exception as e:
                return f"Outlook batch error: {e}"
            finally:
                pythoncom.CoUninitialize()
# ... existing code ...
