import base64
import io
import json
import os
import re
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
# COLOR & FORMATTING HELPERS
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
    """Normalize cell values cleanly for display."""
    if raw_val is None:
        return ""
    if isinstance(raw_val, float) and raw_val.is_integer():
        return str(int(raw_val))
    return str(raw_val).strip()


def is_dark_hex(hex_color):
    """Return True if a #RRGGBB color is dark enough to need white text."""
    if not hex_color or not isinstance(hex_color, str):
        return False
    h = hex_color.lstrip("#")
    if len(h) != 6:
        return False
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        luminance = (0.299 * r + 0.587 * g + 0.114 * b)
        return luminance < 140
    except Exception:
        return False


def _apply_tint(hex_color, tint):
    """Apply Office OpenXML tint (-1.0 to 1.0) to a #RRGGBB hex color."""
    if not hex_color or not tint:
        return hex_color
    h = hex_color.lstrip("#")
    if len(h) != 6:
        return hex_color
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

    def adj(c):
        if tint > 0:
            return int(round(c + (255 - c) * tint))
        return int(round(c * (1.0 + tint)))

    return f"#{max(0, min(255, adj(r))):02X}{max(0, min(255, adj(g))):02X}{max(0, min(255, adj(b))):02X}"


def extract_theme_palette(wb):
    """Extract the workbook's actual Office theme color palette (fallback to Office default)."""
    default_palette = [
        "#FFFFFF", "#000000", "#E7E6E6", "#44546A",
        "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
        "#5B9BD5", "#70AD47", "#0563C1", "#954F72",
    ]
    try:
        theme_xml = getattr(wb, "loaded_theme", None)
        if not theme_xml:
            return default_palette
        root = ET.fromstring(theme_xml)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        clr_scheme = root.find(".//a:clrScheme", ns)
        if clr_scheme is None:
            return default_palette
        order = [
            "lt1", "dk1", "lt2", "dk2",
            "accent1", "accent2", "accent3", "accent4",
            "accent5", "accent6", "hlink", "folHlink",
        ]
        palette = []
        for idx, tag in enumerate(order):
            el = clr_scheme.find(f"a:{tag}", ns)
            hex_val = default_palette[idx]
            if el is not None:
                srgb = el.find("a:srgbClr", ns)
                sys_clr = el.find("a:sysClr", ns)
                if srgb is not None and srgb.attrib.get("val"):
                    hex_val = "#" + srgb.attrib["val"].upper()
                elif sys_clr is not None and sys_clr.attrib.get("lastClr"):
                    hex_val = "#" + sys_clr.attrib["lastClr"].upper()
            palette.append(hex_val)
        return palette
    except Exception:
        return default_palette


def openpyxl_color_to_hex(color_obj, theme_palette):
    """Resolve an openpyxl Color object (rgb, theme+tint, or indexed) to #RRGGBB."""
    if color_obj is None:
        return None
    ctype = getattr(color_obj, "type", None)
    if ctype == "rgb":
        rgb = str(color_obj.rgb or "")
        if len(rgb) == 8:
            if rgb.startswith("00"):
                return None
            rgb = rgb[2:]
        if re.fullmatch(r"[0-9A-Fa-f]{6}", rgb):
            return "#" + rgb.upper()
    elif ctype == "theme":
        t_idx = getattr(color_obj, "theme", None)
        if isinstance(t_idx, int) and 0 <= t_idx < len(theme_palette):
            base = theme_palette[t_idx]
            tint = getattr(color_obj, "tint", 0.0) or 0.0
            return _apply_tint(base, tint)
    elif ctype == "indexed":
        idx = getattr(color_obj, "indexed", None)
        if isinstance(idx, int) and 0 <= idx < len(COLOR_INDEX):
            raw = COLOR_INDEX[idx]
            if len(raw) == 8:
                raw = raw[2:]
            if re.fullmatch(r"[0-9A-Fa-f]{6}", raw):
                return "#" + raw.upper()
    return None


def build_cf_lookup(ws, theme_palette):
    """Parse basic openpyxl Conditional Formatting rules (cellIs / containsText) for fallback mode."""
    entries = []
    try:
        for cf in ws.conditional_formatting:
            ranges = list(cf.sqref.ranges)
            for rule in cf.rules:
                dxf = getattr(rule, "dxf", None)
                if dxf is None:
                    continue
                st = {}
                if dxf.fill and dxf.fill.fgColor:
                    bg = openpyxl_color_to_hex(dxf.fill.fgColor, theme_palette)
                    if bg:
                        st["bg"] = bg
                if dxf.font:
                    fg = openpyxl_color_to_hex(dxf.font.color, theme_palette)
                    if fg:
                        st["color"] = fg
                    if dxf.font.bold:
                        st["bold"] = True
                    if dxf.font.italic:
                        st["italic"] = True
                if not st:
                    continue
                entries.append({
                    "ranges": ranges,
                    "type": getattr(rule, "type", ""),
                    "operator": getattr(rule, "operator", ""),
                    "formula": list(getattr(rule, "formula", []) or []),
                    "text": getattr(rule, "text", None),
                    "style": st,
                })
    except Exception:
        pass
    return entries


def _matches_cf_rule(entry, raw_val, str_val):
    rtype = entry["type"]
    op = entry["operator"]
    formulas = entry["formula"]
    if rtype == "containsText" and entry["text"]:
        return str(entry["text"]).lower() in str_val.lower()
    if rtype == "cellIs" and formulas:
        target = str(formulas[0]).strip().strip('"')
        if op == "equal":
            if str_val.lower() == target.lower():
                return True
            try:
                return float(raw_val) == float(target)
            except Exception:
                return False
        try:
            num_val = float(raw_val)
            num_t = float(target)
            if op == "lessThan":
                return num_val < num_t
            if op == "lessThanOrEqual":
                return num_val <= num_t
            if op == "greaterThan":
                return num_val > num_t
            if op == "greaterThanOrEqual":
                return num_val >= num_t
            if op == "notEqual":
                return num_val != num_t
        except Exception:
            return False
    return False


def extract_cell_style_openpyxl(cell, row_num, col_num, raw_val, str_val, theme_palette, cf_entries):
    """Extract fill, font color, bold, italic, and alignment from an openpyxl cell + CF rules."""
    st = {}
    fill = getattr(cell, "fill", None)
    if fill and fill.fill_type and fill.fill_type != "none":
        bg = openpyxl_color_to_hex(fill.fgColor, theme_palette)
        if bg:
            st["bg"] = bg

    font = getattr(cell, "font", None)
    if font:
        fg = openpyxl_color_to_hex(font.color, theme_palette)
        if fg:
            st["color"] = fg
        if font.bold:
            st["bold"] = True
        if font.italic:
            st["italic"] = True

    align = getattr(cell, "alignment", None)
    if align and align.horizontal in ("left", "center", "right"):
        st["align"] = align.horizontal

    for entry in cf_entries:
        in_range = any(
            rng.min_row <= row_num <= rng.max_row and rng.min_col <= col_num <= rng.max_col
            for rng in entry["ranges"]
        )
        if in_range and _matches_cf_rule(entry, raw_val, str_val):
            st.update(entry["style"])

    return st if st else None


# ==============================================================================
# NATIVE EXCEL DISPLAYFORMAT ENGINE (EXACT ON-SCREEN COLORS & CONDITIONAL FORMATS)
# ==============================================================================
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

            raw_rows = list(raw_matrix)
            while raw_rows and all(v is None or str(v).strip() == "" for v in raw_rows[-1]):
                raw_rows.pop()
            if not raw_rows:
                return {"ok": False, "error": f"Worksheet '{active_sheet}' is empty."}

            # Fall back to fast openpyxl for very large sheets (>600 rows)
            if len(raw_rows) > 600:
                return None

            grid = [[_format_cell_val(v) for v in r] for r in raw_rows]
            num_rows = len(grid)
            num_cols = len(grid[0])

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

            active_col_indices = sorted({h["idx"] for h in header_cols})
            style_grid = [[None] * num_cols for _ in range(num_rows)]
            merge_master = {}
            merge_covered = set()
            seen_merge_addrs = set()
            align_map = {-4131: "left", -4108: "center", -4152: "right"}

            for r in range(header_row_idx, num_rows):
                if all(grid[r][c] == "" for c in range(num_cols)):
                    continue
                excel_r = start_r + r
                for c in active_col_indices:
                    excel_c = start_c + c
                    cell = ws.Cells(excel_r, excel_c)

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

                    df = cell.DisplayFormat
                    st = {}
                    interior = df.Interior
                    if interior.ColorIndex != -4142:
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
                engine_used="Excel DisplayFormat",
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
# OPENPYXL FALLBACK PARSER
# ==============================================================================
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
        engine_used="openpyxl",
    )


def _build_regions_from_grid(
    grid, style_grid, merge_master, merge_covered,
    header_row_idx, header_cols, sheet_names, active_sheet,
    engine_used="openpyxl",
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

        matches_header = sum(
            1 for h in header_cols if h["idx"] < len(row) and row[h["idx"]] and row[h["idx"]].lower() == h["name"].lower()
        )
        if matches_header >= 2:
            header_cols = [{"name": v, "idx": idx} for idx, v in enumerate(row) if v != ""]
            last_group_vals = {}
            continue

        first_two = " ".join(non_empty[:2]).lower()
        if re.search(r"\btotal\b", first_two) and not any(re.match(r"^[A-Z]\d{3}$", v, re.I) for v in non_empty):
            continue

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
        body { font-family: 'Inter', sans-serif; }
        .email-body {
            font-family: Calibri, Arial, sans-serif;
            font-size: 11pt;
            color: #000000;
        }
        .email-table {
            border-collapse: collapse;
            width: auto;
            border: 1px solid #000000;
            font-size: 10.5pt;
        }
        .email-table th {
            background-color: #e6005c;
            color: #ffffff;
            border: 1px solid #000000;
            padding: 5px 12px;
            text-align: center;
            font-weight: bold;
            white-space: nowrap;
        }
        .email-table td {
            border: 1px solid #000000;
            padding: 4px 10px;
            text-align: center;
            background-color: #fdf6e3;
            vertical-align: middle;
            white-space: nowrap;
        }
        .email-table.excel-retained {
            border: 1px solid #d1d5db;
        }
        .email-table.excel-retained th,
        .email-table.excel-retained td {
            border: 1px solid #d1d5db;
        }
        .email-table .totals-row td {
            background-color: #ffff00;
            font-weight: bold;
        }
        .email-table .note-row td {
            background-color: #ffff00;
            font-weight: bold;
        }
        .email-table .note-row td.note-value {
            background-color: #ffffff;
            color: #1f4e79;
            font-weight: normal;
            text-align: left;
            padding-left: 10px;
        }
        .text-green { color: #00b050 !important; font-weight: bold; }
        .text-red { color: #ff0000 !important; font-weight: bold; }
        .bg-red-row td { background-color: #ff0000 !important; color: #ffffff !important; font-weight: bold; }
        .bg-red-row td.text-green, .bg-red-row td.text-red { color: #ffffff !important; }
    </style>
</head>
<body class="bg-slate-50 min-h-screen text-slate-800 antialiased pb-12">

    <header class="bg-white border-b border-slate-200 sticky top-0 z-10 shadow-sm">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex flex-wrap justify-between items-center gap-4">
            <div class="flex items-center space-x-3">
                <div class="bg-indigo-600 text-white p-2 rounded-lg shadow-sm">
                    <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"></path>
                    </svg>
                </div>
                <div>
                    <div class="flex items-center gap-2">
                        <h1 class="text-lg font-bold text-slate-900 tracking-tight">Outlook Email Automator</h1>
                        <span id="memoryBadge" class="text-[10px] font-semibold px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200">
                            Windows Memory Active
                        </span>
                    </div>
                    <p class="text-xs text-slate-500">Auto-saving routing, templates, fonts, signature &amp; native Excel formatting</p>
                </div>
            </div>

            <div class="flex flex-wrap items-center gap-2">
                <span id="copyFeedback" class="text-emerald-600 bg-emerald-50 border border-emerald-200 px-3 py-1 rounded-full text-xs font-medium opacity-0 transition-opacity duration-200">
                    Copied
                </span>
                <button onclick="copyText('emailSubject')" class="inline-flex items-center gap-1.5 bg-white border border-slate-300 text-slate-700 px-3 py-2 rounded-lg hover:bg-slate-50 transition text-xs font-semibold shadow-sm">
                    Copy Subject
                </button>
                <button onclick="copyRichText('emailBodyContent')" class="inline-flex items-center gap-1.5 bg-white border border-slate-300 text-slate-700 px-3 py-2 rounded-lg hover:bg-slate-50 transition text-xs font-semibold shadow-sm">
                    Copy Body
                </button>
                <button onclick="openOutlookDraft()" class="inline-flex items-center gap-1.5 bg-indigo-600 text-white px-3.5 py-2 rounded-lg hover:bg-indigo-700 transition text-xs font-semibold shadow-sm">
                    Open in Outlook
                </button>
                <button onclick="batchOutlookDrafts()" class="inline-flex items-center gap-1.5 bg-emerald-600 text-white px-3.5 py-2 rounded-lg hover:bg-emerald-700 transition text-xs font-semibold shadow-sm">
                    Batch Outlook Drafts
                </button>
                <button onclick="exportEML()" class="inline-flex items-center gap-1.5 bg-slate-800 text-white px-3 py-2 rounded-lg hover:bg-slate-900 transition text-xs font-semibold shadow-sm">
                    .EML
                </button>
                <button onclick="batchExportZip()" class="inline-flex items-center gap-1.5 bg-slate-700 text-white px-3 py-2 rounded-lg hover:bg-slate-800 transition text-xs font-semibold shadow-sm">
                    .ZIP
                </button>
            </div>
        </div>
    </header>

    <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mt-6">
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
            <div class="lg:col-span-5 space-y-4">
                <div id="statusToast" class="hidden rounded-lg p-3 text-xs font-medium border flex items-center justify-between transition-all">
                    <span id="statusToastMsg"></span>
                    <button onclick="document.getElementById('statusToast').classList.add('hidden')" class="text-slate-400 hover:text-slate-600 font-bold ml-2">x</button>
                </div>

                <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4 space-y-3.5">
                    <div class="flex flex-wrap items-center justify-between gap-2">
                        <span class="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-1.5">
                            <span class="w-2 h-2 rounded-full bg-indigo-500"></span>
                            1. Data Sources
                        </span>
                        <div class="flex items-center gap-1.5">
                            <button onclick="toggleBubble('tableUploadBubble')" id="btnTableBubble" class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-indigo-50 text-indigo-700 hover:bg-indigo-100 border border-indigo-200 transition cursor-pointer">
                                <span id="tableBubbleText">Table (.xlsx)</span>
                            </button>
                            <button onclick="toggleBubble('routingUploadBubble')" id="btnRoutingBubble" class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-slate-100 text-slate-700 hover:bg-slate-200 border border-slate-200 transition cursor-pointer">
                                <span id="routingBubbleText">Routing (To/CC)</span>
                            </button>
                        </div>
                    </div>

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

                    <hr class="border-slate-100">

                    <div class="space-y-2.5">
                        <h2 class="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-1.5">
                            <span class="w-2 h-2 rounded-full bg-indigo-500"></span>
                            2. Parameters &amp; Layout
                        </h2>
                        <div class="grid grid-cols-3 gap-2">
                            <div>
                                <label for="regionSelect" class="block text-[10px] font-semibold text-slate-500 uppercase mb-0.5">Region (RSE)</label>
                                <select id="regionSelect" class="w-full border border-slate-300 rounded-md px-2 py-1.5 text-xs bg-white font-semibold text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none cursor-pointer" onchange="renderEmail()">
                                    <option value="R4">R4</option>
                                    <option value="A2">A2</option>
                                    <option value="N4">N4</option>
                                </select>
                            </div>
                            <div>
                                <label for="shortDate" class="block text-[10px] font-semibold text-slate-500 uppercase mb-0.5">Short Date</label>
                                <input type="text" id="shortDate" class="w-full border border-slate-300 rounded-md px-2 py-1.5 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none" oninput="scheduleTextRender()">
                            </div>
                            <div>
                                <label for="longDate" class="block text-[10px] font-semibold text-slate-500 uppercase mb-0.5">Long Date</label>
                                <input type="text" id="longDate" class="w-full border border-slate-300 rounded-md px-2 py-1.5 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none" oninput="scheduleTextRender()">
                            </div>
                        </div>

                        <div class="grid grid-cols-2 gap-x-3 gap-y-1.5 pt-1 text-[11px] text-slate-700">
                            <label class="inline-flex items-center gap-1.5 cursor-pointer">
                                <input type="checkbox" id="mergeGroupCells" checked onchange="onSettingChange()" class="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5">
                                <span>Merge repeated cells</span>
                            </label>
                            <label class="inline-flex items-center gap-1.5 cursor-pointer">
                                <input type="checkbox" id="showTotalRow" checked onchange="onSettingChange()" class="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5">
                                <span>Show Total row</span>
                            </label>
                            <label class="inline-flex items-center gap-1.5 cursor-pointer">
                                <input type="checkbox" id="includeNoteRows" checked onchange="onSettingChange()" class="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5">
                                <span>Include bottom notes</span>
                            </label>
                            <label class="inline-flex items-center gap-1.5 cursor-pointer font-medium text-indigo-700">
                                <input type="checkbox" id="retainExcelDesign" checked onchange="onSettingChange()" class="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5">
                                <span>Retain Excel table design</span>
                            </label>
                        </div>

                        <div class="grid grid-cols-3 gap-2 pt-1.5 border-t border-slate-100">
                            <div>
                                <label for="globalFontFamily" class="block text-[10px] font-semibold text-slate-500 uppercase mb-0.5">Email Font</label>
                                <select id="globalFontFamily" onchange="onSettingChange()" class="w-full border border-slate-300 rounded-md px-2 py-1 text-xs bg-white text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                    <option value="Calibri, Arial, sans-serif" selected>Calibri</option>
                                    <option value="Segoe UI, Arial, sans-serif">Segoe UI</option>
                                    <option value="Arial, Helvetica, sans-serif">Arial</option>
                                    <option value="Aptos, Calibri, sans-serif">Aptos</option>
                                    <option value="Verdana, Geneva, sans-serif">Verdana</option>
                                    <option value="Tahoma, Geneva, sans-serif">Tahoma</option>
                                    <option value="Times New Roman, Times, serif">Times New Roman</option>
                                </select>
                            </div>
                            <div>
                                <label for="bodyFontSize" class="block text-[10px] font-semibold text-slate-500 uppercase mb-0.5">Body &amp; Sig Size</label>
                                <select id="bodyFontSize" onchange="onSettingChange()" class="w-full border border-slate-300 rounded-md px-2 py-1 text-xs bg-white text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                    <option value="9pt">9 pt</option>
                                    <option value="10pt">10 pt</option>
                                    <option value="10.5pt">10.5 pt</option>
                                    <option value="11pt" selected>11 pt</option>
                                    <option value="12pt">12 pt</option>
                                    <option value="13pt">13 pt</option>
                                    <option value="14pt">14 pt</option>
                                </select>
                            </div>
                            <div>
                                <label for="tableFontSize" class="block text-[10px] font-semibold text-slate-500 uppercase mb-0.5">Table Font Size</label>
                                <select id="tableFontSize" onchange="onSettingChange()" class="w-full border border-slate-300 rounded-md px-2 py-1 text-xs bg-white text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                    <option value="8.5pt">8.5 pt</option>
                                    <option value="9pt">9 pt</option>
                                    <option value="9.5pt">9.5 pt</option>
                                    <option value="10pt">10 pt</option>
                                    <option value="10.5pt" selected>10.5 pt</option>
                                    <option value="11pt">11 pt</option>
                                    <option value="12pt">12 pt</option>
                                </select>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4 space-y-2.5">
                    <div class="flex justify-between items-center">
                        <h2 class="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-1.5">
                            <span class="w-2 h-2 rounded-full bg-indigo-500"></span>
                            3. Customize Template
                        </h2>
                        <span class="text-[10px] text-slate-400">Auto-saved to disk</span>
                    </div>
                    <div class="flex flex-wrap gap-1.5 items-center bg-slate-50 p-1.5 rounded-lg border border-slate-200/80">
                        <span class="text-[10px] text-slate-500 font-semibold uppercase px-1">Variables:</span>
                        <span draggable="true" ondragstart="handleDragStart(event, '{region}')" onclick="insertVariable('{region}')" class="cursor-pointer select-none text-[11px] font-mono bg-white hover:bg-indigo-50 text-indigo-700 px-2 py-0.5 rounded-full border border-indigo-200 transition">{region}</span>
                        <span draggable="true" ondragstart="handleDragStart(event, '{shortDate}')" onclick="insertVariable('{shortDate}')" class="cursor-pointer select-none text-[11px] font-mono bg-white hover:bg-indigo-50 text-indigo-700 px-2 py-0.5 rounded-full border border-indigo-200 transition">{shortDate}</span>
                        <span draggable="true" ondragstart="handleDragStart(event, '{longDate}')" onclick="insertVariable('{longDate}')" class="cursor-pointer select-none text-[11px] font-mono bg-white hover:bg-indigo-50 text-indigo-700 px-2 py-0.5 rounded-full border border-indigo-200 transition">{longDate}</span>
                    </div>
                    <div>
                        <label class="block text-[10px] font-semibold text-slate-500 uppercase mb-0.5">Subject Line</label>
                        <input type="text" id="subjectTemplate" onfocus="lastFocusedTarget = this" ondragover="allowDrop(event)" ondrop="handleDrop(event)" class="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none" value="RE: Shelving Status of Ventrax HFA Inhaler {shortDate} of {region}" oninput="scheduleTextRender(true)">
                    </div>
                    <div>
                        <label class="block text-[10px] font-semibold text-slate-500 uppercase mb-0.5">Opening Body Text</label>
                        <div class="flex flex-wrap items-center gap-1 bg-slate-100 border border-b-0 border-slate-300 rounded-t-lg px-2 py-1">
                            <button type="button" onmousedown="event.preventDefault(); execBodyCmd('bold')" class="px-2 py-0.5 rounded hover:bg-white text-xs font-bold text-slate-800">B</button>
                            <button type="button" onmousedown="event.preventDefault(); execBodyCmd('italic')" class="px-2 py-0.5 rounded hover:bg-white text-xs italic font-serif text-slate-800">I</button>
                            <button type="button" onmousedown="event.preventDefault(); execBodyCmd('underline')" class="px-2 py-0.5 rounded hover:bg-white text-xs underline text-slate-800">U</button>
                            <span class="h-3.5 w-[1px] bg-slate-300 mx-0.5"></span>
                            <button type="button" onmousedown="event.preventDefault(); changeRichFontSize('bodyEditor', -1)" class="px-1.5 py-0.5 rounded hover:bg-white text-[11px] font-semibold text-slate-700">A-</button>
                            <button type="button" onmousedown="event.preventDefault(); changeRichFontSize('bodyEditor', 1)" class="px-1.5 py-0.5 rounded hover:bg-white text-[11px] font-semibold text-slate-700">A+</button>
                            <select onchange="applyExactFontSize('bodyEditor', this.value); this.selectedIndex=0;" class="text-[11px] bg-white border border-slate-200 rounded px-1 py-0.5 text-slate-700 cursor-pointer">
                                <option value="">Size</option>
                                <option value="9pt">9 pt</option>
                                <option value="10pt">10 pt</option>
                                <option value="11pt">11 pt</option>
                                <option value="12pt">12 pt</option>
                                <option value="14pt">14 pt</option>
                                <option value="16pt">16 pt</option>
                            </select>
                            <span class="h-3.5 w-[1px] bg-slate-300 mx-0.5"></span>
                            <label class="inline-flex items-center gap-1 px-1.5 py-0.5 rounded hover:bg-white cursor-pointer text-[11px] font-semibold text-slate-700">
                                <span class="underline decoration-2" id="colorIndicator">A</span>
                                <input type="color" value="#e6005c" onchange="applyBodyColor(this.value)" class="w-3.5 h-3.5 p-0 border-0 bg-transparent cursor-pointer">
                            </label>
                            <button type="button" onmousedown="event.preventDefault(); applyBodyColor('#000000')" class="w-3.5 h-3.5 rounded-full bg-black border border-slate-300"></button>
                            <button type="button" onmousedown="event.preventDefault(); applyBodyColor('#1f4e79')" class="w-3.5 h-3.5 rounded-full bg-[#1f4e79] border border-slate-300"></button>
                            <button type="button" onmousedown="event.preventDefault(); applyBodyColor('#dc2626')" class="w-3.5 h-3.5 rounded-full bg-red-600 border border-slate-300"></button>
                            <button type="button" onmousedown="event.preventDefault(); applyBodyColor('#16a34a')" class="w-3.5 h-3.5 rounded-full bg-green-600 border border-slate-300"></button>
                            <button type="button" onmousedown="event.preventDefault(); execBodyCmd('hiliteColor', '#fef08a')" class="px-1.5 py-0.5 rounded bg-yellow-200 text-[10px] font-bold text-slate-800">H</button>
                            <button type="button" onmousedown="event.preventDefault(); execBodyCmd('removeFormat'); onSettingChange();" class="ml-auto px-1.5 py-0.5 rounded hover:bg-white text-[10px] text-slate-500">Clear</button>
                        </div>
                        <div id="bodyEditor" contenteditable="true" onfocus="lastFocusedTarget = this" ondragover="allowDrop(event)" ondrop="handleDrop(event)" oninput="scheduleTextRender(true)" class="w-full min-h-[85px] border border-slate-300 rounded-b-lg px-2.5 py-2 text-xs text-slate-800 bg-white focus:ring-2 focus:ring-indigo-500 focus:outline-none leading-relaxed">Dear Vai,<br><br>Please find below the MPE-wise shelving status of Ventrax HFA Inhaler (VHI) as of {longDate}</div>
                    </div>
                </div>

                <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
                    <div class="flex items-center justify-between">
                        <label class="inline-flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" id="includeSignature" checked onchange="onSettingChange()" class="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5 cursor-pointer">
                            <span class="text-xs font-bold text-slate-900 uppercase tracking-wider">4. Append Signature</span>
                        </label>
                        <button onclick="toggleBubble('signatureEditorDrawer')" class="text-[11px] font-semibold text-indigo-600 hover:text-indigo-800 px-2 py-0.5 rounded-full hover:bg-indigo-50 transition">
                            Edit Details
                        </button>
                    </div>

                    <div id="signatureEditorDrawer" class="mt-3 pt-2.5 border-t border-slate-100 space-y-2 text-xs">
                        <div class="flex items-center justify-between">
                            <span class="text-[10px] font-semibold text-slate-500 uppercase">Signature Editor (Auto-Saved)</span>
                            <div class="flex items-center gap-2">
                                <label class="text-[10px] text-indigo-600 hover:text-indigo-800 font-semibold cursor-pointer">
                                    + Custom Logo
                                    <input type="file" accept="image/*" class="hidden" onchange="handleSigLogoUpload(event)">
                                </label>
                                <button onclick="resetDefaultSignature()" class="text-[10px] text-slate-400 hover:text-slate-600">Reset</button>
                            </div>
                        </div>

                        <div class="flex flex-wrap items-center gap-1 bg-slate-100 border border-b-0 border-slate-300 rounded-t-lg px-2 py-1">
                            <button type="button" onmousedown="event.preventDefault(); execSigCmd('bold')" class="px-1.5 py-0.5 rounded hover:bg-white text-[11px] font-bold text-slate-800">B</button>
                            <button type="button" onmousedown="event.preventDefault(); execSigCmd('italic')" class="px-1.5 py-0.5 rounded hover:bg-white text-[11px] italic font-serif text-slate-800">I</button>
                            <button type="button" onmousedown="event.preventDefault(); execSigCmd('underline')" class="px-1.5 py-0.5 rounded hover:bg-white text-[11px] underline text-slate-800">U</button>
                            <span class="h-3 w-[1px] bg-slate-300 mx-0.5"></span>
                            <button type="button" onmousedown="event.preventDefault(); changeRichFontSize('sigPasteBox', -1)" class="px-1.5 py-0.5 rounded hover:bg-white text-[10px] font-semibold text-slate-700">A-</button>
                            <button type="button" onmousedown="event.preventDefault(); changeRichFontSize('sigPasteBox', 1)" class="px-1.5 py-0.5 rounded hover:bg-white text-[10px] font-semibold text-slate-700">A+</button>
                            <select onchange="applyExactFontSize('sigPasteBox', this.value); this.selectedIndex=0;" class="text-[10px] bg-white border border-slate-200 rounded px-1 py-0.5 text-slate-700 cursor-pointer">
                                <option value="">Size</option>
                                <option value="8pt">8 pt</option>
                                <option value="9pt">9 pt</option>
                                <option value="10pt">10 pt</option>
                                <option value="11pt">11 pt</option>
                                <option value="12pt">12 pt</option>
                                <option value="14pt">14 pt</option>
                            </select>
                            <span class="h-3 w-[1px] bg-slate-300 mx-0.5"></span>
                            <input type="color" value="#1f4e79" onchange="applySigColor(this.value)" class="w-3.5 h-3.5 p-0 border-0 bg-transparent cursor-pointer">
                            <button type="button" onmousedown="event.preventDefault(); execSigCmd('removeFormat'); onSettingChange();" class="ml-auto px-1.5 py-0.5 rounded hover:bg-white text-[10px] text-slate-500">Clear</button>
                        </div>

                        <div id="sigPasteBox" contenteditable="true" oninput="scheduleSigRender(true)" class="w-full min-h-[90px] max-h-48 overflow-y-auto border border-slate-300 rounded-b-lg p-2.5 bg-slate-50/50 focus:bg-white focus:ring-2 focus:ring-indigo-500 focus:outline-none text-[11px] leading-relaxed">
                            <p style="margin: 0 0 8px 0; color: #1f4e79; font-weight: bold;">Thanks and regards,</p>
                            <p style="margin: 0; color: #595959;">
                                <span style="color: #2e75b6; font-weight: bold;">S.M. Masrakul Islam Medha</span> | Brand Executive, Strategic Brand Management | Beximco Pharmaceuticals PLC. | 19 Dhanmondi R/A, Road #7, Dhaka 1205 | Bangladesh
                            </p>
                            <p style="margin: 3px 0 0 0; color: #595959;">
                                | <strong>Cell</strong> +88 <a href="tel:+8801760021915" style="color: #954f72; text-decoration: underline;">01760021915</a> | <strong>Fax</strong> <a href="tel:+880258613888" style="color: #954f72; text-decoration: underline;">+880 2 58613888</a> | <a href="mailto:mke@bpl.net" style="color: #0563c1; text-decoration: underline;">mke@bpl.net</a> | <a href="http://www.beximcopharma.com" style="color: #0563c1; text-decoration: underline;">www.beximcopharma.com</a>
                            </p>
                        </div>

                        <div class="flex items-center justify-between text-[10px] text-slate-500 bg-slate-50 px-2.5 py-1.5 rounded-md border border-slate-200/70">
                            <label class="inline-flex items-center gap-1.5 cursor-pointer font-medium text-slate-600">
                                <input type="checkbox" id="autoDetectLogo" checked onchange="onSettingChange()" class="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-3 h-3">
                                Auto-show Beximco Pharma logo
                            </label>
                            <button id="clearCustomLogoBtn" onclick="clearCustomSigLogo()" class="hidden text-rose-600 hover:underline font-semibold">Remove Uploaded Logo</button>
                        </div>
                    </div>
                </div>
            </div>

            <div class="lg:col-span-7">
                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                    <div class="bg-slate-800 text-slate-200 px-4 py-2.5 text-xs font-medium flex items-center justify-between">
                        <div class="flex items-center space-x-2">
                            <span class="w-2.5 h-2.5 rounded-full bg-rose-500 inline-block"></span>
                            <span class="w-2.5 h-2.5 rounded-full bg-amber-400 inline-block"></span>
                            <span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span>
                            <span class="ml-2 text-slate-300 font-semibold">Live Outlook Draft Preview</span>
                        </div>
                        <span class="text-[11px] text-slate-400">Edits to To/Cc are remembered per region</span>
                    </div>

                    <div class="bg-slate-50/70 border-b border-slate-200 p-4 text-xs space-y-2">
                        <div class="grid grid-cols-[55px_1fr] items-center">
                            <span class="text-slate-400 font-medium">From</span>
                            <span class="text-slate-700 font-medium px-2 py-0.5">SBM-(MKE)-S.M. Masrakul Islam Medha</span>
                        </div>
                        <div class="grid grid-cols-[55px_1fr] items-center">
                            <span class="text-slate-400 font-medium">To</span>
                            <input type="text" id="emailTo" oninput="saveCurrentRoutingEdit()" class="border border-slate-200 rounded-md px-2.5 py-1 w-full bg-white text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition" value="">
                        </div>
                        <div class="grid grid-cols-[55px_1fr] items-center">
                            <span class="text-slate-400 font-medium">Cc</span>
                            <input type="text" id="emailCc" oninput="saveCurrentRoutingEdit()" class="border border-slate-200 rounded-md px-2.5 py-1 w-full bg-white text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition" value="">
                        </div>
                        <div class="grid grid-cols-[55px_1fr] items-center pt-1.5 border-t border-slate-200/80">
                            <span class="text-slate-400 font-medium">Subject</span>
                            <div class="text-slate-900 font-bold text-sm px-2 py-0.5" id="emailSubject"></div>
                        </div>
                    </div>

                    <div class="p-6 md:p-8 bg-white overflow-x-auto">
                        <div class="email-body" id="emailBodyContent">
                            <div id="dynamicBodyText" class="text-gray-900"></div>
                            <div id="tableContainer"></div>
                            <div id="signatureContainer"></div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </main>

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

        let routingData = {
            "R4": { to: "PSL-(NAL)-Md. Nur-A- Alam", cc: "SBM-(ALM)-Mahabub Alam.; PSL-(KNH)-Kazi Nurul Hoque; PSL-(MDR)-Masud Rana khan; PSL-(ANL)-Ahsanul Kabir;" },
            "A2": { to: "PSL-(MZO)-Md. Zakir Hossain", cc: "SBM-(ALM)-Mahabub Alam.; PSL-(KNH)-Kazi Nurul Hoque; PSL-(MDR)-Masud Rana khan; PSL-(ANL)-Ahsanul Kabir;" },
            "N4": { to: "PSL-(KMO)-Kazi Md. Mostafa", cc: "SBM-(RML)-Noor-Ul Azam Khan Romel; SBM-(ALM)-Mahabub Alam.; PSL-(KNH)-Kazi Nurul Hoque;" }
        };

        let tableData = {
            "R4": [
                { "ASE": "R410", "T.Code": "R411", "Shelving Till 15": 1, "Shelving Till 19": 3, "Shelving Till 25": 14, "Shelving Till 28": 14, "New Shelving": 0, "Remarks": "" },
                { "ASE": "R410", "T.Code": "R412", "Shelving Till 15": 1, "Shelving Till 19": 7, "Shelving Till 25": 7, "Shelving Till 28": 9, "New Shelving": 2, "Remarks": "Increased but still below 10. Need more shelving." },
                { "ASE": "R420", "T.Code": "R421", "Shelving Till 15": 4, "Shelving Till 19": 4, "Shelving Till 25": 12, "Shelving Till 28": 12, "New Shelving": 0, "Remarks": "" },
                { "ASE": "R420", "T.Code": "R423", "Shelving Till 15": 4, "Shelving Till 19": 4, "Shelving Till 25": 5, "Shelving Till 28": 13, "New Shelving": 8, "Remarks": "" }
            ],
            "A2": [
                { "RSE": "A2", "ASE": "A210", "MPE": "A211", "Total PB": 3, "Remaining PB Selection": 2 },
                { "RSE": "A2", "ASE": "A220", "MPE": "A221", "Total PB": "No PB selected", "Remaining PB Selection": "No PB selected" },
                { "RSE": "A2", "ASE": "A220", "MPE": "A222", "Total PB": "No PB selected", "Remaining PB Selection": "No PB selected" },
                { "RSE": "A2", "ASE": "A220", "MPE": "A233", "Total PB": "No PB selected", "Remaining PB Selection": "No PB selected" },
                { "RSE": "A2", "ASE": "A250", "MPE": "A251", "Total PB": 1, "Remaining PB Selection": 4 }
            ],
            "N4": [
                { "ASE": "N410", "T.Code": "N411", "Shelving Till 15": 12, "Shelving Till 19": 18, "Shelving Till 25": 20, "Shelving Till 28": 22, "New Shelving": 2, "Remarks": "" },
                { "ASE": "N420", "T.Code": "N421", "Shelving Till 15": 4, "Shelving Till 19": 10, "Shelving Till 25": 14, "Shelving Till 28": 12, "New Shelving": -2, "Remarks": "" }
            ]
        };

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

        async function loadMemoryFromDisk() {
            let raw = null;
            try {
                if (window.pywebview && window.pywebview.api) {
                    raw = await window.pywebview.api.load_profile();
                } else {
                    raw = localStorage.getItem('outlook_automator_profile');
                }
            } catch (e) { console.error("Memory load error:", e); }

            if (!raw) return;
            lastSavedJson = raw;
            try {
                const p = JSON.parse(raw);
                if (p.routingData) routingData = p.routingData;
                if (p.subjectTemplate !== undefined) document.getElementById('subjectTemplate').value = p.subjectTemplate;
                if (p.bodyHtml !== undefined) document.getElementById('bodyEditor').innerHTML = p.bodyHtml;
                if (p.sigHtml !== undefined) document.getElementById('sigPasteBox').innerHTML = p.sigHtml;
                if (p.includeSignature !== undefined) document.getElementById('includeSignature').checked = p.includeSignature;
                if (p.autoDetectLogo !== undefined) document.getElementById('autoDetectLogo').checked = p.autoDetectLogo;
                if (p.customSigLogoData) {
                    customSigLogoData = p.customSigLogoData;
                    document.getElementById('clearCustomLogoBtn').classList.remove('hidden');
                }
                if (p.globalFontFamily) {
                    const fontSel = document.getElementById('globalFontFamily');
                    const cleanVal = p.globalFontFamily.replace(/['"]/g, '');
                    const matchOpt = Array.from(fontSel.options).find(o => o.value === p.globalFontFamily || o.value === cleanVal);
                    fontSel.value = matchOpt ? matchOpt.value : fontSel.options[0].value;
                }
                if (p.bodyFontSize) document.getElementById('bodyFontSize').value = p.bodyFontSize;
                if (p.tableFontSize) document.getElementById('tableFontSize').value = p.tableFontSize;
                if (p.mergeGroupCells !== undefined) document.getElementById('mergeGroupCells').checked = p.mergeGroupCells;
                if (p.showTotalRow !== undefined) document.getElementById('showTotalRow').checked = p.showTotalRow;
                if (p.includeNoteRows !== undefined) document.getElementById('includeNoteRows').checked = p.includeNoteRows;
                if (p.retainExcelDesign !== undefined) document.getElementById('retainExcelDesign').checked = p.retainExcelDesign;
                const rCount = Object.keys(routingData).length;
                if (rCount > 0) document.getElementById('routingBubbleText').innerText = `Routing (${rCount})`;
            } catch (e) { console.error("Profile parse error:", e); }
        }

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

        function saveCurrentRoutingEdit() {
            const region = document.getElementById('regionSelect').value;
            if (!region) return;
            if (!routingData[region]) routingData[region] = {};
            routingData[region].to = document.getElementById('emailTo').value;
            routingData[region].cc = document.getElementById('emailCc').value;
            scheduleSaveMemory();
        }

        function toggleBubble(drawerId) {
            const drawers = ['tableUploadBubble', 'routingUploadBubble', 'signatureEditorDrawer'];
            drawers.forEach(id => {
                const el = document.getElementById(id);
                if (!el) return;
                if (id === drawerId) el.classList.toggle('hidden');
                else if (drawerId !== 'signatureEditorDrawer' && id !== 'signatureEditorDrawer') el.classList.add('hidden');
            });
        }

        function showStatus(msg, isError = false) {
            const toast = document.getElementById('statusToast');
            const msgEl = document.getElementById('statusToastMsg');
            toast.className = `rounded-lg p-3 text-xs font-medium border flex items-center justify-between transition-all ${
                isError ? 'bg-rose-50 text-rose-700 border-rose-200' : 'bg-emerald-50 text-emerald-700 border-emerald-200'
            }`;
            msgEl.innerText = msg;
            toast.classList.remove('hidden');
            setTimeout(() => { toast.classList.add('hidden'); }, 4500);
        }

        function execBodyCmd(cmd, val = null) {
            const editor = document.getElementById('bodyEditor');
            editor.focus();
            document.execCommand(cmd, false, val);
            scheduleTextRender(true);
        }
        function applyBodyColor(hex) {
            document.getElementById('colorIndicator').style.textDecorationColor = hex;
            execBodyCmd('foreColor', hex);
        }
        function execSigCmd(cmd, val = null) {
            const editor = document.getElementById('sigPasteBox');
            editor.focus();
            document.execCommand(cmd, false, val);
            scheduleSigRender(true);
        }
        function applySigColor(hex) { execSigCmd('foreColor', hex); }

        function applyExactFontSize(editorId, sizeStr) {
            if (!sizeStr) return;
            const editor = document.getElementById(editorId);
            editor.focus();
            const sel = window.getSelection();
            if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return;
            const range = sel.getRangeAt(0);
            const span = document.createElement('span');
            span.style.fontSize = sizeStr;
            span.appendChild(range.extractContents());
            range.insertNode(span);
            sel.removeAllRanges();
            const newRange = document.createRange();
            newRange.selectNodeContents(span);
            sel.addRange(newRange);
            if (editorId === 'bodyEditor') scheduleTextRender(true);
            else scheduleSigRender(true);
        }

        function changeRichFontSize(editorId, delta) {
            const editor = document.getElementById(editorId);
            editor.focus();
            const sel = window.getSelection();
            if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return;
            let currentPt = parseFloat(document.getElementById('bodyFontSize').value) || 11;
            const anchorEl = sel.anchorNode.nodeType === 3 ? sel.anchorNode.parentElement : sel.anchorNode;
            if (anchorEl && editor.contains(anchorEl)) {
                const px = parseFloat(window.getComputedStyle(anchorEl).fontSize);
                if (!isNaN(px)) currentPt = Math.round(px * 0.75);
            }
            const nextPt = Math.max(8, Math.min(28, currentPt + delta));
            applyExactFontSize(editorId, `${nextPt}pt`);
        }

        function handleDragStart(e, variableText) {
            e.dataTransfer.setData('text/plain', variableText);
            e.dataTransfer.effectAllowed = 'copy';
        }
        function allowDrop(e) { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; }
        function handleDrop(e) {
            e.preventDefault();
            const variableText = e.dataTransfer.getData('text/plain');
            if (!variableText) return;
            insertIntoTarget(e.currentTarget, variableText);
        }
        function insertVariable(variableText) {
            const target = lastFocusedTarget || document.getElementById('bodyEditor');
            insertIntoTarget(target, variableText);
        }
        function insertIntoTarget(target, textToInsert) {
            target.focus();
            if (target.isContentEditable) {
                document.execCommand('insertText', false, textToInsert);
            } else {
                const start = target.selectionStart || 0;
                const end = target.selectionEnd || 0;
                target.value = target.value.substring(0, start) + textToInsert + target.value.substring(end);
                target.selectionStart = target.selectionEnd = start + textToInsert.length;
            }
            scheduleTextRender(true);
        }

        function handleSigLogoUpload(event) {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                customSigLogoData = e.target.result;
                document.getElementById('clearCustomLogoBtn').classList.remove('hidden');
                scheduleSigRender(true);
            };
            reader.readAsDataURL(file);
        }
        function clearCustomSigLogo() {
            customSigLogoData = null;
            document.getElementById('clearCustomLogoBtn').classList.add('hidden');
            scheduleSigRender(true);
        }
        function resetDefaultSignature() {
            customSigLogoData = null;
            document.getElementById('clearCustomLogoBtn').classList.add('hidden');
            document.getElementById('sigPasteBox').innerHTML = `
                <p style="margin: 0 0 8px 0; color: #1f4e79; font-weight: bold;">Thanks and regards,</p>
                <p style="margin: 0; color: #595959;">
                    <span style="color: #2e75b6; font-weight: bold;">S.M. Masrakul Islam Medha</span> | Brand Executive, Strategic Brand Management | Beximco Pharmaceuticals PLC. | 19 Dhanmondi R/A, Road #7, Dhaka 1205 | Bangladesh
                </p>
                <p style="margin: 3px 0 0 0; color: #595959;">
                    | <strong>Cell</strong> +88 <a href="tel:+8801760021915" style="color: #954f72; text-decoration: underline;">01760021915</a> | <strong>Fax</strong> <a href="tel:+880258613888" style="color: #954f72; text-decoration: underline;">+880 2 58613888</a> | <a href="mailto:mke@bpl.net" style="color: #0563c1; text-decoration: underline;">mke@bpl.net</a> | <a href="http://www.beximcopharma.com" style="color: #0563c1; text-decoration: underline;">www.beximcopharma.com</a>
                </p>
            `;
            scheduleSigRender(true);
        }
        function getAutoLogoHTML(pastedText) {
            if (customSigLogoData) return `<div style="margin-top: 8px;"><img src="${customSigLogoData}" alt="Signature Logo" style="max-height: 38px; width: auto; display: block;"></div>`;
            const autoDetect = document.getElementById('autoDetectLogo');
            if (!autoDetect || !autoDetect.checked) return '';
            const box = document.getElementById('sigPasteBox');
            if (box && box.querySelector('img')) return '';
            const lower = pastedText.toLowerCase();
            if (lower.includes('beximco') || lower.includes('bpl.net') || lower.includes('medha')) {
                return `<div style="margin-top: 8px;"><img src="${DEFAULT_LOGO_B64}" alt="Beximco Pharma Logo" style="height: 38px; width: auto; display: block;"></div>`;
            }
            return '';
        }
        function generateSignatureHTML(fontFamily, fontSize) {
            const include = document.getElementById('includeSignature');
            if (!include || !include.checked) return '';
            const box = document.getElementById('sigPasteBox');
            if (!box) return '';
            const rawHtml = box.innerHTML.trim();
            const plainText = box.innerText.trim();
            if (!plainText && !box.querySelector('img')) return '';
            return `<br><br><div style="font-family: ${fontFamily}; font-size: ${fontSize}; color: #595959; line-height: 1.45;">${rawHtml}${getAutoLogoHTML(plainText)}</div>`;
        }

        // --- EXCEL UPLOAD VIA DIRECT NATIVE DISK PICKER (ZERO BASE64 LAG) ---
        async function applyParsedWorkbookResult(res) {
            if (!res || res.cancelled) return;
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

        function generateTableHTML(region, fontFamily, tableFontSize) {
            const data = tableData[region];
            if (!data || data.length === 0) return "<p class='text-rose-600 italic'>No data available for this region.</p>";

            const excludeKeys = ['to', 'cc', 'isredrow', 'cc: zsi', 'fixed cc', '_styles'];
            const columns = regionHeaders[region] || Object.keys(data[0]).filter(key => !excludeKeys.includes(key.toLowerCase().trim()));
            const mergeEnabled = document.getElementById('mergeGroupCells')?.checked ?? true;
            const showTotalRow = document.getElementById('showTotalRow')?.checked ?? true;
            const includeNotes = document.getElementById('includeNoteRows')?.checked ?? true;
            const retainDesign = document.getElementById('retainExcelDesign')?.checked ?? false;

            const notes = includeNotes ? (regionNotes[region] || []) : [];
            const flaggedCodes = new Set();
            notes.forEach(n => {
                String(n.detail || '').split(/[,;\s]+/).forEach(tok => { if (tok.trim()) flaggedCodes.add(tok.trim().toUpperCase()); });
            });

            let totals = {};
            let isColumnNumeric = {};
            columns.forEach(col => {
                totals[col] = 0;
                const numericValues = data.map(r => r[col]).filter(v => v !== '' && !isNaN(v));
                let isNum = numericValues.length > 0;
                if (/code|ase|rse|mpe|territory|specialty|speciality|remarks/i.test(col)) isNum = false;
                isColumnNumeric[col] = isNum;
            });

            const cellMap = Array.from({ length: data.length }, () => {
                const obj = {};
                columns.forEach(col => { obj[col] = { rs: 1, cs: 1 }; });
                return obj;
            });

            const explicitMerges = regionExplicitMerges[region] || [];

            if (mergeEnabled) {
                // 1. Apply exact 2D merges from Excel
                for (let r = 0; r < data.length; r++) {
                    const mRow = explicitMerges[r] || {};
                    columns.forEach(col => {
                        if (mRow[col] && mRow[col].covered) {
                            cellMap[r][col] = null;
                        } else if (mRow[col] && (mRow[col].rs > 1 || mRow[col].cs > 1)) {
                            cellMap[r][col] = { rs: mRow[col].rs || 1, cs: mRow[col].cs || 1 };
                        }
                    });
                }

                // 2. Smart horizontal merge per row for repeated status phrases (e.g. "No PB selected")
                for (let r = 0; r < data.length; r++) {
                    let cIdx = 0;
                    while (cIdx < columns.length) {
                        const col = columns[cIdx];
                        if (!cellMap[r][col] || cellMap[r][col].cs > 1 || cellMap[r][col].rs > 1) { cIdx++; continue; }
                        const val = String(data[r][col] ?? '').trim();
                        if (val === '' || !isNaN(val) || /code|ase|rse|mpe|remarks/i.test(col)) { cIdx++; continue; }

                        let cs = 1;
                        while (cIdx + cs < columns.length) {
                            const nextCol = columns[cIdx + cs];
                            if (!cellMap[r][nextCol] || cellMap[r][nextCol].cs > 1 || cellMap[r][nextCol].rs > 1) break;
                            const nextVal = String(data[r][nextCol] ?? '').trim();
                            if (nextVal === val) cs++;
                            else break;
                        }
                        if (cs > 1) {
                            cellMap[r][col].cs = cs;
                            for (let k = 1; k < cs; k++) cellMap[r][columns[cIdx + k]] = null;
                        }
                        cIdx += cs;
                    }
                }

                // 3. Vertical merge strictly for hierarchy columns (RSE, ASE, Specialty)
                const hierarchyCols = new Set(['rse', 'region', 'ase', 'specialty', 'speciality', 'zone', 'territory']);
                columns.forEach((col, cIdx) => {
                    if (!hierarchyCols.has(col.toLowerCase())) return;
                    let r = 0;
                    while (r < data.length) {
                        const cellInfo = cellMap[r][col];
                        if (!cellInfo || cellInfo.rs > 1) { r++; continue; }
                        const val = String(data[r][col] ?? '').trim();
                        if (val === '') { r++; continue; }

                        const parentSig = columns.slice(0, cIdx).map(pCol => String(data[r][pCol] ?? '').trim()).join('|');
                        let span = 1;
                        while (r + span < data.length) {
                            const nextInfo = cellMap[r + span][col];
                            if (!nextInfo || nextInfo.rs > 1) break;
                            const nextVal = String(data[r + span][col] ?? '').trim();
                            const nextParentSig = columns.slice(0, cIdx).map(pCol => String(data[r + span][pCol] ?? '').trim()).join('|');
                            if (nextVal === val && nextParentSig === parentSig) span++;
                            else break;
                        }
                        if (span > 1) {
                            cellMap[r][col].rs = span;
                            for (let k = 1; k < span; k++) {
                                cellMap[r + k][col] = null;
                            }
                        }
                        r += span;
                    }
                });
            }

            const tableCls = retainDesign ? "email-table excel-retained" : "email-table";
            let html = `<table class="${tableCls}" style="font-family: ${fontFamily}; font-size: ${tableFontSize};"><thead><tr>`;
            columns.forEach(col => {
                let thStyle = "";
                if (retainDesign) {
                    const hs = (regionHeaderStyles[region] && regionHeaderStyles[region][col]) ? regionHeaderStyles[region][col] : {};
                    const bg = hs.bg || "#ffffff";
                    const fg = hs.color || "#000000";
                    const al = hs.align || "center";
                    thStyle = `background-color: ${bg}; color: ${fg}; text-align: ${al};`;
                }
                html += `<th style="${thStyle}">${col}</th>`;
            });
            html += `</tr></thead><tbody>`;

            data.forEach((row, rIdx) => {
                let rowClass = (!retainDesign && row.isRedRow) ? "class='bg-red-row'" : "";
                html += `<tr ${rowClass}>`;
                columns.forEach(col => {
                    let val = row[col] !== undefined ? row[col] : '';
                    if (isColumnNumeric[col] && val !== '' && !isNaN(val)) totals[col] += parseFloat(val) || 0;

                    const info = cellMap[rIdx][col];
                    if (mergeEnabled && info === null) return;

                    const rsAttr = (mergeEnabled && info && info.rs > 1) ? ` rowspan="${info.rs}"` : '';
                    const csAttr = (mergeEnabled && info && info.cs > 1) ? ` colspan="${info.cs}"` : '';

                    let cellClass = "";
                    let inlineStyles = [];

                    if (retainDesign) {
                        const cs = (row._styles && row._styles[col]) ? row._styles[col] : {};
                        inlineStyles.push(`background-color: ${cs.bg || '#ffffff'}`);
                        inlineStyles.push(`color: ${cs.color || '#000000'}`);
                        if (cs.bold) inlineStyles.push(`font-weight: bold`);
                        if (cs.italic) inlineStyles.push(`font-style: italic`);
                        if (cs.align) inlineStyles.push(`text-align: ${cs.align}`);
                    } else {
                        if (col.toLowerCase().includes('new shelving') || col.toLowerCase() === 'new') {
                            let num = parseInt(val);
                            if (num > 0) cellClass = "text-green";
                            else if (num <= 0) cellClass = "text-red";
                        }
                        if (/no pb selected/i.test(String(val))) {
                            cellClass = "text-red";
                        }
                        if (/^(mpe|t\.?code)$/i.test(col) && flaggedCodes.has(String(val).trim().toUpperCase())) {
                            cellClass = "text-red";
                        }
                        if (col.toLowerCase() === 'remarks') {
                            inlineStyles.push('text-align: left', 'padding-left: 10px');
                        }
                    }

                    const styleStr = inlineStyles.length > 0 ? ` style="${inlineStyles.join('; ')};"` : '';
                    html += `<td${rsAttr}${csAttr} class="${cellClass}"${styleStr}>${val}</td>`;
                });
                html += `</tr>`;
            });

            if (showTotalRow) {
                html += `<tr class="totals-row">`;
                columns.forEach((col, index) => {
                    if (index === 0) html += `<td>${region} Total</td>`;
                    else if (isColumnNumeric[col]) html += `<td>${totals[col]}</td>`;
                    else html += `<td></td>`;
                });
                html += `</tr>`;
            }

            if (notes.length > 0) {
                const labelSpan = Math.min(3, Math.max(1, Math.floor(columns.length / 2)));
                const detailSpan = Math.max(1, columns.length - labelSpan);
                notes.forEach(n => {
                    html += `<tr class="note-row"><td colspan="${labelSpan}">${n.label}</td><td colspan="${detailSpan}" class="note-value">${n.detail}</td></tr>`;
                });
            }

            html += `</tbody></table>`;
            return html;
        }

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

        function buildEmailStyles(fontFamily, bodyFontSize, tableFontSize) {
            return `
                <style>
                    body { font-family: ${fontFamily}; font-size: ${bodyFontSize}; color: #000000; }
                    table { border-collapse: collapse; width: auto; border: 1px solid #000000; font-family: ${fontFamily}; font-size: ${tableFontSize}; }
                    th { background-color: #e6005c; color: #ffffff; border: 1px solid #000000; padding: 5px 12px; text-align: center; font-weight: bold; white-space: nowrap; }
                    td { border: 1px solid #000000; padding: 4px 10px; text-align: center; background-color: #fdf6e3; vertical-align: middle; white-space: nowrap; }
                    table.excel-retained { border: 1px solid #d1d5db; }
                    table.excel-retained th, table.excel-retained td { border: 1px solid #d1d5db; }
                    .totals-row td { background-color: #ffff00; font-weight: bold; }
                    .note-row td { background-color: #ffff00; font-weight: bold; }
                    .note-row td.note-value { background-color: #ffffff; color: #1f4e79; font-weight: normal; text-align: left; padding-left: 10px; }
                    .text-green { color: #00b050 !important; font-weight: bold; }
                    .text-red { color: #ff0000 !important; font-weight: bold; }
                    .bg-red-row td { background-color: #ff0000 !important; color: #ffffff !important; font-weight: bold; }
                    .bg-red-row td.text-green, .bg-red-row td.text-red { color: #ffffff !important; }
                </style>
            `;
        }

        async function openOutlookDraft() {
            const to = document.getElementById('emailTo').value || '';
            const cc = document.getElementById('emailCc').value || '';
            const subject = document.getElementById('emailSubject').innerText || 'Email';
            const fontFamily = document.getElementById('globalFontFamily')?.value || "Calibri, Arial, sans-serif";
            const bodyFontSize = document.getElementById('bodyFontSize')?.value || "11pt";
            const tableFontSize = document.getElementById('tableFontSize')?.value || "10.5pt";
            const bodyContent = document.getElementById('emailBodyContent').innerHTML;
            const fullHtml = `<!DOCTYPE html><html><head><meta charset="utf-8">${buildEmailStyles(fontFamily, bodyFontSize, tableFontSize)}</head><body>${bodyContent}</body></html>`;

            if (window.pywebview && window.pywebview.api) {
                const res = await window.pywebview.api.open_outlook_draft(to, cc, subject, fullHtml);
                if (res === "OK") showStatus("Opened draft directly in Microsoft Outlook!");
                else showStatus(res, true);
            } else {
                exportEML();
            }
        }

        async function batchOutlookDrafts() {
            if (!window.pywebview || !window.pywebview.api) {
                batchExportZip();
                return;
            }
            const regions = Object.keys(tableData);
            if (regions.length === 0) { showStatus("No data available.", true); return; }
            const dateShort = document.getElementById('shortDate').value || "";
            const dateLong = document.getElementById('longDate').value || "";
            const subjectTemplate = document.getElementById('subjectTemplate').value || "";
            const bodyTemplateHtml = document.getElementById('bodyEditor').innerHTML || "";
            const fontFamily = document.getElementById('globalFontFamily')?.value || "Calibri, Arial, sans-serif";
            const bodyFontSize = document.getElementById('bodyFontSize')?.value || "11pt";
            const tableFontSize = document.getElementById('tableFontSize')?.value || "10.5pt";
            const styles = buildEmailStyles(fontFamily, bodyFontSize, tableFontSize);

            const drafts = [];
            for (const region of regions) {
                if (!tableData[region] || tableData[region].length === 0) continue;
                const routing = routingData[region] || { to: "", cc: "" };
                const finalSubject = subjectTemplate.replace(/\{region\}/gi, region).replace(/\{shortDate\}/gi, dateShort).replace(/\{longDate\}/gi, dateLong);
                const finalBody = bodyTemplateHtml.replace(/\{region\}/gi, region).replace(/\{shortDate\}/gi, dateShort).replace(/\{longDate\}/gi, dateLong);
                const tableHtml = generateTableHTML(region, fontFamily, tableFontSize);
                const sigHtml = generateSignatureHTML(fontFamily, bodyFontSize);
                const fullHtml = `<!DOCTYPE html><html><head><meta charset="utf-8">${styles}</head><body><div style="font-family: ${fontFamily}; font-size: ${bodyFontSize};">${finalBody}<br><br>${tableHtml}${sigHtml}</div></body></html>`;
                drafts.push({ to: routing.to || "", cc: routing.cc || "", subject: finalSubject, html: fullHtml });
            }
            const res = await window.pywebview.api.batch_outlook_drafts(JSON.stringify(drafts));
            if (res.startsWith("OK:")) showStatus(res.slice(3));
            else showStatus(res, true);
        }

        function showFeedback() {
            const feedback = document.getElementById('copyFeedback');
            feedback.style.opacity = '1';
            setTimeout(() => { feedback.style.opacity = '0'; }, 2000);
        }
        function copyText(elementId) {
            const text = document.getElementById(elementId).innerText;
            const ta = document.createElement("textarea");
            ta.value = text;
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
            showFeedback();
        }
        function copyRichText(elementId) {
            const el = document.getElementById(elementId);
            const sel = window.getSelection();
            const range = document.createRange();
            range.selectNodeContents(el);
            sel.removeAllRanges();
            sel.addRange(range);
            document.execCommand('copy');
            sel.removeAllRanges();
            showFeedback();
        }

        async function exportEML() {
            const to = document.getElementById('emailTo').value || '';
            const cc = document.getElementById('emailCc').value || '';
            const subject = document.getElementById('emailSubject').innerText || 'Email';
            const fontFamily = document.getElementById('globalFontFamily')?.value || "Calibri, Arial, sans-serif";
            const bodyFontSize = document.getElementById('bodyFontSize')?.value || "11pt";
            const tableFontSize = document.getElementById('tableFontSize')?.value || "10.5pt";
            const bodyContent = document.getElementById('emailBodyContent').innerHTML;
            const fullHtml = `<!DOCTYPE html><html><head><meta charset="utf-8">${buildEmailStyles(fontFamily, bodyFontSize, tableFontSize)}</head><body>${bodyContent}</body></html>`;
            const emlContent = [`To: ${to}`, `Cc: ${cc}`, `Subject: ${subject}`, `X-Unsent: 1`, `Content-Type: text/html; charset="utf-8"`, ``, fullHtml].join('\r\n');
            const filename = `${subject.replace(/[^a-z0-9]/gi, '_').substring(0, 150)}.eml`;

            if (window.pywebview && window.pywebview.api) {
                const res = await window.pywebview.api.save_text_file(filename, emlContent);
                if (res) showStatus("Saved .EML file!");
            } else {
                const blob = new Blob([emlContent], { type: 'message/rfc822' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url; a.download = filename;
                document.body.appendChild(a); a.click(); document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }
        }

        async function batchExportZip() {
            if (Object.keys(tableData).length === 0) { showStatus("No data available to export.", true); return; }
            const zip = new JSZip();
            const dateShort = document.getElementById('shortDate').value || "";
            const dateLong = document.getElementById('longDate').value || "";
            const subjectTemplate = document.getElementById('subjectTemplate').value || "";
            const bodyTemplateHtml = document.getElementById('bodyEditor').innerHTML || "";
            const fontFamily = document.getElementById('globalFontFamily')?.value || "Calibri, Arial, sans-serif";
            const bodyFontSize = document.getElementById('bodyFontSize')?.value || "11pt";
            const tableFontSize = document.getElementById('tableFontSize')?.value || "10.5pt";
            const emailStyles = buildEmailStyles(fontFamily, bodyFontSize, tableFontSize);

            for (const region in tableData) {
                const data = tableData[region];
                if (!data || data.length === 0) continue;
                const routing = routingData[region] || { to: "", cc: "" };
                const finalSubject = subjectTemplate.replace(/\{region\}/gi, region).replace(/\{shortDate\}/gi, dateShort).replace(/\{longDate\}/gi, dateLong);
                const finalBody = bodyTemplateHtml.replace(/\{region\}/gi, region).replace(/\{shortDate\}/gi, dateShort).replace(/\{longDate\}/gi, dateLong);
                const tableHtml = generateTableHTML(region, fontFamily, tableFontSize);
                const sigHtml = generateSignatureHTML(fontFamily, bodyFontSize);
                const fullHtml = `<!DOCTYPE html><html><head><meta charset="utf-8">${emailStyles}</head><body><div style="font-family: ${fontFamily}; font-size: ${bodyFontSize};">${finalBody}<br><br>${tableHtml}${sigHtml}</div></body></html>`;
                const emlContent = [`To: ${routing.to || ''}`, `Cc: ${routing.cc || ''}`, `Subject: ${finalSubject}`, `X-Unsent: 1`, `Content-Type: text/html; charset="utf-8"`, ``, fullHtml].join('\r\n');
                zip.file(`${finalSubject.replace(/[^a-z0-9]/gi, '_').substring(0, 150)}_${region}.eml`, emlContent);
            }

            const zipName = `Batch_Emails_${dateShort.replace(/\./g, '_')}.zip`;
            if (window.pywebview && window.pywebview.api) {
                const b64 = await zip.generateAsync({ type: "base64" });
                const res = await window.pywebview.api.save_base64_file(zipName, b64);
                if (res) showStatus("Saved Batch ZIP file!");
            } else {
                const content = await zip.generateAsync({ type: "blob" });
                const url = URL.createObjectURL(content);
                const a = document.createElement('a');
                a.href = url; a.download = zipName;
                document.body.appendChild(a); a.click(); document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }
        }

        function initializeDates() {
            const today = new Date();
            const dd = String(today.getDate()).padStart(2, '0');
            const mm = String(today.getMonth() + 1).padStart(2, '0');
            const yy = String(today.getFullYear()).slice(-2);
            document.getElementById('shortDate').value = `${dd}.${mm}.${yy}`;
            const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
            document.getElementById('longDate').value = `${today.getDate()} ${months[today.getMonth()]} ${today.getFullYear()}`;
        }

        async function bootApp() {
            initializeDates();
            await loadMemoryFromDisk();
            renderEmail();
        }

        window.addEventListener('pywebviewready', bootApp);
        window.onload = function() {
            setTimeout(() => { if (!window.pywebview) bootApp(); }, 150);
        };
    </script>
</body>
</html>
"""


# ==============================================================================
# PYWEBVIEW DESKTOP BRIDGE API
# ==============================================================================
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
            file_types=("Excel Files (*.xlsx;*.xlsm;*.xls)", "All files (*.*)"),
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
            file_types=("Excel Files (*.xlsx;*.xlsm)", "All files (*.*)"),
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
        if not self.window:
            return False
        result = self.window.create_file_dialog(
            webview.SAVE_DIALOG, save_filename=default_name
        )
        if result:
            target = result[0] if isinstance(result, (list, tuple)) else result
            Path(target).write_text(text_content, encoding="utf-8")
            return True
        return False

    def save_base64_file(self, default_name, b64_data):
        if not self.window:
            return False
        result = self.window.create_file_dialog(
            webview.SAVE_DIALOG, save_filename=default_name
        )
        if result:
            target = result[0] if isinstance(result, (list, tuple)) else result
            Path(target).write_bytes(base64.b64decode(b64_data))
            return True
        return False

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


def main():
    api = DesktopApi()
    window = webview.create_window(
        "Outlook Email Automator - Windows Desktop Edition",
        html=HTML_UI,
        js_api=api,
        width=1380,
        height=880,
        min_size=(1024, 680),
    )
    api.window = window
    webview.start()


if __name__ == "__main__":
    main()
