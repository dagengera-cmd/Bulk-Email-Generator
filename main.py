import base64
import json
import re
from pathlib import Path
import webview

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import win32com.client
    import pythoncom
    HAS_OUTLOOK = True
except ImportError:
    HAS_OUTLOOK = False

PROFILE_DIR = Path.home() / ".outlook_email_automator"
PROFILE_FILE = PROFILE_DIR / "profile.json"

# ... existing code ...
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <!-- SheetJS for Demo Routing template export & fallback -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js" crossorigin="anonymous"></script>
    <!-- JSZip for batch zip export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js" crossorigin="anonymous"></script>
    <style>
# ... existing code ...
            let html = `<table class="email-table" style="font-family: ${fontFamily}; font-size: ${tableFontSize};"><thead><tr>`;
            columns.forEach(col => {
                let thStyle = "";
                if (retainDesign) {
                    const hs = (regionHeaderStyles[region] && regionHeaderStyles[region][col]) || {};
                    const bg = hs.bg || "#ffffff";
                    const fg = hs.color || "#000000";
                    thStyle += `background-color: ${bg}; color: ${fg}; `;
                    if (hs.align) thStyle += `text-align: ${hs.align}; `;
                }
                html += `<th style="${thStyle}">${col}</th>`;
            });
            html += `</tr></thead><tbody>`;

            data.forEach((row, rIdx) => {
                let rowClass = row.isRedRow ? "class='bg-red-row'" : "";
                html += `<tr ${rowClass}>`;
                columns.forEach(col => {
                    let val = row[col] !== undefined ? row[col] : '';
                    if (isColumnNumeric[col] && val !== '') totals[col] += parseFloat(val) || 0;

                    const info = cellMap[rIdx][col];
                    if (mergeEnabled && info === null) return;

                    const rsAttr = (mergeEnabled && info && info.rs > 1) ? ` rowspan="${info.rs}"` : '';
                    const csAttr = (mergeEnabled && info && info.cs > 1) ? ` colspan="${info.cs}"` : '';

                    let cellClass = "";
                    let inlineStyles = [];

                    if (retainDesign) {
                        const cs = (row._styles && row._styles[col]) || {};
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
                        if (/no\s*pb\s*selected/i.test(String(val))) {
                            cellClass = "text-red";
                        }
                        if (/^(mpe|t\.?code)$/i.test(col) && flaggedCodes.has(String(val).trim().toUpperCase())) {
                            cellClass = "text-red";
                        }
                    }

                    if (col.toLowerCase() === 'remarks') {
                        inlineStyles.push('text-align: left', 'padding-left: 10px');
                    }

                    const styleStr = inlineStyles.length > 0 ? ` style="${inlineStyles.join('; ')};"` : '';
                    html += `<td${rsAttr}${csAttr} class="${cellClass}"${styleStr}>${val}</td>`;
                });
                html += `</tr>`;
            });
# ... existing code ...
        // --- EXCEL PARSER POWERED BY PYTHON OPENPYXL (EXACT STYLES & 2D MERGES) ---
        async function handleExcelUpload(event) {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = async function(e) {
                const bytes = new Uint8Array(e.target.result);
                let binary = '';
                for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
                const b64 = btoa(binary);

                if (window.pywebview && window.pywebview.api) {
                    const resJson = await window.pywebview.api.parse_excel_b64(b64, "");
                    applyBackendParseResult(JSON.parse(resJson), b64);
                } else {
                    showStatus("Excel backend parser requires the Windows desktop app.", true);
                }
            };
            reader.readAsArrayBuffer(file);
        }

        let lastUploadedExcelB64 = "";

        async function handleSheetChange() {
            if (!lastUploadedExcelB64 || !window.pywebview || !window.pywebview.api) return;
            const name = document.getElementById('sheetSelect').value;
            const resJson = await window.pywebview.api.parse_excel_b64(lastUploadedExcelB64, name);
            applyBackendParseResult(JSON.parse(resJson), lastUploadedExcelB64);
        }

        function applyBackendParseResult(res, b64) {
            if (res.error) {
                showStatus(res.error, true);
                return;
            }
            lastUploadedExcelB64 = b64;
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
                Object.keys(res.extractedRouting).forEach(reg => {
                    if (!routingData[reg]) routingData[reg] = {};
                    if (res.extractedRouting[reg].to) routingData[reg].to = res.extractedRouting[reg].to;
                    if (res.extractedRouting[reg].cc) routingData[reg].cc = res.extractedRouting[reg].cc;
                });
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
                document.getElementById('tableBubbleText').innerText = `Table (${regions.size} RSE)`;
                showStatus(`Loaded ${regions.length} region(s) with exact Excel styles & 2D merges!`);
            } else {
                showStatus("No valid region rows found in this sheet.", true);
            }
        }
# ... existing code ...
                // 3. Smart vertical merge ONLY for hierarchy group columns (RSE, ASE, Specialty, Zone, Territory)
                const groupHierarchyCols = new Set(['rse', 'region', 'ase', 'specialty', 'speciality', 'zone', 'territory']);
                columns.forEach((col, cIdx) => {
                    if (!groupHierarchyCols.has(col.toLowerCase().trim())) return;
                    let r = 0;
                    while (r < data.length) {
                        const cellInfo = cellMap[r][col];
                        if (!cellInfo || cellInfo.rs > 1) { r++; continue; }
                        const val = String(data[r][col] ?? '').trim();
                        if (val === '' || !isNaN(val)) { r++; continue; }

                        const curCs = cellInfo.cs;
                        const parentSig = columns.slice(0, cIdx).map(pCol => String(data[r][pCol] ?? '').trim()).join('|');
                        let span = 1;
                        while (r + span < data.length) {
                            const nextInfo = cellMap[r + span][col];
                            if (!nextInfo || nextInfo.rs > 1 || nextInfo.cs !== curCs) break;
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
# ... existing code ...
def _excel_color_to_hex(color_obj):
    if not color_obj:
        return None
    ctype = getattr(color_obj, "type", None)
    if ctype == "rgb" and color_obj.rgb:
        rgb = str(color_obj.rgb).strip()
        if len(rgb) == 8:
            if rgb.startswith("00"):
                return None
            rgb = rgb[2:]
        if re.match(r"^[0-9A-Fa-f]{6}$", rgb):
            return f"#{rgb}"
    elif ctype == "theme":
        theme_palette = [
            "#ffffff", "#000000", "#e7e6e6", "#44546a", "#4472c4",
            "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47"
        ]
        idx = getattr(color_obj, "theme", None)
        if isinstance(idx, int) and 0 <= idx < len(theme_palette):
            base = theme_palette[idx]
            tint = getattr(color_obj, "tint", 0.0) or 0.0
            if tint != 0.0:
                r = int(base[1:3], 16)
                g = int(base[3:5], 16)
                b = int(base[5:7], 16)
                def adj(c):
                    return round(c + (255 - c) * tint) if tint > 0 else round(c * (1 + tint))
                return "#{:02x}{:02x}{:02x}".format(
                    max(0, min(255, adj(r))),
                    max(0, min(255, adj(g))),
                    max(0, min(255, adj(b)))
                )
            return base
    return None


def _extract_openpyxl_style(cell):
    style = {}
    fill = getattr(cell, "fill", None)
    if fill and getattr(fill, "fill_type", None) not in (None, "none"):
        bg = _excel_color_to_hex(getattr(fill, "fgColor", None))
        if bg:
            style["bg"] = bg
    font = getattr(cell, "font", None)
    if font:
        fg = _excel_color_to_hex(getattr(font, "color", None))
        if fg:
            style["color"] = fg
        if getattr(font, "bold", False):
            style["bold"] = True
        if getattr(font, "italic", False):
            style["italic"] = True
    align = getattr(cell, "alignment", None)
    if align and getattr(align, "horizontal", None):
        style["align"] = align.horizontal
    return style if style else None


class DesktopApi:
    def __init__(self):
        self.window = None

    def parse_excel_b64(self, b64_data, target_sheet=""):
        if not HAS_OPENPYXL:
            return json.dumps({"error": "openpyxl is not installed."})
        try:
            import io
            raw_bytes = base64.b64decode(b64_data)
            wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
            sheet_names = wb.sheetnames
            active_sheet = target_sheet if (target_sheet and target_sheet in sheet_names) else sheet_names[0]
            ws = wb[active_sheet]

            max_r = ws.max_row or 0
            max_c = ws.max_column or 0
            if max_r == 0 or max_c == 0:
                return json.dumps({"error": "The selected worksheet is empty."})

            grid = []
            style_grid = []
            for r in range(1, max_r + 1):
                row_vals = []
                s_vals = []
                for c in range(1, max_c + 1):
                    cell = ws.cell(row=r, column=c)
                    val = "" if cell.value is None else str(cell.value).strip()
                    row_vals.append(val)
                    s_vals.append(_extract_openpyxl_style(cell))
                grid.append(row_vals)
                style_grid.append(s_vals)

            merge_master = {}
            merge_covered = {}
            for m in ws.merged_cells.ranges:
                min_col, min_row, max_col, max_row = m.bounds
                top_r, top_c = min_row - 1, min_col - 1
                rs = max_row - min_row + 1
                cs = max_col - min_col + 1
                top_val = grid[top_r][top_c]
                top_style = style_grid[top_r][top_c]
                if rs > 1 or cs > 1:
                    merge_master[f"{top_r}:{top_c}"] = {"rs": rs, "cs": cs}
                for r in range(top_r, max_row):
                    for c in range(top_c, max_col):
                        if not (r == top_r and c == top_c) and (rs > 1 or cs > 1):
                            merge_covered[f"{r}:{c}"] = True
                        if grid[r][c] == "" and top_val != "":
                            grid[r][c] = top_val
                        if not style_grid[r][c] and top_style:
                            style_grid[r][c] = top_style

            header_row_idx = -1
            header_cols = []
            for r, row in enumerate(grid):
                rse_idx = next((i for i, v in enumerate(row) if re.match(r"^(rse|region)$", v, re.I)), -1)
                if rse_idx != -1 and len([v for v in row if v != ""]) >= 2:
                    header_row_idx = r
                    header_cols = [{"name": v, "idx": i} for i, v in enumerate(row) if v != ""]
                    break

            if header_row_idx == -1:
                return json.dumps({"error": "Could not find an 'RSE' or 'Region' header column in this sheet."})

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

                repeated_matches = sum(
                    1 for h in header_cols
                    if row[h["idx"]] and row[h["idx"]].lower() == h["name"].lower()
                )
                if repeated_matches >= 2:
                    header_cols = [{"name": v, "idx": i} for i, v in enumerate(row) if v != ""]
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
                    c_idx = h["idx"]
                    name = h["name"]
                    row_obj[name] = row[c_idx]
                    if style_grid[r][c_idx]:
                        row_styles[name] = style_grid[r][c_idx]
                    if f"{r}:{c_idx}" in merge_master:
                        row_merge_info[name] = merge_master[f"{r}:{c_idx}"]
                    elif f"{r}:{c_idx}" in merge_covered:
                        row_merge_info[name] = {"covered": True}

                rse_col = next((h for h in header_cols if re.match(r"^(rse|region)$", h["name"], re.I)), None)
                reg_val = row_obj[rse_col["name"]] if rse_col else ""
                if not reg_val:
                    code_col = next((h for h in header_cols if re.match(r"^(ase|mpe|t\.?code|territory)$", h["name"], re.I)), None)
                    code_val = row_obj[code_col["name"]] if code_col else ""
                    m_pref = re.match(r"^([A-Za-z]+\d)", str(code_val))
                    if m_pref:
                        reg_val = m_pref.group(1).upper()
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
                        if style_grid[header_row_idx][h["idx"]]:
                            region_header_styles[current_region][h["name"]] = style_grid[header_row_idx][h["idx"]]

                if row_obj.get("To") or row_obj.get("TO"):
                    extracted_routing.setdefault(current_region, {})["to"] = row_obj.get("To") or row_obj.get("TO")
                cc_list = []
                if row_obj.get("CC: ZSI"):
                    cc_list.append(row_obj["CC: ZSI"])
                if row_obj.get("Fixed CC"):
                    cc_list.append(row_obj["Fixed CC"])
                if cc_list:
                    extracted_routing.setdefault(current_region, {})["cc"] = "; ".join(cc_list)
                elif row_obj.get("CC") or row_obj.get("Cc"):
                    extracted_routing.setdefault(current_region, {})["cc"] = row_obj.get("CC") or row_obj.get("Cc")

                rem_low = str(row_obj.get("Remarks", "")).lower()
                row_obj["isRedRow"] = ("0 steps" in rem_low) or ("immediate attention" in rem_low)
                row_obj["_styles"] = row_styles
                table_data[current_region].append(row_obj)
                region_explicit_merges[current_region].append(row_merge_info)

            return json.dumps({
                "sheetNames": sheet_names,
                "activeSheet": active_sheet,
                "tableData": table_data,
                "regionNotes": region_notes,
                "regionHeaders": region_headers,
                "regionHeaderStyles": region_header_styles,
                "regionExplicitMerges": region_explicit_merges,
                "extractedRouting": extracted_routing
            })
        except Exception as e:
            return json.dumps({"error": f"Excel parse error: {e}"})
# ... existing code ...
