import base64
import json
import os
from pathlib import Path
import webview

try:
    import win32com.client
    import pythoncom
    HAS_OUTLOOK = True
except ImportError:
    HAS_OUTLOOK = False

PROFILE_DIR = Path.home() / ".outlook_email_automator"
PROFILE_FILE = PROFILE_DIR / "profile.json"

HTML_UI = r"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Outlook Email Automator - Desktop Edition</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js" crossorigin="anonymous"></script>
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
            border: 1px solid black;
            font-size: 10.5pt;
        }
        .email-table th {
            background-color: #e6005c;
            color: white;
            border: 1px solid black;
            padding: 4px 10px;
            text-align: center;
            font-weight: bold;
        }
        .email-table td {
            border: 1px solid black;
            padding: 4px 10px;
            text-align: center;
            background-color: #fdf6e3;
            vertical-align: middle;
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
        .bg-red-row td { background-color: #ff0000 !important; color: white !important; font-weight: bold; }
        .bg-red-row td.text-green, .bg-red-row td.text-red { color: white !important; }
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
                    <p class="text-xs text-slate-500">Auto-saving routing, templates, fonts, signature &amp; merged tables</p>
                </div>
            </div>

            <div class="flex flex-wrap items-center gap-2">
                <span id="copyFeedback" class="text-emerald-600 bg-emerald-50 border border-emerald-200 px-3 py-1 rounded-full text-xs font-medium opacity-0 transition-opacity duration-200">
                    Saved / Copied
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
                            <label for="excelUpload" class="text-[11px] font-bold text-indigo-950">Upload Table Excel (.xlsx)</label>
                            <span class="text-[10px] text-indigo-600 font-medium">Merged &amp; multi-block supported</span>
                        </div>
                        <input type="file" id="excelUpload" accept=".xlsx, .xls" class="block w-full text-[11px] text-slate-600 file:mr-2 file:py-1 file:px-3 file:rounded-full file:border-0 file:text-[11px] file:font-semibold file:bg-indigo-600 file:text-white hover:file:bg-indigo-700 cursor-pointer" onchange="handleExcelUpload(event)">
                        <div id="sheetSelectorRow" class="hidden pt-1 flex items-center justify-between gap-2">
                            <label for="sheetSelect" class="text-[10px] font-semibold text-indigo-900 uppercase">Active Sheet:</label>
                            <select id="sheetSelect" onchange="handleSheetChange()" class="border border-indigo-200 rounded px-2 py-0.5 text-xs bg-white font-semibold text-slate-800"></select>
                        </div>
                    </div>

                    <div id="routingUploadBubble" class="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2 hidden">
                        <div class="flex justify-between items-center">
                            <label for="routingUpload" class="text-[11px] font-bold text-slate-700">Upload Routing (.xlsx)</label>
                            <button onclick="downloadDemoRouting()" class="text-[10px] bg-white border border-slate-200 px-2 py-0.5 rounded-full text-indigo-600 hover:text-indigo-800 font-semibold">
                                Get Demo File
                            </button>
                        </div>
                        <input type="file" id="routingUpload" accept=".xlsx, .xls" class="block w-full text-[11px] text-slate-600 file:mr-2 file:py-1 file:px-3 file:rounded-full file:border-0 file:text-[11px] file:font-semibold file:bg-slate-700 file:text-white hover:file:bg-slate-800 cursor-pointer" onchange="handleRoutingUpload(event)">
                        <p class="text-[10px] text-slate-500">Saved permanently on Windows once uploaded. Columns: <code class="text-slate-700">RSE, To, CC: ZSI, Fixed CC</code></p>
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
                                <input type="text" id="shortDate" class="w-full border border-slate-300 rounded-md px-2 py-1.5 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none" oninput="renderEmail()">
                            </div>
                            <div>
                                <label for="longDate" class="block text-[10px] font-semibold text-slate-500 uppercase mb-0.5">Long Date</label>
                                <input type="text" id="longDate" class="w-full border border-slate-300 rounded-md px-2 py-1.5 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none" oninput="renderEmail()">
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
                                <input type="checkbox" id="retainExcelDesign" onchange="onSettingChange()" class="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5">
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
                        <input type="text" id="subjectTemplate" onfocus="lastFocusedTarget = this" ondragover="allowDrop(event)" ondrop="handleDrop(event)" class="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none" value="RE: Shelving Status of Ventrax HFA Inhaler {shortDate} of {region}" oninput="onSettingChange()">
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
                        <div id="bodyEditor" contenteditable="true" onfocus="lastFocusedTarget = this" ondragover="allowDrop(event)" ondrop="handleDrop(event)" oninput="onSettingChange()" class="w-full min-h-[85px] border border-slate-300 rounded-b-lg px-2.5 py-2 text-xs text-slate-800 bg-white focus:ring-2 focus:ring-indigo-500 focus:outline-none leading-relaxed">Dear Vai,<br><br>Please find below the MPE-wise shelving status of Ventrax HFA Inhaler (VHI) as of {longDate}</div>
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

                        <div id="sigPasteBox" contenteditable="true" oninput="onSettingChange()" class="w-full min-h-[90px] max-h-48 overflow-y-auto border border-slate-300 rounded-b-lg p-2.5 bg-slate-50/50 focus:bg-white focus:ring-2 focus:ring-indigo-500 focus:outline-none text-[11px] leading-relaxed">
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
        let currentWorkbook = null;
        let regionNotes = {};
        let regionHeaders = {};
        let regionHeaderStyles = {};
        let regionExplicitMerges = {};
        let saveTimer = null;

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
            saveTimer = setTimeout(saveMemoryToDisk, 400);
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
            try {
                if (window.pywebview && window.pywebview.api) {
                    await window.pywebview.api.save_profile(JSON.stringify(payload));
                } else {
                    localStorage.setItem('outlook_automator_profile', JSON.stringify(payload));
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
            onSettingChange();
        }
        function applyBodyColor(hex) {
            document.getElementById('colorIndicator').style.textDecorationColor = hex;
            execBodyCmd('foreColor', hex);
        }
        function execSigCmd(cmd, val = null) {
            const editor = document.getElementById('sigPasteBox');
            editor.focus();
            document.execCommand(cmd, false, val);
            onSettingChange();
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
            onSettingChange();
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
            onSettingChange();
        }

        function handleSigLogoUpload(event) {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                customSigLogoData = e.target.result;
                document.getElementById('clearCustomLogoBtn').classList.remove('hidden');
                onSettingChange();
            };
            reader.readAsDataURL(file);
        }
        function clearCustomSigLogo() {
            customSigLogoData = null;
            document.getElementById('clearCustomLogoBtn').classList.add('hidden');
            onSettingChange();
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
            onSettingChange();
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

        function excelColorToHex(colorObj) {
            if (!colorObj) return null;
            if (colorObj.rgb) {
                let rgb = String(colorObj.rgb).trim();
                if (rgb.length === 8) rgb = rgb.slice(2);
                if (/^[0-9A-Fa-f]{6}$/.test(rgb)) return '#' + rgb;
            }
            const themePalette = ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
            if (typeof colorObj.theme === 'number' && themePalette[colorObj.theme]) {
                let base = themePalette[colorObj.theme];
                if (typeof colorObj.tint === 'number' && colorObj.tint !== 0) {
                    const r = parseInt(base.slice(1, 3), 16);
                    const g = parseInt(base.slice(3, 5), 16);
                    const b = parseInt(base.slice(5, 7), 16);
                    const adj = (c) => colorObj.tint > 0
                        ? Math.round(c + (255 - c) * colorObj.tint)
                        : Math.round(c * (1 + colorObj.tint));
                    return '#' + [adj(r), adj(g), adj(b)].map(x => Math.max(0, Math.min(255, x)).toString(16).padStart(2, '0')).join('');
                }
                return base;
            }
            return null;
        }

        function extractCellStyle(cell) {
            if (!cell || !cell.s) return null;
            const s = cell.s;
            const styleObj = {};
            const fill = s.fill || s;
            if (fill && fill.fgColor && (!fill.patternType || fill.patternType !== 'none')) {
                const bg = excelColorToHex(fill.fgColor);
                if (bg) styleObj.bg = bg;
            }
            if (s.font) {
                const fg = excelColorToHex(s.font.color);
                if (fg) styleObj.color = fg;
                if (s.font.bold) styleObj.bold = true;
                if (s.font.italic) styleObj.italic = true;
            }
            if (s.alignment && s.alignment.horizontal) {
                styleObj.align = s.alignment.horizontal;
            }
            return Object.keys(styleObj).length > 0 ? styleObj : null;
        }

        function handleExcelUpload(event) {
            const file = event.target.files[0];
            if (!file) return;
            showStatus("Reading Excel file...");
            const reader = new FileReader();
            reader.onload = function(e) {
                setTimeout(() => {
                    const data = new Uint8Array(e.target.result);
                    currentWorkbook = XLSX.read(data, { type: 'array', cellStyles: true });
                    const sheetSelect = document.getElementById('sheetSelect');
                    sheetSelect.innerHTML = '';
                    currentWorkbook.SheetNames.forEach(name => {
                        const opt = document.createElement('option');
                        opt.value = name;
                        opt.text = name;
                        sheetSelect.appendChild(opt);
                    });
                    document.getElementById('sheetSelectorRow').classList.toggle('hidden', currentWorkbook.SheetNames.length <= 1);
                    parseWorksheet(currentWorkbook.Sheets[currentWorkbook.SheetNames[0]]);
                }, 20);
            };
            reader.readAsArrayBuffer(file);
        }

        function handleSheetChange() {
            if (!currentWorkbook) return;
            const name = document.getElementById('sheetSelect').value;
            parseWorksheet(currentWorkbook.Sheets[name]);
        }

        function parseWorksheet(worksheet) {
            if (!worksheet || !worksheet['!ref']) {
                showStatus("The selected worksheet is empty.", true);
                return;
            }
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            const maxRows = Math.min(range.e.r, range.s.r + 2500);
            const maxCols = Math.min(range.e.c, range.s.c + 30);
            const grid = [];
            const styleGrid = [];
            for (let r = range.s.r; r <= maxRows; r++) {
                const rowArr = [];
                const sRowArr = [];
                for (let c = range.s.c; c <= maxCols; c++) {
                    const cell = worksheet[XLSX.utils.encode_cell({ r, c })];
                    let val = "";
                    if (cell) val = (cell.w !== undefined && cell.w !== null) ? String(cell.w).trim() : (cell.v !== undefined && cell.v !== null ? String(cell.v).trim() : "");
                    rowArr.push(val);
                    sRowArr.push(extractCellStyle(cell));
                }
                grid.push(rowArr);
                styleGrid.push(sRowArr);
            }

            const mergeMaster = {};
            const mergeCovered = {};
            const merges = worksheet['!merges'] || [];
            merges.forEach(m => {
                const topRow = m.s.r - range.s.r;
                const bottomRow = m.e.r - range.s.r;
                const leftCol = m.s.c - range.s.c;
                const rightCol = m.e.c - range.s.c;
                if (topRow < 0 || leftCol < 0 || !grid[topRow]) return;
                const topVal = grid[topRow][leftCol];
                const topStyle = styleGrid[topRow][leftCol];
                const rs = (bottomRow - topRow) + 1;
                const cs = (rightCol - leftCol) + 1;
                if (rs > 1 || cs > 1) {
                    mergeMaster[`${topRow}:${leftCol}`] = { rs, cs };
                }
                for (let r = topRow; r <= bottomRow; r++) {
                    if (!grid[r]) continue;
                    for (let c = leftCol; c <= rightCol; c++) {
                        if (!(r === topRow && c === leftCol) && (rs > 1 || cs > 1)) {
                            mergeCovered[`${r}:${c}`] = true;
                        }
                        if (grid[r][c] === "" && topVal !== "") grid[r][c] = topVal;
                        if (!styleGrid[r][c] && topStyle) styleGrid[r][c] = topStyle;
                    }
                }
            });

            let headerRowIdx = -1;
            let headerCols = [];
            for (let r = 0; r < grid.length; r++) {
                const rseIdx = grid[r].findIndex(v => /^(rse|region)$/i.test(v));
                if (rseIdx !== -1 && grid[r].filter(v => v !== "").length >= 2) {
                    headerRowIdx = r;
                    headerCols = grid[r].map((v, idx) => ({ name: v, idx })).filter(h => h.name !== "");
                    break;
                }
            }
            if (headerRowIdx === -1) {
                showStatus("Could not find an 'RSE' or 'Region' header column in this sheet.", true);
                return;
            }

            tableData = {};
            regionNotes = {};
            regionHeaders = {};
            regionHeaderStyles = {};
            regionExplicitMerges = {};
            const regions = new Set();
            let currentRegion = "";
            let lastGroupVals = {};
            const groupColNames = new Set(['rse', 'region', 'ase', 'specialty', 'speciality', 'zone', 'territory']);
            const routingColNames = new Set(['to', 'cc', 'cc: zsi', 'fixed cc']);

            for (let r = headerRowIdx + 1; r < grid.length; r++) {
                const row = grid[r];
                const nonEmpty = row.filter(v => v !== "");
                if (nonEmpty.length === 0) continue;

                const isRepeatedHeader = headerCols.filter(h => row[h.idx] && row[h.idx].toLowerCase() === h.name.toLowerCase()).length >= 2;
                if (isRepeatedHeader) {
                    headerCols = grid[r].map((v, idx) => ({ name: v, idx })).filter(h => h.name !== "");
                    lastGroupVals = {};
                    continue;
                }

                const firstTwoText = nonEmpty.slice(0, 2).join(" ").toLowerCase();
                if (/\btotal\b/.test(firstTwoText) && !nonEmpty.some(v => /^[A-Z]\d{3}$/i.test(v))) continue;

                const noteLabelIdx = row.findIndex(v => /did not select|not selected|note\s*:/i.test(v));
                if (noteLabelIdx !== -1) {
                    const label = row[noteLabelIdx];
                    const detail = row.slice(noteLabelIdx + 1).filter(v => v !== "").join(", ");
                    if (currentRegion) {
                        if (!regionNotes[currentRegion]) regionNotes[currentRegion] = [];
                        regionNotes[currentRegion].push({ label, detail });
                    }
                    continue;
                }

                const rowObj = {};
                const rowStyles = {};
                const rowMergeInfo = {};
                headerCols.forEach(h => {
                    rowObj[h.name] = row[h.idx] !== undefined ? row[h.idx] : "";
                    if (styleGrid[r][h.idx]) rowStyles[h.name] = styleGrid[r][h.idx];
                    if (mergeMaster[`${r}:${h.idx}`]) rowMergeInfo[h.name] = mergeMaster[`${r}:${h.idx}`];
                    else if (mergeCovered[`${r}:${h.idx}`]) rowMergeInfo[h.name] = { covered: true };
                });

                const rseCol = headerCols.find(h => /^(rse|region)$/i.test(h.name));
                let regVal = rseCol ? rowObj[rseCol.name] : "";
                if (!regVal) {
                    const codeCol = headerCols.find(h => /^(ase|mpe|t\.?code|territory)$/i.test(h.name));
                    const codeVal = codeCol ? rowObj[codeCol.name] : "";
                    const prefixMatch = String(codeVal).match(/^([A-Za-z]+\d)/);
                    if (prefixMatch) regVal = prefixMatch[1].toUpperCase();
                    else if (currentRegion) regVal = currentRegion;
                }
                if (!regVal) continue;

                if (regVal !== currentRegion) {
                    currentRegion = regVal;
                    lastGroupVals = {};
                }
                if (rseCol) rowObj[rseCol.name] = currentRegion;

                headerCols.forEach(h => {
                    const low = h.name.toLowerCase();
                    if (groupColNames.has(low)) {
                        if (rowObj[h.name] !== "") lastGroupVals[low] = rowObj[h.name];
                        else if (lastGroupVals[low]) rowObj[h.name] = lastGroupVals[low];
                    }
                });

                regions.add(currentRegion);
                if (!tableData[currentRegion]) {
                    tableData[currentRegion] = [];
                    regionHeaders[currentRegion] = headerCols.filter(h => !routingColNames.has(h.name.toLowerCase())).map(h => h.name);
                    regionHeaderStyles[currentRegion] = {};
                    regionExplicitMerges[currentRegion] = [];
                    headerCols.forEach(h => {
                        if (styleGrid[headerRowIdx][h.idx]) regionHeaderStyles[currentRegion][h.name] = styleGrid[headerRowIdx][h.idx];
                    });
                }

                if (rowObj['To'] || rowObj['TO']) {
                    if (!routingData[currentRegion]) routingData[currentRegion] = {};
                    routingData[currentRegion].to = rowObj['To'] || rowObj['TO'];
                }
                let combinedCC = [];
                if (rowObj['CC: ZSI']) combinedCC.push(rowObj['CC: ZSI']);
                if (rowObj['Fixed CC']) combinedCC.push(rowObj['Fixed CC']);
                if (combinedCC.length > 0) {
                    if (!routingData[currentRegion]) routingData[currentRegion] = {};
                    routingData[currentRegion].cc = combinedCC.join('; ');
                } else if (rowObj['CC'] || rowObj['Cc']) {
                    if (!routingData[currentRegion]) routingData[currentRegion] = {};
                    routingData[currentRegion].cc = rowObj['CC'] || rowObj['Cc'];
                }

                rowObj.isRedRow = String(rowObj['Remarks'] || '').toLowerCase().includes('0 steps') || String(rowObj['Remarks'] || '').toLowerCase().includes('immediate attention');
                rowObj._styles = rowStyles;
                tableData[currentRegion].push(rowObj);
                regionExplicitMerges[currentRegion].push(rowMergeInfo);
            }

            const select = document.getElementById('regionSelect');
            select.innerHTML = '';
            regions.forEach(region => {
                const option = document.createElement('option');
                option.value = region;
                option.text = region;
                select.appendChild(option);
            });

            if (regions.size > 0) {
                renderEmail();
                scheduleSaveMemory();
                document.getElementById('tableBubbleText').innerText = `Table (${regions.size} RSE)`;
                showStatus(`Loaded ${regions.size} region(s) with 2D merged-cell & Excel design support!`);
            } else {
                showStatus("No valid region rows found in this sheet.", true);
            }
        }

        function handleRoutingUpload(event) {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
                let count = 0;
                json.forEach(row => {
                    const region = row['RSE'] || row['Region'];
                    if (!region) return;
                    if (!routingData[region]) routingData[region] = {};
                    if (row['To'] || row['TO']) routingData[region].to = row['To'] || row['TO'];
                    let combinedCC = [];
                    if (row['CC: ZSI']) combinedCC.push(row['CC: ZSI']);
                    if (row['Fixed CC']) combinedCC.push(row['Fixed CC']);
                    if (combinedCC.length > 0) routingData[region].cc = combinedCC.join('; ');
                    else if (row['CC'] || row['Cc'] || row['cc']) routingData[region].cc = row['CC'] || row['Cc'] || row['cc'];
                    count++;
                });
                renderEmail();
                saveMemoryToDisk();
                if (count > 0) {
                    document.getElementById('routingBubbleText').innerText = `Routing (${count})`;
                    document.getElementById('routingUploadBubble').classList.add('hidden');
                    showStatus(`Saved routing (To/CC) for ${count} regions to Windows memory!`);
                }
            };
            reader.readAsArrayBuffer(file);
        }

        async function downloadDemoRouting() {
            const ws_data = [
                ["RSE", "To", "CC: ZSI", "Fixed CC"],
                ["R4", "PSL-(NAL)-Md. Nur-A- Alam", "SBM-(ALM)-Mahabub Alam.", "PSL-(KNH)-Kazi Nurul Hoque; PSL-(MDR)-Masud Rana khan"],
                ["A2", "PSL-(MZO)-Md. Zakir Hossain", "SBM-(ALM)-Mahabub Alam.", "PSL-(KNH)-Kazi Nurul Hoque; PSL-(MDR)-Masud Rana khan"],
                ["N4", "PSL-(KMO)-Kazi Md. Mostafa", "SBM-(RML)-Noor-Ul Azam Khan Romel", "PSL-(KNH)-Kazi Nurul Hoque"]
            ];
            const ws = XLSX.utils.aoa_to_sheet(ws_data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "RoutingTemplate");
            if (window.pywebview && window.pywebview.api) {
                const b64 = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
                const res = await window.pywebview.api.save_base64_file("Demo_Routing_Template.xlsx", b64);
                if (res) showStatus("Saved Demo_Routing_Template.xlsx!");
            } else {
                XLSX.writeFile(wb, "Demo_Routing_Template.xlsx");
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
                const textValues = data.map(r => String(r[col] ?? '').trim()).filter(v => v !== '' && isNaN(v));
                let isNum = numericValues.length > 0;
                if (/code|ase|rse|mpe|territory|specialty|speciality/i.test(col)) isNum = false;
                isColumnNumeric[col] = isNum;
            });

            const cellMap = Array.from({ length: data.length }, () => {
                const obj = {};
                columns.forEach(col => { obj[col] = { rs: 1, cs: 1 }; });
                return obj;
            });

            if (mergeEnabled) {
                // 1. Horizontal merge per row for repeated non-numeric status phrases (e.g. "No PB selected")
                for (let r = 0; r < data.length; r++) {
                    let cIdx = 0;
                    while (cIdx < columns.length) {
                        const col = columns[cIdx];
                        const val = String(data[r][col] ?? '').trim();
                        if (val === '' || !isNaN(val) || /code|ase|rse|mpe|remarks/i.test(col)) { cIdx++; continue; }

                        let cs = 1;
                        while (cIdx + cs < columns.length) {
                            const nextCol = columns[cIdx + cs];
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

                // 2. Vertical merge strictly for hierarchy columns (RSE, ASE, Specialty)
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

            let html = `<table class="email-table" style="font-family: ${fontFamily}; font-size: ${tableFontSize};"><thead><tr>`;
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
                let rowClass = row.isRedRow ? "class='bg-red-row'" : "";
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
                        inlineStyles.push(`color: ${cs.color || (/no pb selected/i.test(String(val)) ? '#ff0000' : '#000000')}`);
                        if (cs.bold || /no pb selected/i.test(String(val))) inlineStyles.push(`font-weight: bold`);
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
                    }

                    if (col.toLowerCase() === 'remarks') {
                        inlineStyles.push('text-align: left', 'padding-left: 10px');
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

        function renderEmail() {
            const region = document.getElementById('regionSelect').value || "N/A";
            const dateShort = document.getElementById('shortDate').value || "";
            const dateLong = document.getElementById('longDate').value || "";
            const fontFamily = document.getElementById('globalFontFamily')?.value || "Calibri, Arial, sans-serif";
            const bodyFontSize = document.getElementById('bodyFontSize')?.value || "11pt";
            const tableFontSize = document.getElementById('tableFontSize')?.value || "10.5pt";

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

            const routing = routingData[region] || { to: "", cc: "" };
            document.getElementById('emailTo').value = routing.to || "";
            document.getElementById('emailCc').value = routing.cc || "";

            document.getElementById('tableContainer').innerHTML = generateTableHTML(region, fontFamily, tableFontSize);
            document.getElementById('signatureContainer').innerHTML = generateSignatureHTML(fontFamily, bodyFontSize);
        }

        function buildEmailStyles(fontFamily, bodyFontSize, tableFontSize) {
            return `
                <style>
                    body { font-family: ${fontFamily}; font-size: ${bodyFontSize}; color: #000000; }
                    table { border-collapse: collapse; width: auto; border: 1px solid black; font-family: ${fontFamily}; font-size: ${tableFontSize}; }
                    th { background-color: #e6005c; color: white; border: 1px solid black; padding: 4px 10px; text-align: center; font-weight: bold; }
                    td { border: 1px solid black; padding: 4px 10px; text-align: center; background-color: #fdf6e3; vertical-align: middle; }
                    .totals-row td { background-color: #ffff00; font-weight: bold; }
                    .note-row td { background-color: #ffff00; font-weight: bold; }
                    .note-row td.note-value { background-color: #ffffff; color: #1f4e79; font-weight: normal; text-align: left; padding-left: 10px; }
                    .text-green { color: #00b050 !important; font-weight: bold; }
                    .text-red { color: #ff0000 !important; font-weight: bold; }
                    .bg-red-row td { background-color: #ff0000 !important; color: white !important; font-weight: bold; }
                    .bg-red-row td.text-green, .bg-red-row td.text-red { color: white !important; }
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


class DesktopApi:
    def __init__(self):
        self.window = None

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
            PROFILE_FILE.write_text(json_str, encoding="utf-8")
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
        if not HAS_OUTLOOK:
            return "pywin32 not available; use .EML export instead."
        try:
            pythoncom.CoInitialize()
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

    def batch_outlook_drafts(self, drafts_json):
        if not HAS_OUTLOOK:
            return "pywin32 not available; use .ZIP export instead."
        try:
            drafts = json.loads(drafts_json)
            pythoncom.CoInitialize()
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
