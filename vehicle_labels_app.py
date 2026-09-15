#!/usr/bin/env python3
"""
Vehicle Labels Application
All-in-one local GUI for converting, viewing, and printing vehicle labels.
No browser required. Matches the HTML template exactly.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont
import os
import sys
import re
import json
import math
import socket
import subprocess
import shutil
import zipfile
import io
import time
import threading
import xml.etree.ElementTree as ET

if getattr(sys, "frozen", False):
    ROOT = os.path.dirname(os.path.abspath(sys.executable))
else:
    ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(os.environ.get("LOCALAPPDATA", ROOT), "VehicleLabeler")
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")
IMPORTS_DIR = os.path.join(ROOT, "Imports")
FORMATTED_DIR = os.path.join(ROOT, "Formatted")
LABELS_DIR = os.path.join(ROOT, "Labels")
TEMPLATE_PATH = os.path.join(ROOT, "Template", "vehicle_labels.html")

for d in [IMPORTS_DIR, FORMATTED_DIR, LABELS_DIR, CONFIG_DIR]:
    os.makedirs(d, exist_ok=True)


def js_round(x):
    return int(math.floor(x + 0.5))


# ---------------------------------------------------------------------------
# Code 128-B barcode encoder (ported from HTML template)
# ---------------------------------------------------------------------------
C128 = [
    "11011001100","11001101100","11001100110","10010011000","10010001100",
    "10001001100","10011001000","10011000100","10001100100","11001001000",
    "11001000100","11000100100","10110011100","10011011100","10011001110",
    "10111001100","10011101100","10011100110","11001110010","11001011100",
    "11001001110","11011100100","11001110100","11101101110","11101001100",
    "11100101100","11100100110","11101100100","11100110100","11100110010",
    "11011011000","11011000110","11000110110","10100011000","10001011000",
    "10001000110","10110001000","10001101000","10001100010","11010001000",
    "11000101000","11000100010","10110111000","10110001110","10001101110",
    "10111011000","10111000110","10001110110","11101110110","11010001110",
    "11000101110","11011101000","11011100010","11011101110","11101011000",
    "11101000110","11100010110","11101101000","11101100010","11100011010",
    "11101111010","11001000010","11110001010","10100110000","10100001100",
    "10010110000","10010000110","10000101100","10000100110","10110010000",
    "10110000100","10011010000","10011000010","10000110100","10000110010",
    "11000010010","11001010000","11110111010","11000010100","10001111010",
    "10100111100","10010111100","10010011110","10111100100","10011110100",
    "10011110010","11110100100","11110010100","11110010010","11011011110",
    "11011110110","11110110110","10101111000","10100011110","10001011110",
    "10111101000","10111100010","11110101000","11110100010","10111011110",
    "10111101110","11101011110","11110101110","11010000100","11010010000",
    "11010011100","11000111010",
]


def code128_bits(text):
    clean = "".join(c if 32 <= ord(c) <= 126 else " " for c in text)
    codes = [104]
    for ch in clean:
        codes.append(ord(ch) - 32)
    total = 104
    for i in range(1, len(codes)):
        total += codes[i] * i
    codes.append(total % 103)
    codes.append(106)
    bits = "".join(C128[c] for c in codes)
    return bits + "11"


# ---------------------------------------------------------------------------
# Config helpers
# ---------------------------------------------------------------------------
def load_config():
    defaults = {"printer_ip": "192.0.2.1", "printer_port": 9100, "dpi": 300,
                "combo": True, "copies": 1, "adj_x": 0.0, "adj_y": 0.35}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                data = json.load(f)
            defaults["printer_ip"] = data.get("printer", {}).get("ip", defaults["printer_ip"])
            defaults["printer_port"] = data.get("printer", {}).get("port", defaults["printer_port"])
            defaults["dpi"] = data.get("dpi", defaults["dpi"])
            defaults["adj_x"] = data.get("adj_x", defaults["adj_x"])
            defaults["adj_y"] = data.get("adj_y", defaults["adj_y"])
        except Exception:
            pass
    return defaults


def save_config(cfg):
    data = {"printer": {"ip": cfg["printer_ip"], "port": cfg["printer_port"],
                         "web": {"port": 8080}},
            "dpi": cfg["dpi"], "adj_x": cfg["adj_x"], "adj_y": cfg["adj_y"]}
    try:
        with open(CONFIG_PATH, "w") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Convert-IN: .xls text report -> formatted .xlsx (ported from PowerShell)
# ---------------------------------------------------------------------------
def convert_report(source_path, dest_path):
    with open(source_path, "r", encoding="utf-8", errors="replace") as f:
        lines = [l.rstrip("\r\n") for l in f.readlines()]

    header_line = None
    for l in lines:
        if "Location" in l and "VIN" in l:
            header_line = l
            break
    if not header_line:
        raise ValueError("Could not find the column header line containing 'VIN'")

    model_pos = header_line.find("Model")
    color_pos = header_line.find("Color")
    if color_pos < 0:
        color_pos = header_line.find("Co ")

    vin_re = re.compile(r'(?<![A-Z0-9])[A-Z0-9]{17}(?![A-Z0-9])')
    ym_re = re.compile(r'(?<![A-Z0-9-])(\d{2})\s+([A-Z]{2,})(?![A-Z0-9-])')
    count_re = re.compile(r'(\d+)\s*$')

    vehicles = []
    for line in lines:
        if not line.strip():
            continue
        m = vin_re.search(line)
        if not m:
            continue

        vin = m.group()
        before = line[:m.start()].strip()
        loc_tokens = [t for t in before.split() if t]

        loc_a = ""
        loc_b = ""
        if len(loc_tokens) >= 4:
            half = len(loc_tokens) // 2
            loc_a = " ".join(loc_tokens[:half])
            loc_b = " ".join(loc_tokens[half:])
        elif len(loc_tokens) == 3:
            loc_a = loc_tokens[0]
            loc_b = " ".join(loc_tokens[1:])
        elif len(loc_tokens) == 2:
            loc_a = loc_tokens[0]
            loc_b = loc_tokens[1]
        elif len(loc_tokens) == 1:
            loc_a = loc_tokens[0]
            loc_b = loc_tokens[0]

        after = line[m.end():]
        ym = ym_re.search(after)
        if not ym:
            continue
        year = ym.group(1)
        make = ym.group(2)

        stock_part = after[:ym.start()].strip()
        stock_tokens = [t for t in stock_part.split() if t]
        stock = ""
        vtype = ""
        if stock_tokens:
            stock = stock_tokens[0]
            i = 1
            while i < len(stock_tokens) and re.match(r'^[A-Za-z]{1,5}$', stock_tokens[i]):
                stock = stock + " " + stock_tokens[i]
                i += 1
            if i < len(stock_tokens):
                vtype = " ".join(stock_tokens[i:])

        tail = after[ym.end():]
        count = ""
        cm = count_re.search(tail)
        if cm:
            count = cm.group(1)
        mid = tail[:len(tail) - (len(cm.group()) if cm else 0)].strip()

        model = ""
        color = ""
        if len(line) > color_pos:
            cz = line[color_pos:]
            czm = count_re.search(cz)
            if czm:
                cz = cz[:czm.start()]
            color = cz.strip()
        if len(line) > model_pos + 1:
            take = min(color_pos - model_pos, len(line) - model_pos)
            if take > 0:
                model = line[model_pos:model_pos + take].strip()
        if not model and mid:
            mid_no = mid.replace(color, "").strip() if color and color in mid else mid
            mt = [t for t in mid_no.split() if t]
            if mt:
                model = mt[0]
        if not color and mid:
            ct = [t for t in mid.split() if t]
            if len(ct) > 1:
                color = " ".join(ct[1:])

        if not vtype:
            vtype = model

        vehicles.append({
            "location": loc_a, "location2": loc_b, "vin": vin, "stock": stock,
            "type": vtype, "yearMake": f"{year} {make}", "color": color, "count": count,
        })

    if not vehicles:
        raise ValueError("No vehicle rows found")

    out_dir = os.path.dirname(dest_path)
    os.makedirs(out_dir, exist_ok=True)

    tmp_path = dest_path + ".tmp.xlsx"
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _content_types_xml())
        zf.writestr("_rels/.rels", _root_rels_xml())
        zf.writestr("xl/workbook.xml", _workbook_xml())
        zf.writestr("xl/_rels/workbook.xml.rels", _wb_rels_xml())
        zf.writestr("xl/worksheets/sheet1.xml", _sheet_xml(vehicles))

    shutil.move(tmp_path, dest_path)
    return len(vehicles)


def _esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _content_types_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '</Types>')


def _root_rels_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')


def _workbook_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Inventory" sheetId="1" r:id="rId1"/></sheets></workbook>')


def _wb_rels_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '</Relationships>')


def _sheet_xml(vehicles):
    headers = ["Location", "Location", "VIN", "Stock #", "Type", "Year Make", "Color", "Co"]
    rows = []
    rows.append('<row r="3"><c r="A3" t="inlineStr"><is><t>Current</t></is></c>'
                '<c r="B3" t="inlineStr"><is><t>Count</t></is></c>'
                '<c r="C3" t="inlineStr"><is><t>Stock</t></is></c></row>')
    row = '<row r="4">'
    for ci, h in enumerate(headers):
        ref = chr(65 + ci) + "4"
        row += f'<c r="{ref}" t="inlineStr"><is><t>{_esc(h)}</t></is></c>'
    rows.append(row + '</row>')

    for ri, v in enumerate(vehicles, start=5):
        vals = [v["location"], v.get("location2", ""), v["vin"], v["stock"],
                v["type"], v["yearMake"], v["color"], v["count"]]
        row = f'<row r="{ri}">'
        for ci, val in enumerate(vals):
            ref = chr(65 + ci) + str(ri)
            row += f'<c r="{ref}" t="inlineStr"><is><t>{_esc(str(val))}</t></is></c>'
        rows.append(row + '</row>')

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetData>' + "".join(rows) + '</sheetData></worksheet>')


def save_formatted_xlsx(vehicles, dest_path):
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    tmp = dest_path + ".tmp.xlsx"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _content_types_xml())
        zf.writestr("_rels/.rels", _root_rels_xml())
        zf.writestr("xl/workbook.xml", _workbook_xml())
        zf.writestr("xl/_rels/workbook.xml.rels", _wb_rels_xml())
        zf.writestr("xl/worksheets/sheet1.xml", _sheet_xml(vehicles))
    shutil.move(tmp, dest_path)


# ---------------------------------------------------------------------------
# Parse formatted .xlsx -> vehicles list (ported from Make-Labels.ps1)
# ---------------------------------------------------------------------------
def parse_xlsx(path):
    NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    fs = open(path, "rb")
    zf = zipfile.ZipFile(fs, "r")

    def read_entry(name):
        for e in zf.infolist():
            if e.filename.lower() == name.lower():
                return zf.read(e).decode("utf-8")
        return None

    strings = []
    ss_text = read_entry("xl/sharedStrings.xml")
    if ss_text:
        ss_doc = ET.fromstring(ss_text)
        for si in ss_doc.findall(".//s:si", NS):
            t_el = si.find(".//s:t", NS)
            strings.append(t_el.text if t_el is not None and t_el.text else "")

    sheet_target = "xl/worksheets/sheet1.xml"
    wb_text = read_entry("xl/workbook.xml")
    if wb_text:
        wb_doc = ET.fromstring(wb_text)
        sheet_el = wb_doc.find(".//s:sheet", NS)
        if sheet_el is not None:
            rid = sheet_el.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") or sheet_el.get("r:id", "")
            rels_text = read_entry("xl/_rels/workbook.xml.rels")
            if rels_text and rid:
                rels_doc = ET.fromstring(rels_text)
                for rel in rels_doc.findall(".//s:Relationship", NS):
                    if rel.get("Id") == rid:
                        t = rel.get("Target", "").lstrip("/")
                        if not t.startswith("xl/"):
                            t = "xl/" + t
                        sheet_target = t
                        break

    sheet_text = read_entry(sheet_target)
    if not sheet_text:
        for e in zf.infolist():
            if e.filename.startswith("xl/worksheets/") and e.filename.endswith(".xml"):
                sheet_text = zf.read(e).decode("utf-8")
                break
    if not sheet_text:
        zf.close()
        fs.close()
        raise ValueError("No worksheet found")

    def col_num(ref):
        n = 0
        for ch in ref:
            if ch.isalpha():
                n = n * 26 + (ord(ch.upper()) - 64)
        return n

    sheet_doc = ET.fromstring(sheet_text)
    rows_raw = []
    for row_el in sheet_doc.findall(".//s:row", NS):
        cells = {}
        for c_el in row_el.findall("s:c", NS):
            ref = c_el.get("r", "")
            if not ref:
                continue
            ref_letters = re.sub(r"\d", "", ref)
            idx = col_num(ref_letters)
            t = c_el.get("t", "")
            val = ""
            if t == "s":
                v_el = c_el.find("s:v", NS)
                if v_el is not None and v_el.text:
                    i = int(v_el.text)
                    if 0 <= i < len(strings):
                        val = strings[i]
            elif t == "inlineStr":
                is_el = c_el.find("s:is", NS)
                if is_el is not None:
                    t_el = is_el.find("s:t", NS)
                    if t_el is not None and t_el.text:
                        val = t_el.text
            else:
                v_el = c_el.find("s:v", NS)
                if v_el is not None and v_el.text:
                    val = v_el.text
            cells[idx] = val.strip()
        rows_raw.append(cells)

    zf.close()
    fs.close()

    header_idx = -1
    for i, cells in enumerate(rows_raw):
        if any(v.upper() == "VIN" for v in cells.values()):
            header_idx = i
            break
    if header_idx < 0:
        raise ValueError('Could not find a header row containing "VIN"')

    vehicles = []
    for i in range(header_idx + 1, len(rows_raw)):
        c = rows_raw[i]
        vin = c.get(3, "").upper()
        if not vin:
            continue
        vehicles.append({
            "location": c.get(1, ""),
            "vin": vin,
            "stock": c.get(4, ""),
            "type": c.get(5, ""),
            "yearMake": c.get(6, ""),
            "color": c.get(7, ""),
            "count": c.get(8, ""),
        })

    if not vehicles:
        raise ValueError("No vehicle rows found below the header")
    return vehicles


# ---------------------------------------------------------------------------
# Layout defaults (matching HTML template exactly)
# ---------------------------------------------------------------------------
LAYOUT_DEFAULTS = {
    "stockY": 0.10, "stockH": 0.60, "tstock": 0.60,
    "barY": 0.80,   "barH": 0.95,
    "vinY": 1.90,   "vinH": 0.20,  "tvin": 0.20,
    "infoY": 2.25,  "infoH": 0.20, "tinfo": 0.20,
    "mstockY": 0.035, "mstockH": 0.079, "tmstock": 0.079,
    "mbarY": 0.140,   "mbarH": 0.403,
    "mvinY": 0.613,   "mvinH": 0.070,  "tmvin": 0.070,
    "miniY": 0.0,
}
W_RATIO = {"stock": 0.583, "vin": 0.65, "info": 0.60, "mstock": 1.22, "mvin": 1.0}

LAYOUT_PATH = os.path.join(CONFIG_DIR, "layout.json")


def load_layout():
    L = dict(LAYOUT_DEFAULTS)
    if os.path.exists(LAYOUT_PATH):
        try:
            with open(LAYOUT_PATH, "r") as f:
                saved = json.load(f)
            L.update(saved)
        except Exception:
            pass
    return L


def save_layout(L):
    try:
        with open(LAYOUT_PATH, "w") as f:
            json.dump(L, f, indent=2)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# ZPL generation (ported from HTML template buildZPL — exact match)
# ---------------------------------------------------------------------------
def zpl_safe(s):
    out = []
    for c in str(s):
        cc = ord(c)
        if cc == 94:
            out.append("_5E")
        elif cc == 126:
            out.append("_7E")
        elif 32 <= cc <= 126:
            out.append(c)
    return "".join(out)


def car_info(v):
    parts = [v.get("yearMake", ""), v.get("type", "")]
    parts = [p for p in parts if p]
    return (" ".join(parts) + " \u2013 " + v.get("color", "")).strip()


def build_zpl(vehicles, dpi=300, combo=True, copies=1, adj_x=0.0, adj_y=0.35, L=None):
    if L is None:
        L = load_layout()
    w_in = 4.0
    h_in = 5.0 if combo else 2.0
    pw = js_round(w_in * dpi)
    ll = js_round(h_in * dpi)
    main_top = 1.75 if combo else 0
    n_x = js_round(adj_x * dpi)
    n_y = js_round(adj_y * dpi)

    DARKNESS = 25
    SPEED = 4

    y_stock = js_round((main_top + L["stockY"]) * dpi)
    h_stock = js_round(L["stockH"] * dpi)
    w_stock = js_round(L["tstock"] * W_RATIO["stock"] * dpi)
    t_stock = js_round(L["tstock"] * dpi)
    y_bar   = js_round((main_top + L["barY"]) * dpi)
    h_bar   = js_round(L["barH"] * dpi)
    y_vin   = js_round((main_top + L["vinY"]) * dpi)
    h_vin   = js_round(L["vinH"] * dpi)
    w_vin   = js_round(L["tvin"] * W_RATIO["vin"] * dpi)
    t_vin   = js_round(L["tvin"] * dpi)
    y_info  = js_round((main_top + L["infoY"]) * dpi)
    h_info  = js_round(L["infoH"] * dpi)
    w_info  = js_round(L["tinfo"] * W_RATIO["info"] * dpi)
    t_info  = js_round(L["tinfo"] * dpi)

    out = "~SD" + str(DARKNESS) + "\r\n"
    fine_count = 0

    for v in vehicles:
        payload = v["vin"]
        modules = 11 * (len(payload) + 2) + 13
        m = math.floor((pw * 0.94) / modules)
        if m < 1: m = 1
        if m > 4: m = 4
        if m < 2: fine_count += 1
        bw = modules * m
        x_bar = max(0, js_round((pw - bw) / 2))

        out += "^XA^CI28^PW" + str(pw) + "^LL" + str(ll) + "^LH" + str(n_x) + "," + str(n_y) + "^PR" + str(SPEED) + "\r\n"
        out += "^FO0," + str(y_stock) + "^FB" + str(pw) + ",1,0,C,0^A0N," + str(t_stock) + "," + str(w_stock) + "^FD" + zpl_safe(v["stock"]) + "^FS\r\n"
        out += "^FO" + str(x_bar) + "," + str(y_bar) + "^BY" + str(m) + ",3," + str(h_bar) + "^BCN," + str(h_bar) + ",N,N,N^FD" + zpl_safe(payload) + "^FS\r\n"
        out += "^FO0," + str(y_vin) + "^FB" + str(pw) + ",1,0,C,0^A0N," + str(t_vin) + "," + str(w_vin) + "^FD" + zpl_safe(v["vin"]) + "^FS\r\n"
        out += "^FO0," + str(y_info) + "^FB" + str(pw) + ",1,0,C,0^A0N," + str(t_info) + "," + str(w_info) + "^FD" + zpl_safe(car_info(v)) + "^FS\r\n"

        if combo:
            mw = js_round(2 * dpi)
            mh = js_round(0.875 * dpi)
            tMstock = js_round(L["tmstock"] * dpi)
            mWstock = js_round(L["tmstock"] * W_RATIO["mstock"] * dpi)
            mHbar   = js_round(L["mbarH"] * dpi)
            tMvin   = js_round(L["tmvin"] * dpi)
            mWvin   = js_round(L["tmvin"] * W_RATIO["mvin"] * dpi)
            for ix, iy in [(0, 0), (2, 0), (0, 0.875), (2, 0.875)]:
                x0 = ix * dpi
                y0 = (iy + L["miniY"]) * dpi
                out += "^FO" + str(x0) + "," + str(y0 + js_round(L["mstockY"] * dpi)) + "^FB" + str(mw) + ",1,0,C,0^A0N," + str(tMstock) + "," + str(mWstock) + "^FD" + zpl_safe(v["stock"]) + "^FS\r\n"

                vin_modules = 11 * (len(v["vin"]) + 2) + 13
                mm = math.floor((mw * 0.88) / vin_modules)
                if mm < 1: mm = 1
                if mm > 3: mm = 3
                vinBW = vin_modules * mm
                xVinBar = x0 + max(0, js_round((mw - vinBW) / 2))
                yVinBar = y0 + js_round(L["mbarY"] * dpi)
                out += "^FO" + str(xVinBar) + "," + str(yVinBar) + "^BY" + str(mm) + ",2," + str(mHbar) + "^BCN," + str(mHbar) + ",N,N,N^FD" + zpl_safe(v["vin"]) + "^FS\r\n"

                out += "^FO" + str(x0) + "," + str(y0 + js_round(L["mvinY"] * dpi)) + "^FB" + str(mw) + ",1,0,C,0^A0N," + str(tMvin) + "," + str(mWvin) + "^FD" + zpl_safe(v["vin"]) + "^FS\r\n"

        out += "^PQ" + str(copies) + ",1,0,Y\r\n"
        out += "^XZ\r\n"

    return out, len(vehicles), fine_count


# ---------------------------------------------------------------------------
# Send ZPL to printer via TCP
# ---------------------------------------------------------------------------
def send_zpl_to_printer(zpl_text, printer_ip, printer_port):
    data = zpl_text.encode("utf-8")
    client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    client.settimeout(10)
    client.connect((printer_ip, printer_port))
    client.sendall(data)
    time.sleep(0.8)
    client.close()
    return len(data)


# ---------------------------------------------------------------------------
# Label drawing helpers (matches HTML template rendering exactly)
# ---------------------------------------------------------------------------
def draw_label_on_canvas(c, v, cx, cy, w_px, h_px, combo, L, adj_x, adj_y):
    """Draw a single label onto a tkinter Canvas at (cx,cy) size (w_px, h_px).
    Matches the HTML template's mainElHTML + miniCellHTML exactly at 96 px/in."""
    PX = 96.0
    W = w_px
    H = h_px
    main_top_px = (1.75 if combo else 0) * PX
    shift_x = adj_x * PX
    shift_y = adj_y * PX

    def pct_y(val_in):
        return cy + shift_y + (val_in / (H / PX)) * H

    def pct_x(val_in):
        return cx + shift_x + (val_in / (W / PX)) * W

    def pct_w(val_in):
        return (val_in / (W / PX)) * W

    def pct_h(val_in):
        return (val_in / (H / PX)) * H

    def fs_stock():
        return max(5, js_round(L["tstock"] * PX))
    def fs_vin():
        return max(5, js_round(L["tvin"] * PX))
    def fs_info():
        return max(5, js_round(L["tinfo"] * PX))

    avail_w = W * 0.94

    def draw_text_centered(text, y_center, fs, font_family="Arial", bold=True,
                           center_x=None, max_w=None):
        actual_fs = fs
        use_x = center_x if center_x is not None else cx + W / 2
        use_w = max_w if max_w is not None else avail_w
        f = tkfont.Font(family=font_family, size=-actual_fs, weight="bold" if bold else "normal")
        tw = f.measure(text)
        if tw > use_w and tw > 0:
            actual_fs = max(5, int(actual_fs * use_w / tw))
            f = tkfont.Font(family=font_family, size=-actual_fs, weight="bold" if bold else "normal")
        c.create_text(use_x, y_center, text=text, font=f, anchor="center",
                      width=int(use_w) + 4)
        return actual_fs

    def draw_barcode(payload, bar_x, bar_y, bar_w, bar_h):
        bits = code128_bits(payload)
        n = len(bits)
        if n == 0:
            return
        bw = bar_w / n
        for bi, b in enumerate(bits):
            if b == "1":
                bx = bar_x + bi * bw
                c.create_rectangle(bx, bar_y, bx + bw, bar_y + bar_h,
                                   fill="black", outline="")

    def draw_rect_border(x, y, w, h, dash=(4, 4)):
        c.create_rectangle(x, y, x + w, y + h, fill="white",
                           outline="#999", dash=dash, width=1)

    label_x = cx
    label_y = cy
    label_w = W
    label_h = H

    draw_rect_border(label_x, label_y, label_w, label_h)

    top = label_y + shift_y
    left = label_x + shift_x

    if combo:
        cellH = 0.875
        mini_specs = [(0, 0), (0, 2), (0.875, 0), (0.875, 2)]
        for my, ix in mini_specs:
            cellY_in = my + L["miniY"]
            cellX = left + (ix / 4.0) * label_w
            cellY = top + (cellY_in / 5.0) * label_h
            cellW = (2.0 / 4.0) * label_w
            cellHpx = (cellH / 5.0) * label_h

            c.create_rectangle(cellX, cellY, cellX + cellW, cellY + cellHpx,
                               fill="white", outline="#c0c0c0")

            fs_ms = max(3, js_round(L["tmstock"] * PX))
            ms_y = cellY + (L["mstockY"] / cellH) * cellHpx + (L["mstockH"] / cellH) * cellHpx / 2
            draw_text_centered(v["stock"], ms_y, fs_ms, "Arial", True,
                               center_x=cellX + cellW / 2, max_w=cellW)

            mbar_y = cellY + (L["mbarY"] / cellH) * cellHpx
            mbar_h = (L["mbarH"] / cellH) * cellHpx
            mbar_w = cellW * 0.96
            mbar_x = cellX + cellW * 0.02
            draw_barcode(v["vin"], mbar_x, mbar_y, mbar_w, mbar_h)

            fs_mv = max(3, js_round(L["tmvin"] * PX))
            mv_y = cellY + (L["mvinY"] / cellH) * cellHpx + (L["mvinH"] / cellH) * cellHpx / 2
            draw_text_centered(v["vin"], mv_y, fs_mv, "Consolas", True,
                               center_x=cellX + cellW / 2, max_w=cellW)

    mainTop_in = 1.75 if combo else 0
    main_h_in = (5.0 if combo else 2.0) - mainTop_in

    stock_top_in = mainTop_in + L["stockY"]
    stock_h_in = L["stockH"]
    stock_y_px = top + (stock_top_in / (H / PX)) * H
    stock_h_px = (stock_h_in / (H / PX)) * H
    draw_text_centered(v["stock"], stock_y_px + stock_h_px / 2, fs_stock(), "Arial", True)

    bar_top_in = mainTop_in + L["barY"]
    bar_h_in = L["barH"]
    bar_y_px = top + (bar_top_in / (H / PX)) * H
    bar_h_px = (bar_h_in / (H / PX)) * H
    bar_w_px = W * 0.96
    bar_x_px = left + W * 0.02
    draw_barcode(v["vin"], bar_x_px, bar_y_px, bar_w_px, bar_h_px)

    vin_top_in = mainTop_in + L["vinY"]
    vin_h_in = L["vinH"]
    vin_y_px = top + (vin_top_in / (H / PX)) * H
    vin_h_px = (vin_h_in / (H / PX)) * H
    draw_text_centered(v["vin"], vin_y_px + vin_h_px / 2, fs_vin(), "Consolas", True)

    info_top_in = mainTop_in + L["infoY"]
    info_h_in = L["infoH"]
    info_y_px = top + (info_top_in / (H / PX)) * H
    info_h_px = (info_h_in / (H / PX)) * H
    draw_text_centered(car_info(v), info_y_px + info_h_px / 2, fs_info(), "Arial", True)


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------
class VehicleLabelsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Vehicle Labels")
        self.root.geometry("1200x780")
        self.root.minsize(960, 640)

        self.cfg = load_config()
        self.vehicles = []
        self.filtered = []
        self.locations = []

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", rowheight=22, font=("Consolas", 9))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        self._build_ui()

    def _build_ui(self):
        toolbar = ttk.Frame(self.root, padding=4)
        toolbar.pack(fill=tk.X)

        ttk.Button(toolbar, text="1. Import report", command=self._import_xls).pack(side=tk.LEFT, padx=2)
        self.lbl_info = ttk.Label(toolbar, text="No data loaded")
        self.lbl_info.pack(side=tk.LEFT, padx=8)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4)

        ttk.Label(toolbar, text="Search:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._apply_filter())
        ttk.Entry(toolbar, textvariable=self.search_var, width=16).pack(side=tk.LEFT, padx=2)

        ttk.Label(toolbar, text="Location:").pack(side=tk.LEFT, padx=(6, 0))
        self.loc_var = tk.StringVar(value="All")
        self.loc_combo = ttk.Combobox(toolbar, textvariable=self.loc_var, state="readonly", width=12)
        self.loc_combo.pack(side=tk.LEFT, padx=2)
        self.loc_combo.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())

        ttk.Button(toolbar, text="All", command=self._select_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="None", command=self._select_none).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4)

        ttk.Label(toolbar, text="Copies:").pack(side=tk.LEFT)
        self.copies_var = tk.StringVar(value="1")
        ttk.Spinbox(toolbar, from_=1, to=99, textvariable=self.copies_var, width=4).pack(side=tk.LEFT, padx=2)

        ttk.Label(toolbar, text="Combo:").pack(side=tk.LEFT, padx=(6, 0))
        self.combo_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(toolbar, variable=self.combo_var, command=self._refresh_preview).pack(side=tk.LEFT)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4)

        ttk.Label(toolbar, text="DPI:").pack(side=tk.LEFT)
        self.dpi_var = tk.StringVar(value=str(self.cfg["dpi"]))
        dpi_cb = ttk.Combobox(toolbar, textvariable=self.dpi_var, values=["300", "203"], width=4, state="readonly")
        dpi_cb.pack(side=tk.LEFT, padx=2)
        dpi_cb.bind("<<ComboboxSelected>>", lambda _: self._refresh_preview())

        ttk.Label(toolbar, text="X shift:").pack(side=tk.LEFT, padx=(6, 0))
        self.adj_x_var = tk.StringVar(value=f"{self.cfg['adj_x']:.2f}")
        ttk.Entry(toolbar, textvariable=self.adj_x_var, width=6).pack(side=tk.LEFT, padx=1)
        self.adj_x_var.trace_add("write", lambda *_: self._refresh_preview())

        ttk.Label(toolbar, text="Y shift:").pack(side=tk.LEFT, padx=(6, 0))
        self.adj_y_var = tk.StringVar(value=f"{self.cfg['adj_y']:.2f}")
        ttk.Entry(toolbar, textvariable=self.adj_y_var, width=6).pack(side=tk.LEFT, padx=1)
        self.adj_y_var.trace_add("write", lambda *_: self._refresh_preview())

        self.lbl_count = ttk.Label(toolbar, text="0 selected")
        self.lbl_count.pack(side=tk.RIGHT, padx=4)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4)

        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=4, pady=2)

        left_frame = ttk.Frame(main_pane)
        right_frame = ttk.Frame(main_pane)
        main_pane.add(left_frame, weight=3)
        main_pane.add(right_frame, weight=2)

        cols = ("stock", "vin", "type", "yearMake", "color", "location", "count")
        self.tree = ttk.Treeview(left_frame, columns=cols, show="headings", selectmode="extended")
        for c, w in [("stock", 90), ("vin", 170), ("type", 60), ("yearMake", 80),
                      ("color", 110), ("location", 100), ("count", 45)]:
            self.tree.heading(c, text=c.upper())
            self.tree.column(c, width=w, minwidth=50)
        self.tree.pack(fill=tk.BOTH, expand=True)

        tree_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree.yview)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_tree_dblclick)

        right_top = ttk.Frame(right_frame)
        right_top.pack(fill=tk.X, padx=4, pady=2)

        ttk.Label(right_top, text="Label Preview", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT)
        self.preview_hint = ttk.Label(right_top, text="(click row to preview)", font=("Segoe UI", 8))
        self.preview_hint.pack(side=tk.LEFT, padx=6)

        preview_frame = ttk.Frame(right_frame)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=4)

        self.preview_canvas = tk.Canvas(preview_frame, bg="#e8eaed", highlightthickness=0)
        preview_scroll = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        self.preview_canvas.configure(yscrollcommand=preview_scroll.set)
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        btn_bar = ttk.Frame(right_frame)
        btn_bar.pack(fill=tk.X, padx=4, pady=4)

        ttk.Button(btn_bar, text="Preview All Labels", command=self._preview_layout).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_bar, text="Layout Editor", command=self._open_layout_editor).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_bar, text="Download ZPL", command=self._download_zpl).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_bar, text="Send to Printer", command=self._send_to_printer).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_bar, text="Printer...", command=self._printer_settings).pack(side=tk.LEFT, padx=2)

        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, padx=4, pady=(0, 4))
        self.status_var = tk.StringVar(value="Load a report (CDK .xls or formatted .xlsx) to begin - no labels selected by default")
        ttk.Label(status_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X)

    def _import_xls(self):
        path = filedialog.askopenfilename(
            title="Select an inventory report",
            filetypes=[("Inventory reports", "*.xls;*.xlsx"),
                       ("CDK report (.xls)", "*.xls"),
                       ("Formatted workbook (.xlsx)", "*.xlsx"),
                       ("All files", "*.*")],
            initialdir=IMPORTS_DIR,
        )
        if not path:
            return

        self.status_var.set("Reading report ...")
        self.root.update_idletasks()

        try:
            ext = os.path.splitext(path)[1].lower()

            if ext == ".xlsx":
                try:
                    vehicles = parse_xlsx(path)
                    if vehicles:
                        self.vehicles = vehicles
                        self._finalize_load()
                        self.status_var.set(
                            f"Loaded {len(self.vehicles)} vehicles from {os.path.basename(path)} - "
                            "none selected. Use All / double-click to select before printing.")
                        return
                except Exception:
                    pass

            self.status_var.set("Converting .xls ...")
            self.root.update_idletasks()

            base = os.path.splitext(os.path.basename(path))[0]
            formatted_path = os.path.join(FORMATTED_DIR, base + "_formatted.xlsx")
            count = convert_report(path, formatted_path)
            self.status_var.set(f"Converted {count} vehicles. Parsing ...")
            self.root.update_idletasks()

            self.vehicles = parse_xlsx(formatted_path)
            self._finalize_load()
            self.status_var.set(
                f"Loaded {len(self.vehicles)} vehicles from {os.path.basename(path)} - "
                "none selected. Use All / double-click to select before printing.")
        except Exception as e:
            messagebox.showerror("Import Error", str(e))
            self.status_var.set("Import failed")

    def _finalize_load(self):
        for v in self.vehicles:
            v["_sel"] = False
        self.locations = sorted(set(v["location"] for v in self.vehicles if v["location"]))
        self.loc_combo["values"] = ["All"] + self.locations
        self.loc_var.set("All")
        self.search_var.set("")
        self._apply_filter()
        self.lbl_info.config(text=f"{len(self.vehicles)} vehicles loaded (none selected)")

    def _apply_filter(self):
        q = self.search_var.get().strip().upper()
        loc = self.loc_var.get()
        self.filtered = []
        for i, v in enumerate(self.vehicles):
            if loc and loc != "All" and v["location"] != loc:
                continue
            if q:
                haystack = (v["stock"] + " " + v["vin"] + " " + v["location"] + " " + v.get("count", "")).upper()
                if q not in haystack:
                    continue
            self.filtered.append((i, v))
        self._refresh_tree()
        self._refresh_preview()

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for idx, v in self.filtered:
            tag = "sel" if v.get("_sel", True) else "uns"
            self.tree.insert("", tk.END, iid=str(idx),
                             values=(v["stock"], v["vin"], v["type"], v["yearMake"],
                                     v["color"], v["location"], v.get("count", "")),
                             tags=(tag,))
        self.tree.tag_configure("sel", background="#ffffff")
        self.tree.tag_configure("uns", background="#f0d0d0")
        self._update_count()

    def _update_count(self):
        sel = sum(1 for v in self.vehicles if v.get("_sel", True))
        self.lbl_count.config(text=f"{sel} of {len(self.vehicles)} selected")

    def _select_all(self):
        for i, v in self.filtered:
            self.vehicles[i]["_sel"] = True
        self._apply_filter()

    def _select_none(self):
        for i, v in self.filtered:
            self.vehicles[i]["_sel"] = False
        self._apply_filter()

    def _on_tree_select(self, event):
        sel = self.tree.selection()
        if sel:
            idx = int(sel[0])
            self._show_preview(self.vehicles[idx])

    def _on_tree_dblclick(self, event):
        sel = self.tree.selection()
        if sel:
            idx = int(sel[0])
            v = self.vehicles[idx]
            v["_sel"] = not v.get("_sel", True)
            self._apply_filter()

    def _get_adj_x(self):
        try: return float(self.adj_x_var.get())
        except: return 0.0

    def _get_adj_y(self):
        try: return float(self.adj_y_var.get())
        except: return 0.35

    def _get_dpi(self):
        try: return int(self.dpi_var.get())
        except: return 300

    def _show_preview(self, v):
        c = self.preview_canvas
        c.delete("all")
        c.update_idletasks()
        cw = c.winfo_width()
        ch = c.winfo_height()
        if cw < 10:
            cw = 400
            ch = 400

        combo = self.combo_var.get()
        L = load_layout()
        px = 96

        w_in = 4.0
        h_in = 5.0 if combo else 2.0
        max_label_w = cw - 30
        max_label_h = ch - 20
        scale_w = max_label_w / (w_in * px)
        scale_h = max_label_h / (h_in * px)
        scale = min(scale_w, scale_h, 1.2)

        w_px = w_in * px * scale
        h_px = h_in * px * scale
        x0 = (cw - w_px) / 2
        y0 = 10

        draw_label_on_canvas(c, v, x0, y0, w_px, h_px, combo, L,
                             self._get_adj_x(), self._get_adj_y())

        c.configure(scrollregion=c.bbox("all"))

    def _refresh_preview(self):
        sel = self.tree.selection()
        if sel:
            idx = int(sel[0])
            if 0 <= idx < len(self.vehicles):
                self._show_preview(self.vehicles[idx])

    def _preview_layout(self):
        selected = [v for v in self.vehicles if v.get("_sel", True)]
        if not selected:
            messagebox.showinfo("Preview", "No labels selected")
            return

        top = tk.Toplevel(self.root)
        top.title(f"Label Preview \u2014 {len(selected)} labels")
        top.geometry("560x720")

        canvas = tk.Canvas(top, bg="#e8eaed", highlightthickness=0)
        scrollbar = ttk.Scrollbar(top, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        inner = tk.Canvas(canvas, bg="#e8eaed", highlightthickness=0, bd=0)
        canvas.create_window((0, 0), window=inner, anchor="nw")

        combo = self.combo_var.get()
        L = load_layout()
        px = 96
        w_in = 4.0
        h_in = 5.0 if combo else 2.0
        label_w = w_in * px
        label_h = h_in * px
        avail_w = 520
        scale = min((avail_w - 40) / label_w, 1.0)
        disp_w = label_w * scale
        disp_h = label_h * scale
        gap = 16

        y_cursor = gap
        for vi, v in enumerate(selected[:200]):
            inner.create_rectangle(10, y_cursor, 10 + disp_w, y_cursor + disp_h,
                                   fill="white", outline="#999", dash=(4, 4), width=1)
            draw_label_on_canvas(inner, v, 10, y_cursor, disp_w, disp_h, combo, L,
                                 self._get_adj_x(), self._get_adj_y())
            inner.create_text(10 + disp_w / 2, y_cursor + disp_h + 10,
                              text=str(vi + 1), font=("Arial", 7), fill="#999")
            y_cursor += disp_h + 28

        inner.configure(width=avail_w, height=y_cursor)
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _open_layout_editor(self):
        L = load_layout()
        EditorWindow(self.root, L, self)

    def _download_zpl(self):
        selected = [v for v in self.vehicles if v.get("_sel", True)]
        if not selected:
            messagebox.showinfo("Download ZPL", "No labels selected")
            return

        copies = max(1, int(self.copies_var.get() or 1))
        combo = self.combo_var.get()
        dpi = self._get_dpi()
        zpl_text, count, fine = build_zpl(selected, dpi=dpi, combo=combo, copies=copies,
                                          adj_x=self._get_adj_x(), adj_y=self._get_adj_y())

        label_note = ""
        if combo:
            label_note = "\nEach form = 1 main label + 4 mini VIN labels (2\u00d72 grid)"
        ok = messagebox.askyesno(
            "Download ZPL",
            f"About to export {count} label(s).\n\n"
            f"Selected: {count} of {len(self.vehicles)} total vehicles.\n"
            f"Copies each: {copies}{label_note}\n\n"
            f"Download ZPL file?",
        )
        if not ok:
            return

        fname = "zebra_labels_" + time.strftime("%Y%m%d-%H%M") + ".zpl"
        path = filedialog.asksaveasfilename(
            title="Save ZPL file",
            initialfile=fname,
            defaultextension=".zpl",
            filetypes=[("ZPL files", "*.zpl")],
            initialdir=os.path.join(ROOT, "ZPL"),
        )
        if not path:
            return

        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w") as f:
            f.write(zpl_text)
        msg = f"Exported {count} label(s) to {os.path.basename(path)}"
        if fine:
            msg += f" ({fine} used very thin bars)"
        self.status_var.set(msg)

    def _send_to_printer(self):
        selected = [v for v in self.vehicles if v.get("_sel", True)]
        if not selected:
            messagebox.showinfo("Print", "No labels selected")
            return

        copies = max(1, int(self.copies_var.get() or 1))
        combo = self.combo_var.get()
        dpi = self._get_dpi()
        zpl_text, count, fine = build_zpl(selected, dpi=dpi, combo=combo, copies=copies,
                                          adj_x=self._get_adj_x(), adj_y=self._get_adj_y())

        label_note = ""
        if combo:
            label_note = "\nEach form = 1 main label + 4 mini VIN labels (2\u00d72 grid)"

        if count >= 50:
            warn = (
                f"WARNING: You are about to send {count} label(s) to the printer!\n\n"
                f"Selected: {count} of {len(self.vehicles)} total vehicles.\n"
                f"Copies each: {copies}{label_note}\n\n"
                f"Printer: {self.cfg['printer_ip']}:{self.cfg['printer_port']}\n\n"
                f"Are you SURE you want to print?"
            )
        else:
            warn = (
                f"About to print {count} label(s).\n\n"
                f"Printer: {self.cfg['printer_ip']}:{self.cfg['printer_port']}\n\n"
                f"Send to printer?"
            )

        if not messagebox.askyesno("Confirm Print", warn):
            return

        self.status_var.set(f"Sending {count} labels to printer ...")
        self.root.update_idletasks()

        def do_send():
            try:
                bytes_sent = send_zpl_to_printer(
                    zpl_text, self.cfg["printer_ip"], self.cfg["printer_port"]
                )
                self.root.after(0, lambda: self.status_var.set(
                    f"Sent {count} labels ({bytes_sent} bytes) to printer"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Print Error", str(e)))
                self.root.after(0, lambda: self.status_var.set("Print failed"))

        threading.Thread(target=do_send, daemon=True).start()

    def _printer_settings(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Printer Settings")
        dlg.transient(self.root)
        dlg.resizable(False, False)

        ip_var = tk.StringVar(value="")
        port_var = tk.StringVar(value=str(self.cfg["printer_port"]))

        frame = ttk.Frame(dlg, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Printer IP address:").grid(row=0, column=0, sticky=tk.W, pady=3)
        ip_entry = ttk.Entry(frame, textvariable=ip_var, width=20)
        ip_entry.grid(row=0, column=1, sticky=tk.EW, padx=(8, 0), pady=3)

        ttk.Label(frame, text="Port (usually 9100):").grid(row=1, column=0, sticky=tk.W, pady=3)
        ttk.Entry(frame, textvariable=port_var, width=20).grid(row=1, column=1, sticky=tk.EW, padx=(8, 0), pady=3)

        ttk.Label(frame, text="Example: 192.0.2.1", foreground="#666").grid(
            row=2, column=0, columnspan=2, sticky=tk.W, pady=(4, 8))

        status = tk.StringVar()
        ttk.Label(frame, textvariable=status, foreground="#b00").grid(
            row=3, column=0, columnspan=2, sticky=tk.W)

        err = {"msg": None}

        def validate():
            ip = ip_var.get().strip()
            if not ip:
                err["msg"] = "IP address cannot be empty"
                return None
            parts = ip.split(".")
            if len(parts) != 4 or not all(p.isdigit() and 0 <= int(p) <= 255 for p in parts):
                err["msg"] = "Invalid IP address (use dotted form, e.g. 192.0.2.1)"
                return None
            try:
                port = int(port_var.get().strip())
                if not (1 <= port <= 65535):
                    raise ValueError
            except ValueError:
                err["msg"] = "Invalid port (1 - 65535)"
                return None
            return ip, port

        def save(_=None):
            val = validate()
            if val is None:
                status.set(err["msg"])
                return
            ip, port = val
            self.cfg["printer_ip"] = ip
            self.cfg["printer_port"] = port
            save_config(self.cfg)
            self.status_var.set(f"Printer set to {ip}:{port}")
            dlg.destroy()

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=4, column=0, columnspan=2, sticky=tk.E, pady=(10, 0))
        ttk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Save", command=save).pack(side=tk.LEFT, padx=2)

        dlg.bind("<Return>", save)
        dlg.grab_set()
        dlg.wait_window()


# ---------------------------------------------------------------------------
# Layout Editor Window
# ---------------------------------------------------------------------------
class EditorWindow:
    SCALE = 120          # px per inch in editor
    MAIN_BLOCKS = [
        ("stock", "STOCK", True),
        ("bar", "BARCODE", False),
        ("vin", "VIN", True),
        ("info", "CAR INFO", True),
    ]
    MINI_BLOCKS = [
        ("mstock", "STOCK", True),
        ("mbar", "BARCODE", False),
        ("mvin", "VIN", True),
    ]
    MINI_CELLS = [(0, 0), (2, 0), (0, 0.875), (2, 0.875)]

    def __init__(self, parent, layout, app):
        self.app = app
        self.L = dict(layout)
        self.top = tk.Toplevel(parent)
        self.top.title("Layout Editor")
        self.top.geometry("1000x720")
        self.top.transient(parent)
        self.top.grab_set()

        hint = ttk.Label(self.top,
                         text="Drag a block vertically to move it. Resize via its bottom-right "
                              "corner. Drag empty mini-cell space to move the whole 2x2 band. "
                              "Changes are saved to the preview live.",
                         font=("Segoe UI", 9), wraplength=950)
        hint.pack(padx=8, pady=(8, 2), anchor=tk.W)

        body = ttk.Frame(self.top)
        body.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        canvas_frame = ttk.Frame(body)
        canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_frame, bg="#f8f8f8", highlightthickness=1,
                                highlightbackground="#ccc", width=520, height=600)
        vsb = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        hsb = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        ctrl_frame = ttk.Frame(body, width=300)
        ctrl_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(8, 0))
        ctrl_frame.pack_propagate(False)
        self._build_controls(ctrl_frame)

        btn_bar = ttk.Frame(self.top)
        btn_bar.pack(fill=tk.X, padx=8, pady=4)
        ttk.Button(btn_bar, text="Reset Defaults", command=self._reset).pack(side=tk.LEFT)
        ttk.Label(btn_bar, text="Mini block values apply to all 4 cells",
                  font=("Segoe UI", 8), foreground="#666").pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_bar, text="Save & Close", command=self._close).pack(side=tk.RIGHT)

        self.dragging = None
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)

        self._draw()

    # ---------------- controls ----------------
    def _build_controls(self, parent):
        grp_main = ttk.LabelFrame(parent, text="Main 4 in x 3 in")
        grp_main.pack(fill=tk.X, pady=2)
        for key, label, is_text in self.MAIN_BLOCKS:
            self._add_block_row(grp_main, label, key, is_text)

        grp_band = ttk.LabelFrame(parent, text="Minis 2 in x 1 in (all 4)")
        grp_band.pack(fill=tk.X, pady=2)
        row = ttk.Frame(grp_band)
        row.pack(fill=tk.X, pady=1)
        ttk.Label(row, text="GROUP Y", width=9, font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT)
        g_var = tk.StringVar(value=f"{self.L['miniY']:.2f}")
        ttk.Entry(row, textvariable=g_var, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Label(row, text="in").pack(side=tk.LEFT)
        ttk.Label(row, text="(band vertical offset)", font=("Segoe UI", 8),
                  foreground="#666").pack(side=tk.LEFT, padx=6)
        g_var.trace_add("write", lambda *_: self._on_group_change(g_var))

        grp_mini = ttk.LabelFrame(parent, text="Mini block positions (per cell)")
        grp_mini.pack(fill=tk.X, pady=2)
        for key, label, is_text in self.MINI_BLOCKS:
            self._add_block_row(grp_mini, label, key, is_text)

    def _add_block_row(self, parent, label, key, is_text):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=1)
        ttk.Label(row, text=label, width=9, font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT)
        ttk.Label(row, text="Y").pack(side=tk.LEFT)
        y_var = tk.StringVar(value=f"{self.L[key + 'Y']:.2f}")
        ttk.Entry(row, textvariable=y_var, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Label(row, text="H").pack(side=tk.LEFT)
        h_var = tk.StringVar(value=f"{self.L[key + 'H']:.2f}")
        ttk.Entry(row, textvariable=h_var, width=6).pack(side=tk.LEFT, padx=2)
        if is_text:
            ttk.Label(row, text="T").pack(side=tk.LEFT)
            t_var = tk.StringVar(value=f"{self.L['t' + key]:.2f}")
            ttk.Entry(row, textvariable=t_var, width=6).pack(side=tk.LEFT, padx=2)

        def on_change(*_):
            try:
                self.L[key + "Y"] = max(0.0, min(10.0, float(y_var.get())))
                self.L[key + "H"] = max(0.03, min(10.0, float(h_var.get())))
                if is_text:
                    self.L["t" + key] = max(0.02, min(10.0, float(t_var.get())))
                self._commit()
            except ValueError:
                pass

        y_var.trace_add("write", on_change)
        h_var.trace_add("write", on_change)
        if is_text:
            t_var.trace_add("write", on_change)

    def _on_group_change(self, var):
        try:
            self.L["miniY"] = max(-1.0, min(0.5, float(var.get())))
            self._commit()
        except ValueError:
            pass

    # ---------------- drawing ----------------
    def _block_rects(self):
        """Return (key, zone, cell, x, y, w, h, zone_h) for every block in screen px."""
        S = self.SCALE
        L = self.L
        rects = []
        zy = 1.75 * S
        zw = 4.0 * S
        for key, label, is_text in self.MAIN_BLOCKS:
            by = zy + L[key + "Y"] * S
            bh = max(6, L[key + "H"] * S)
            rects.append((key, "main", None, 0.0, by, zw - 2, bh, 3.0))
        for mx, my in self.MINI_CELLS:
            zx = mx * S
            zy = (my + L["miniY"]) * S
            zw = 2.0 * S
            for key, label, is_text in self.MINI_BLOCKS:
                by = zy + L[key + "Y"] * S
                bh = max(6, L[key + "H"] * S)
                rects.append((key, "mini", (mx, my), zx, by, zw - 2, bh, 0.875))
        return rects

    def _draw_barcode(self, bx, by, bw, bh, key):
        c = self.canvas
        bits = code128_bits("LABEL-" + key)
        n = len(bits)
        if not n or bw < 4 or bh < 2:
            return
        unit = bw / n
        x = bx
        for b in bits:
            if b == "1":
                c.create_rectangle(x, by, x + unit, by + bh,
                                   fill="black", outline="")
            x += unit

    def _draw(self):
        c = self.canvas
        c.delete("all")
        S = self.SCALE
        L = self.L

        lw = 4 * S
        lh = 5 * S
        c.create_rectangle(0, 0, lw, lh, outline="#666", width=1)

        main_y = 1.75 * S
        c.create_rectangle(0, main_y, lw, main_y + 3 * S,
                           outline="#888", dash=(6, 3))

        for mx, my in self.MINI_CELLS:
            zx = mx * S
            zy = (my + L["miniY"]) * S
            c.create_rectangle(zx, zy, zx + 2 * S, zy + 0.875 * S,
                               outline="#bbb", dash=(4, 2))

        LABELS = {k: lb for k, lb, t in self.MAIN_BLOCKS + self.MINI_BLOCKS}
        for key, zone, cell, zx, zy, zw, bh, zone_h in self._block_rects():
            is_text = key in ("stock", "vin", "info", "mstock", "mvin")
            fill = "#e8f0fe" if is_text else "#ffffff"
            c.create_rectangle(zx + 1, zy, zx + zw, zy + bh,
                               fill=fill, outline="#1a73e8", width=1)
            label = LABELS[key]
            if is_text:
                t_val = L.get("t" + key, 0)
                fs = max(7, int((t_val if t_val else L[key + "H"]) * S * 0.55))
                c.create_text(zx + zw / 2, zy + bh / 2, text=label,
                              font=("Segoe UI", fs, "bold"), fill="#1557b0")
            else:
                self._draw_barcode(zx + 2, zy + 4, max(1, zw - 4), max(2, bh - 8), key)
            c.create_rectangle(zx + zw - 14, zy + bh - 9, zx + zw, zy + bh,
                               fill="#1a73e8", outline="")

        c.configure(scrollregion=(0, -10, lw + 10, lh + 10))

    # ---------------- interaction ----------------
    def _on_press(self, event):
        S = self.SCALE
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        HANDLE = 14

        for key, zone, cell, rx, ry, rw, rh, zh in self._block_rects():
            if rx + rw - HANDLE <= x <= rx + rw and ry + rh - 9 <= y <= ry + rh:
                self.dragging = {"mode": "resize", "key": key, "cell": cell,
                                 "zone_h": zh, "start_y": y,
                                 "start_h": self.L[key + "H"]}
                return
            if rx <= x <= rx + rw and ry <= y <= ry + rh:
                self.dragging = {"mode": "drag", "key": key, "cell": cell,
                                 "zone_h": zh, "start_y": y,
                                 "start_val": self.L[key + "Y"]}
                return

        for mx, my in self.MINI_CELLS:
            zx = mx * S
            zy = (my + self.L["miniY"]) * S
            if zx <= x <= zx + 2 * S and zy <= y <= zy + 0.875 * S:
                self.dragging = {"mode": "band", "start_y": y,
                                 "start_val": self.L["miniY"]}
                return

    def _on_drag(self, event):
        if not self.dragging:
            return
        S = self.SCALE
        d = self.dragging
        d_in = (self.canvas.canvasy(event.y) - d["start_y"]) / S

        if d["mode"] == "band":
            self.L["miniY"] = round(max(-1.0, min(0.5, d["start_val"] + d_in)), 3)
        elif d["mode"] == "resize":
            key = d["key"]
            h = max(0.03, d["start_h"] + d_in)
            y = self.L[key + "Y"]
            if y + h > d["zone_h"]:
                y = max(0.0, d["zone_h"] - h)
            self.L[key + "Y"] = round(y, 3)
            self.L[key + "H"] = round(h, 3)
        else:  # drag
            key = d["key"]
            h = self.L[key + "H"]
            y = max(0.0, min(d["zone_h"] - h, d["start_val"] + d_in))
            self.L[key + "Y"] = round(y, 3)
        self._commit()

    def _on_release(self, event):
        self.dragging = None

    # ---------------- save / apply ----------------
    def _commit(self):
        save_layout(self.L)
        self._draw()
        if self.app:
            try:
                self.app._refresh_preview()
            except Exception:
                pass

    def _reset(self):
        self.L = dict(LAYOUT_DEFAULTS)
        save_layout(self.L)
        if self.app:
            try:
                self.app._refresh_preview()
            except Exception:
                pass
        self._draw()

    def _close(self):
        save_layout(self.L)
        if self.app:
            try:
                self.app.cfg["adj_x"] = self.app._get_adj_x()
                self.app.cfg["adj_y"] = self.app._get_adj_y()
                self.app.cfg["dpi"] = self.app._get_dpi()
                save_config(self.app.cfg)
                self.app._refresh_preview()
            except Exception:
                pass
        self.top.destroy()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    root = tk.Tk()
    app = VehicleLabelsApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
