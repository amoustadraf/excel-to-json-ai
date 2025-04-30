from openpyxl import load_workbook
from openpyxl.styles.colors import COLOR_INDEX
import zipfile
import xml.etree.ElementTree as ET
import json
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# Hide main tkinter window
root = tk.Tk()
root.withdraw()

# Ask user to select a file
file_path = filedialog.askopenfilename(
    title="Select an Excel file",
    filetypes=[("Excel files", "*.xlsx *.xlsm")]
)

# Check if the user cancelled or selected a non-Excel file
if not file_path or not file_path.endswith((".xlsx", ".xlsm")):
    messagebox.showerror("Invalid file", "You must select a valid Excel file (.xlsx or .xlsm).")
    raise SystemExit("Invalid file selected.")

# Function to load theme colors from the Excel file
def load_theme_colors(filename):
    with zipfile.ZipFile(filename, 'r') as archive:
        theme_path = [p for p in archive.namelist() if p.startswith('xl/theme/theme')][0]
        with archive.open(theme_path) as theme_file:
            tree = ET.parse(theme_file)
            root = tree.getroot()
            namespace = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
            theme_elements = root.find('.//a:clrScheme', namespace)
            theme_colors = []
            for elem in theme_elements:
                srgb = elem.find('.//a:srgbClr', namespace)
                sysrgb = elem.find('.//a:sysClr', namespace)
                if srgb is not None:
                    theme_colors.append(srgb.attrib['val'])
                elif sysrgb is not None:
                    theme_colors.append(sysrgb.attrib['lastClr'])
                else:
                    theme_colors.append('000000')
            return theme_colors

# Apply tint to RGB color
def apply_tint(rgb, tint):
    if rgb is None or len(rgb) != 6:
        return None
    r = int(rgb[0:2], 16)
    g = int(rgb[2:4], 16)
    b = int(rgb[4:6], 16)
    def tint_component(comp):
        if tint < 0:
            return int(comp * (1 + tint))
        else:
            return int((255 - comp) * tint + comp)
    r = max(0, min(255, tint_component(r)))
    g = max(0, min(255, tint_component(g)))
    b = max(0, min(255, tint_component(b)))
    return '{:02X}{:02X}{:02X}'.format(r, g, b)

# Validate hex color
def hex_color(obj, theme_colors):
    color = obj.color
    if not color:
        return None
    if getattr(color, 'type', None) == 'rgb' and getattr(color, 'rgb', None):
        c = str(color.rgb)
    elif hasattr(color, 'indexed') and isinstance(color.indexed, int):
        idx = color.indexed
        if 0 <= idx < len(COLOR_INDEX):
            c = COLOR_INDEX[idx]
        else:
            return None
    elif getattr(color, 'theme', None) is not None:
        theme_idx = color.theme
        if 0 <= theme_idx < len(theme_colors):
            base_color = theme_colors[theme_idx]
            tint = getattr(color, 'tint', 0)
            c = apply_tint(base_color, tint)
        else:
            return None
    else:
        return None
    c = c.lstrip('#').upper()
    if len(c) == 6:
        c = 'FF' + c
    if c in ('FF000000', '00000000'):
        return None
    return c

# Load Excel file
excel_file = file_path
theme_colors = load_theme_colors(excel_file)
wb = load_workbook(excel_file, data_only=True)
data = {}

# Iterate through each sheet and extract cell data
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    sheet_data = {}
    for row in ws.iter_rows():
        for cell in row:
            val = cell.value
            if isinstance(val, (datetime.datetime, datetime.date)):
                val = val.strftime("%m/%d/%Y")
            sheet_data[cell.coordinate] = {
                "value": val,
                "bold": bool(cell.font.bold),
                "italic": bool(cell.font.italic),
                "underline": bool(cell.font.underline),
                "strikethrough": bool(cell.font.strike),
                "font_color": hex_color(cell.font, theme_colors),
                "fill_color": hex_color(cell, theme_colors),
                "border_top": cell.border.top.style,
                "border_bottom": cell.border.bottom.style,
                "border_left": cell.border.left.style,
                "border_right": cell.border.right.style,
            }
    data[sheet_name] = sheet_data

# Save extracted data to JSON file
with open("output.json", "w", encoding="utf-8") as f:
    json.dump(data, f, indent=4)