import pandas as pd
import io
import tempfile
import os
import zipfile
from fastapi import FastAPI, UploadFile, Form, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont
from typing import Optional

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Name substitution map ────────────────────────────────────────────────────
SCHOOL_NAME_SUBSTITUTIONS = {
    "Mount Carmel Convent Sr. Sec. School Cement Nagar": "Mount Carmel Convent Sr. Sec. School Ghughus",
}

def apply_name_substitution(name: str) -> str:
    return SCHOOL_NAME_SUBSTITUTIONS.get(name, name)

# ── Helpers ──────────────────────────────────────────────────────────────────
def to_int(val):
    try:
        if pd.isna(val) or str(val).strip() == '' or val is None:
            return None
        return int(float(val))
    except (ValueError, TypeError):
        return None

def format_date_str(d_str):
    if not d_str:
        return ""
    try:
        dt = pd.to_datetime(d_str)
        day = dt.day
        if 4 <= day <= 20 or 24 <= day <= 30:
            suffix = "th"
        else:
            suffix = ["st", "nd", "rd"][day % 10 - 1]
        return f"{day}{suffix}/{dt.strftime('%B')}/{dt.year}"
    except Exception:
        return d_str

def clean_sheet_name(name, existing_names):
    name = str(name).strip()
    invalid_chars = [':', '\\', '/', '?', '*', '[', ']']
    for ch in invalid_chars:
        name = name.replace(ch, '')
    base_name = name[:31].strip()
    if not base_name:
        base_name = "Sheet"
    final_name = base_name
    counter = 1
    while final_name.lower() in [s.lower() for s in existing_names]:
        suffix = f"_{counter}"
        avail_len = 31 - len(suffix)
        final_name = f"{base_name[:avail_len].strip()}{suffix}"
        counter += 1
    return final_name

def safe_filename(name: str) -> str:
    """Strip characters that are invalid in Windows filenames."""
    invalid = r'\/:*?"<>|'
    for ch in invalid:
        name = name.replace(ch, '')
    return name.strip()

def extract_data_bytes(file_bytes, col_school, col_obs, req_teacher_and_unique=False, col_teachers=0, col_unique=0):
    try:
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None)
            dfs_to_process = [df]
        except ValueError:
            try:
                df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='xlrd')
                dfs_to_process = [df]
            except Exception:
                dfs_to_process = pd.read_html(io.BytesIO(file_bytes), header=None)
    except Exception as e:
        raise Exception(f"Failed to read file: {str(e)}")

    data = {}
    school_order = []

    for df_item in dfs_to_process:
        for idx, row in df_item.iterrows():
            try:
                if len(row) <= max(col_school, col_obs,
                                   col_teachers if req_teacher_and_unique else 0,
                                   col_unique if req_teacher_and_unique else 0):
                    continue

                school_raw = row.iloc[col_school]
                if pd.isna(school_raw):
                    continue
                school = str(school_raw).strip()

                if not school or school.lower() in ['nan', 'none', 'school', 'school name'] \
                        or 'observation report' in school.lower() \
                        or 'session :' in school.lower():
                    continue

                location = ""
                loc_col = col_school + 1
                if loc_col < len(row) and loc_col not in [col_obs, col_teachers, col_unique]:
                    loc_val = row.iloc[loc_col]
                    if pd.notna(loc_val):
                        loc_str = str(loc_val).strip()
                        if loc_str and loc_str.lower() not in ['nan', 'none'] and to_int(loc_val) is None:
                            location = loc_str

                obs_int = to_int(row.iloc[col_obs])
                if obs_int is None:
                    continue

                teachers_int = None
                unique_int = None
                if req_teacher_and_unique:
                    teachers_int = to_int(row.iloc[col_teachers])
                    unique_int = to_int(row.iloc[col_unique])
                    if teachers_int is None or unique_int is None:
                        continue

                if location:
                    school_key = f"{school} {location}".upper()
                    display_name = f"{school} {location}"
                else:
                    school_key = school.upper()
                    display_name = school

                # Apply substitution
                display_name = apply_name_substitution(display_name)

                item = {
                    'name': display_name,
                    'school_base': school,
                    'location': location,
                    'obs': obs_int
                }
                if req_teacher_and_unique:
                    item['teachers'] = teachers_int
                    item['unique'] = unique_int

                if school_key not in data:
                    data[school_key] = item
                    school_order.append(school_key)
                else:
                    dup_idx = 2
                    unique_key = f"{school_key} ({dup_idx})"
                    while unique_key in data:
                        dup_idx += 1
                        unique_key = f"{school_key} ({dup_idx})"
                    item['name'] = f"{display_name} ({dup_idx})"
                    data[unique_key] = item
                    school_order.append(unique_key)

            except Exception:
                continue

    return data, school_order


# ── Border helpers ────────────────────────────────────────────────────────────
def make_thin_side():
    return Side(style='thin', color='000000')

def make_thick_side():
    return Side(style='medium', color='000000')

def apply_borders_to_sheet(ws, min_row, max_row, min_col, max_col):
    """Apply thin inner borders and thick outer border to a cell range."""
    thin = make_thin_side()
    thick = make_thick_side()
    no_side = Side(style=None)

    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            top    = thick if r == min_row  else thin
            bottom = thick if r == max_row  else thin
            left   = thick if c == min_col  else thin
            right  = thick if c == max_col  else thin
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)


# ── Image generation ──────────────────────────────────────────────────────────
def generate_school_image(school_name: str, output_data: list) -> bytes:
    """
    Render a styled table as a PNG image and return the raw bytes.
    output_data is a list of [label, value] rows (including the header as first item).
    """
    # ── Layout constants ─────────────────────────────────────────────────────
    COL_A_W   = 650   # px width of Observation column
    COL_B_W   = 450   # px width of Count column
    ROW_H     = 38    # px height per row
    PAD       = 24    # outer padding
    HEADER_H  = 50    # header row height

    # Colours
    HEADER_BG   = (192,  0,  0)   # dark red
    HEADER_FG   = (255, 255, 255)
    ODD_BG      = (255, 255, 255)
    EVEN_BG     = (242, 242, 242)
    TEXT_FG     = (30,  30,  30)
    BORDER_COL  = (0,   0,   0)
    OUTER_W     = 3
    INNER_W     = 1

    rows_count = len(output_data) + 1   # +1 for header row
    total_w = PAD * 2 + COL_A_W + COL_B_W
    total_h = PAD * 2 + HEADER_H + (rows_count - 1) * ROW_H

    img = Image.new('RGB', (total_w, total_h), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)

    # Try to load a system font; fall back to default
    try:
        font_regular = ImageFont.truetype("arial.ttf", 15)
        font_bold    = ImageFont.truetype("arialbd.ttf", 15)
        font_header  = ImageFont.truetype("arialbd.ttf", 16)
    except OSError:
        font_regular = ImageFont.load_default()
        font_bold    = font_regular
        font_header  = font_regular

    x0 = PAD
    y0 = PAD

    # ── Draw header row ───────────────────────────────────────────────────────
    draw.rectangle([x0, y0, x0 + COL_A_W + COL_B_W, y0 + HEADER_H], fill=HEADER_BG)
    draw.text((x0 + 10, y0 + HEADER_H // 2 - 8), "Observation", font=font_header, fill=HEADER_FG)
    count_label = "Count"
    try:
        cl_w = draw.textlength(count_label, font=font_header)
    except AttributeError:
        cl_w = len(count_label) * 9
    draw.text((x0 + COL_A_W + (COL_B_W - cl_w) // 2, y0 + HEADER_H // 2 - 8),
              count_label, font=font_header, fill=HEADER_FG)

    # ── Draw data rows ────────────────────────────────────────────────────────
    for i, (label, value) in enumerate(output_data):
        ry = y0 + HEADER_H + i * ROW_H
        bg = ODD_BG if i % 2 == 0 else EVEN_BG
        draw.rectangle([x0, ry, x0 + COL_A_W + COL_B_W, ry + ROW_H], fill=bg)

        # Label text (left-aligned in col A)
        label_str = str(label) if label else ''
        draw.text((x0 + 10, ry + ROW_H // 2 - 8), label_str, font=font_regular, fill=TEXT_FG)

        # Value text (center-aligned in col B)
        val_str = str(value) if value not in (None, '', 0, 0.0) or value == 0 else ''
        if label == '':
            val_str = ''
        try:
            vw = draw.textlength(val_str, font=font_bold)
        except AttributeError:
            vw = len(val_str) * 9
        draw.text((x0 + COL_A_W + (COL_B_W - vw) // 2, ry + ROW_H // 2 - 8),
                  val_str, font=font_bold, fill=TEXT_FG)

    table_w = COL_A_W + COL_B_W
    table_h = HEADER_H + len(output_data) * ROW_H

    # ── Draw inner grid lines ─────────────────────────────────────────────────
    # Horizontal lines
    for i in range(1, len(output_data) + 1):
        y = y0 + HEADER_H + i * ROW_H
        draw.line([(x0, y), (x0 + table_w, y)], fill=BORDER_COL, width=INNER_W)
    # Header bottom
    draw.line([(x0, y0 + HEADER_H), (x0 + table_w, y0 + HEADER_H)], fill=BORDER_COL, width=INNER_W)
    # Vertical divider between col A and col B
    draw.line([(x0 + COL_A_W, y0), (x0 + COL_A_W, y0 + table_h)], fill=BORDER_COL, width=INNER_W)

    # ── Draw outer thick border ───────────────────────────────────────────────
    x1 = x0 + table_w
    y1 = y0 + table_h
    for t in range(OUTER_W):
        draw.rectangle([x0 - t, y0 - t, x1 + t, y1 + t], outline=BORDER_COL)

    buf = io.BytesIO()
    img.save(buf, format='PNG')
    return buf.getvalue()


# ── API endpoint ──────────────────────────────────────────────────────────────
@app.post("/api/generate")
async def generate_report(
    file1: UploadFile = File(...),
    file2: Optional[UploadFile] = File(None),
    file3: Optional[UploadFile] = File(None),
    year_start: str = Form("1st/April/2025"),
    year_end: str = Form("1st/April/2026"),
    t1_start: str = Form("1/April/2025"),
    t1_end: str = Form("15/oct/2025"),
    t2_start: str = Form("16/Oct/2025"),
    t2_end: str = Form("1st/April/2026"),
    col_school: int = Form(1),
    col_teachers: int = Form(3),
    col_unique: int = Form(6),
    col_obs: int = Form(5),
    col_term_obs: int = Form(8)
):
    f1_bytes = await file1.read()
    f2_bytes = await file2.read() if file2 else None
    f3_bytes = await file3.read() if file3 else None

    col_school    -= 1
    col_teachers  -= 1
    col_unique    -= 1
    col_obs       -= 1
    col_term_obs  -= 1

    data1, order1 = extract_data_bytes(f1_bytes, col_school, col_obs,
                                        req_teacher_and_unique=True,
                                        col_teachers=col_teachers,
                                        col_unique=col_unique)
    data2, _ = extract_data_bytes(f2_bytes, col_school, col_term_obs) if f2_bytes else ({}, [])
    data3, _ = extract_data_bytes(f3_bytes, col_school, col_term_obs) if f3_bytes else ({}, [])

    formatted_year_start = format_date_str(year_start)
    formatted_year_end   = format_date_str(year_end)
    formatted_t1_start   = format_date_str(t1_start)
    formatted_t1_end     = format_date_str(t1_end)
    formatted_t2_start   = format_date_str(t2_start)
    formatted_t2_end     = format_date_str(t2_end)

    tmp_xlsx = tempfile.mktemp(suffix=".xlsx")
    tmp_zip  = tempfile.mktemp(suffix=".zip")

    used_sheet_names = set()

    # ── Collect all school data for image generation ──────────────────────────
    school_image_data = {}   # school_key -> (school_name, output_data)

    with pd.ExcelWriter(tmp_xlsx, engine='openpyxl') as writer:
        for school_key in order1:
            d1 = data1[school_key]
            school_name   = d1.get('name', school_key)
            teachers      = d1.get('teachers', 0)
            obs_all       = d1.get('obs', 0)
            unique        = d1.get('unique', 0)

            target         = teachers * 8
            perc_all       = round((obs_all / target * 100), 2) if target > 0 else 0
            not_observed   = max(0, teachers - unique)

            d2             = data2.get(school_key, data2.get(d1.get('school_base', '').upper(), {}))
            term1_target   = int(target / 2)
            term1_obs      = d2.get('obs', 0)
            perc_1         = round((term1_obs / term1_target * 100), 2) if term1_target > 0 else 0

            d3             = data3.get(school_key, data3.get(d1.get('school_base', '').upper(), {}))
            term2_target   = int(target / 2)
            term2_obs      = d3.get('obs', 0)
            perc_2         = round((term2_obs / term2_target * 100), 2) if term2_target > 0 else 0

            output_data = [
                ['School Name', school_name],
                ['No. of Teachers', teachers],
                ['No. of Target observations (8 per teacher per year x no of teacher)', target],
                [f'Observations till date ({formatted_year_start} to {formatted_year_end})', obs_all],
                ['To observation Percentage till date', perc_all],
                ['Unique Teacher Count', unique],
                ['No. of Teachers not observed once', not_observed],
                ['', '']
            ]

            if file2:
                output_data.extend([
                    [f'Term Target 1 ({formatted_t1_start} to {formatted_t1_end})', term1_target],
                    ['Observation Term 1', term1_obs],
                    ['Percentage of Term 1 observation', perc_1],
                    ['', '']
                ])

            if file3:
                output_data.extend([
                    [f'Term Target 2 ({formatted_t2_start} to {formatted_t2_end})', term2_target],
                    ['Observation Term 2', term2_obs],
                    ['Percentage of Term 2 observation', perc_2]
                ])

            # Store for image generation
            school_image_data[school_key] = (school_name, list(output_data))

            df_out = pd.DataFrame(output_data, columns=['Observation', 'Count'])
            sname  = clean_sheet_name(school_name, used_sheet_names)
            used_sheet_names.add(sname)

            df_out.to_excel(writer, sheet_name=sname, index=False)

            ws = writer.sheets[sname]

            # ── Header styling ────────────────────────────────────────────────
            header_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            for cell in ws["1:1"]:
                cell.fill   = header_fill
                cell.font   = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # ── Column widths ─────────────────────────────────────────────────
            ws.column_dimensions['A'].width = 65
            ws.column_dimensions['B'].width = 45

            # ── Centre-align Count column (B) for all data rows ───────────────
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=2):
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')

            # ── Borders: thin inside, thick (medium) outside ──────────────────
            max_r = ws.max_row
            max_c = 2   # columns A and B
            apply_borders_to_sheet(ws, min_row=1, max_row=max_r, min_col=1, max_col=max_c)

    # ── Build ZIP: Excel + per-school PNG images ──────────────────────────────
    with zipfile.ZipFile(tmp_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(tmp_xlsx, "Observation_Report.xlsx")

        for school_key, (school_name, od) in school_image_data.items():
            img_bytes = generate_school_image(school_name, od)
            fname = safe_filename(school_name) + ".png"
            zf.writestr(fname, img_bytes)

    os.remove(tmp_xlsx)

    return FileResponse(
        tmp_zip,
        filename="Observation_Reports.zip",
        media_type="application/zip"
    )
