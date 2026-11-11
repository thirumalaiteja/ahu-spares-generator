import io
import os
import re
import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(
    page_title="AHU/FAHU Spare Parts Generator",
    page_icon="⚙️",
    layout="wide"
)

# ----------------- PARSING LOGIC -----------------
def parse_pdf_datasheet(pdf_bytes, filename="datasheet.pdf"):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    unit_pages = []
    current_unit = None
    
    for page_num in range(len(doc)):
        text = doc[page_num].get_text()
        
        match_size = re.search(r'UNIT SIZE:\s*([A-Za-z0-9\-_]+)', text, re.IGNORECASE)
        if match_size:
            model = match_size.group(1).strip()
            current_unit = {'model': model, 'pages': []}
            unit_pages.append(current_unit)
        elif current_unit is None and page_num == 0:
            match_ref = re.search(r'Unit reference\s*\n\s*([^\n]+)', text, re.IGNORECASE)
            model = match_ref.group(1).strip() if match_ref else os.path.splitext(filename)[0]
            current_unit = {'model': model, 'pages': []}
            unit_pages.append(current_unit)
            
        if current_unit:
            current_unit['pages'].append((page_num + 1, text))

    units_data = []
    
    for u_idx, u in enumerate(unit_pages, 1):
        model_name = u['model']
        unit_items = []
        
        for p_num, text in u['pages']:
            lines = [l.strip() for l in text.splitlines() if l.strip()]
            
            # --- Extract Filters ---
            if any(k in text for k in ['PANEL FILTER', 'RIGID BAG FILTER', 'SOFT BAG FILTER', 'BAG FILTER', 'FILTER']):
                current_filter_prefix = ''
                i = 0
                while i < len(lines):
                    line = lines[i]
                    if re.match(r'^(G[0-9]|F[0-9]|M[0-9]|E[0-9]|H[0-9]|U[0-9])\s+(EN[0-9]+|synthetic|ISO)', line, re.I) or 'PANEL FILTER' in line or 'RIGID BAG FILTER' in line:
                        filter_code = line.split()[0] if not 'FILTER' in line else ''
                        iso_class = ''
                        for j in range(i, min(i + 15, len(lines))):
                            if 'ISO Coarse' in lines[j] or 'ePM' in lines[j]:
                                iso_class = lines[j]
                                break
                        
                        iso_class = re.sub(r'ePM\s*1\s*', 'ePM1 ', iso_class)
                        
                        if filter_code.upper().startswith('G'):
                            current_filter_prefix = f'{filter_code} synthetic filter {iso_class}'.strip()
                        elif filter_code.upper().startswith('F'):
                            current_filter_prefix = f'{filter_code} Rigid bag filter {iso_class}'.strip()
                        else:
                            if 'Coarse' in iso_class:
                                current_filter_prefix = f'G4 synthetic filter {iso_class}'.strip()
                            elif 'ePM' in iso_class:
                                current_filter_prefix = f'F7 Rigid bag filter {iso_class}'.strip()
                            else:
                                current_filter_prefix = f'Synthetic filter {iso_class}'.strip()
                    
                    qty_match = re.match(r'^(\d+)\s*$', line)
                    if qty_match and i + 1 < len(lines):
                        next_line = lines[i + 1]
                        dim_match = re.search(r'(\d+\s*x\s*\d+\s*x\s*\d+\s*mm)', next_line)
                        if dim_match:
                            qty = int(qty_match.group(1))
                            dim = dim_match.group(1)
                            dim_formatted = re.sub(r'\s*x\s*', ' x ', dim)
                            desc = f'{current_filter_prefix} {dim_formatted}'.strip()
                            unit_items.append((desc, qty))
                            i += 1
                    i += 1

            # --- Extract Belts & Pulleys ---
            if 'SUPPLY FAN' in text:
                for line in lines:
                    if line.startswith('Belt '):
                        m = re.search(r'(Belt\s+[A-Za-z0-9\-]+)\s*x(\d+)', line)
                        if m:
                            unit_items.append((' ' + m.group(1), int(m.group(2))))
                    elif line.startswith('Fan PULLEY'):
                        m = re.search(r'(Fan PULLEY\s+[A-Za-z0-9\s]+?)\s*x(\d+)', line)
                        if m:
                            unit_items.append((m.group(1).strip(), int(m.group(2))))
                    elif line.startswith('Motor PULLEY'):
                        m = re.search(r'(Motor PULLEY\s+[A-Za-z0-9\s]+?)\s*x(\d+)', line)
                        if m:
                            unit_items.append((m.group(1).strip(), int(m.group(2))))

        if unit_items:
            units_data.append({
                's_no': u_idx,
                'model': model_name,
                'items': unit_items
            })
            
    doc.close()
    return units_data

# ----------------- EXCEL GENERATOR -----------------
def generate_excel_spares(units_data, template_bytes=None):
    if template_bytes:
        wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
        ws = wb["Sheet2"] if "Sheet2" in wb.sheetnames else wb.active
        for cell_range in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(cell_range))
        for r in range(6, max(ws.max_row + 1, 50)):
            for c in range(1, 10):
                ws.cell(row=r, column=c).value = None
                ws.cell(row=r, column=c).border = Border()
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet2"
        ws.views.sheetView[0].showGridLines = True
        
        col_widths = {'A': 13.0, 'B': 13.0, 'C': 25.33, 'D': 69.33, 'E': 13.0}
        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width

        font_title = Font(name='Calibri', size=11, bold=False)
        font_subtitle = Font(name='Calibri', size=11, bold=True)
        font_header = Font(name='Calibri', size=13, bold=True)
        align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        align_top_left = Alignment(horizontal='left', vertical='top', wrap_text=True)
        
        ws.cell(row=2, column=1, value="Recommended Spare Parts for FAHU (Please note that these details can be found on the datasheets under ordinary maintenance)").font = font_title
        ws.cell(row=2, column=1).alignment = align_top_left
        ws.cell(row=4, column=4, value="SPARE PARTS ").font = font_subtitle
        ws.cell(row=4, column=4).alignment = align_center

        headers = {2: "SI No. ", 3: "Model", 4: "Description", 5: "Qty"}
        for col_idx, text in headers.items():
            cell = ws.cell(row=5, column=col_idx, value=text)
            cell.font = font_header
            cell.alignment = align_center

    font_data = Font(name='Calibri', size=9, bold=False)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    current_row = 6
    for unit in units_data:
        s_no = unit['s_no']
        model = unit['model']
        items = unit['items']
        
        start_row = current_row
        for desc, qty in items:
            c4 = ws.cell(row=current_row, column=4, value=desc)
            c4.font = font_data
            c4.alignment = align_center
            c4.border = thin_border
            
            c5 = ws.cell(row=current_row, column=5, value=qty)
            c5.font = font_data
            c5.alignment = align_center
            c5.border = thin_border
            current_row += 1
            
        end_row = current_row - 1
        if start_row <= end_row:
            ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
            c_sno = ws.cell(row=start_row, column=2, value=s_no)
            c_sno.font = font_data
            c_sno.alignment = align_center
            
            ws.merge_cells(start_row=start_row, start_column=3, end_row=end_row, end_column=3)
            c_mod = ws.cell(row=start_row, column=3, value=f" {model}")
            c_mod.font = font_data
            c_mod.alignment = align_center
            
            for r in range(start_row, end_row + 1):
                for c in [2, 3]:
                    ws.cell(row=r, column=c).border = thin_border

    out_stream = io.BytesIO()
    wb.save(out_stream)
    wb.close()
    return out_stream.getvalue()

# ----------------- PDF EXPORT (OS-INDEPENDENT) -----------------
def generate_pdf_spares(units_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'MainTitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14
    )
    
    elements = [
        Paragraph("Recommended Spare Parts for FAHU (Please note that these details can be found on the datasheets under ordinary maintenance)", title_style),
        Spacer(1, 12),
        Paragraph("<b>SPARE PARTS</b>", ParagraphStyle('Sub', alignment=1, fontSize=12, leading=14)),
        Spacer(1, 10)
    ]
    
    table_data = [["SI No.", "Model", "Description", "Qty"]]
    t_style = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F2F2F2')),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
    ]
    
    curr_row = 1
    for unit in units_data:
        s_no = str(unit['s_no'])
        model = str(unit['model'])
        start = curr_row
        for desc, qty in unit['items']:
            table_data.append([s_no, model, desc, str(qty)])
            curr_row += 1
        end = curr_row - 1
        if start < end:
            t_style.append(('SPAN', (0, start), (0, end)))
            t_style.append(('SPAN', (1, start), (1, end)))

    col_widths = [45, 120, 340, 45]
    pdf_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    pdf_table.setStyle(TableStyle(t_style))
    elements.append(pdf_table)
    
    doc.build(elements)
    return buffer.getvalue()

# ----------------- UI / WORKFLOW -----------------
st.title("🏭 AHU / FAHU / MAHU Spare Parts Generator")
st.write("Upload equipment datasheets to extract filters, belts, and pulleys into standard spare schedules.")

col1, col2 = st.columns([1, 1])
with col1:
    uploaded_pdf = st.file_uploader("1. Upload Datasheet PDF", type=["pdf"])
with col2:
    uploaded_tpl = st.file_uploader("2. (Optional) Custom Template (.xlsx)", type=["xlsx"])

if uploaded_pdf:
    with st.spinner("Extracting unit data..."):
        units_data = parse_pdf_datasheet(uploaded_pdf.read(), uploaded_pdf.name)

    if not units_data:
        st.warning("⚠️ No matching AHU unit models or spare parts detected in this PDF.")
    else:
        st.success(f"Extracted {len(units_data)} unit(s) successfully!")
        
        # Display preview table
        flat_records = []
        for u in units_data:
            for desc, qty in u['items']:
                flat_records.append({
                    "SI No.": u['s_no'],
                    "Model": u['model'],
                    "Description": desc,
                    "Qty": qty
                })
        
        df = pd.DataFrame(flat_records)
        st.subheader("Data Preview")
        st.dataframe(df, use_container_width=True)
        
        # Prepare Downloads
        tpl_bytes = uploaded_tpl.read() if uploaded_tpl else None
        excel_bytes = generate_excel_spares(units_data, tpl_bytes)
        pdf_bytes = generate_pdf_spares(units_data)
        
        base_name = os.path.splitext(uploaded_pdf.name)[0]
        
        c_down1, c_down2 = st.columns(2)
        with c_down1:
            st.download_button(
                label="📥 Download Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"{base_name}_spares.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c_down2:
            st.download_button(
                label="📥 Download PDF Document (.pdf)",
                data=pdf_bytes,
                file_name=f"{base_name}_spares.pdf",
                mime="application/pdf"
            )
