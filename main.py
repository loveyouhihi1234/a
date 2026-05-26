import os
import io
import zipfile
import tempfile
import pandas as pd
import jinja2
from datetime import datetime, timedelta
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from docxtpl import DocxTemplate
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = FastAPI(title="Hệ thống Tự động hóa Hợp đồng")

# Mở khóa bảo mật CORS để web từ GitHub Pages có thể gọi sang Server Render
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =====================================================================
# CÁC HÀM XỬ LÝ DỮ LIỆU CHUNG (Dùng chung cho cả Xe tải và Gạo)
# =====================================================================
def xu_ly_cmt_cccd(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).lower() == 'nan': return ""
    s = str(val).split('.')[0].strip()
    if len(s) == 8 or len(s) == 11: s = "0" + s
    return s

def xu_ly_ngay_excel(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).lower() == 'nan': return ""
    if isinstance(val, (pd.Timestamp, datetime)): return val.strftime("%d/%m/%Y")
    if isinstance(val, str):
        try: return pd.to_datetime(val).strftime("%d/%m/%Y")
        except: return str(val).strip()
    try:
        delta = timedelta(days=float(val))
        date_obj = datetime(1899, 12, 30) + delta
        return date_obj.strftime("%d/%m/%Y")
    except Exception: return str(val)

def format_number(val):
    try:
        if pd.isna(val) or val == "" or str(val).lower() == 'nan': return ""
        return f"{float(val):,.0f}".replace(',', '.')
    except ValueError: return str(val)

def xu_ly_mst(val):
    if pd.isna(val) or str(val).strip() == "": return ""
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    if s.isdigit():
        if len(s) <= 10: return s.zfill(10)
        elif len(s) <= 13:
            s_13 = s.zfill(13)
            return f"{s_13[:10]}-{s_13[10:]}"
    return s

def lay_gia_tri(row, keys):
    for k in keys:
        for col in row.index:
            if k == str(col).strip(): return row[col]
    for k in keys:
        for col in row.index:
            if k in str(col): return row[col]
    return ""

def format_cell(cell, text, bold=False, align='center'):
    cell.text = str(text)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right': p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == 'left': p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = bold

def autofit_to_window(table):
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')

def tao_file_zip(thu_muc_nguon, ten_file_zip_ao):
    """Nén toàn bộ thư mục thành file zip trong RAM rồi bắn trả về trình duyệt"""
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(thu_muc_nguon):
            for file in files:
                file_path = os.path.join(root, file)
                zf.write(file_path, os.path.relpath(file_path, thu_muc_nguon))
    memory_file.seek(0)
    return StreamingResponse(
        memory_file, 
        media_type="application/zip", 
        headers={"Content-Disposition": f"attachment; filename={ten_file_zip_ao}"}
    )

@app.get("/")
def read_root():
    return {"message": "Hệ thống Backend Khởi tạo Hợp đồng Tự động đang hoạt động cực mượt!"}

# =====================================================================
# API 1: XUẤT HỢP ĐỒNG THUÊ XE TẢI
# =====================================================================
def ke_vien_full(table):
    tblPr = table._tbl.tblPr
    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is None:
        tblBorders = OxmlElement('w:tblBorders')
        tblPr.append(tblBorders)
    else: tblBorders.clear()
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
        tblBorders.append(border)

@app.post("/api/hop-dong-xe-tai")
async def tao_hop_dong_xe_tai(excel_file: UploadFile = File(...), word_template: UploadFile = File(...)):
    # Đọc file nhị phân trực tiếp trên RAM do Web bắn sang
    excel_data = await excel_file.read()
    word_data = await word_template.read()
    
    df_raw = pd.read_excel(io.BytesIO(excel_data), engine='openpyxl', header=None)
    
    # Mắt thần AI tìm dòng tiêu đề
    header_idx = -1
    max_score = 0
    keywords = ['so', 'stt', 'số', 'bks', 'biển', 'daidien', 'đại diện', 'xe', 'oto', 'ô tô', 'trongtai', 'tải', 'cmt', 'cccd', 'ngay', 'tien', 'tiền', 'thang', 'tháng', 'tong', 'tổng']
    
    for i, row in df_raw.iterrows():
        row_text = " ".join([str(val).lower() for val in row.values if pd.notna(val)])
        score = sum(1 for k in keywords if k in row_text)
        if score > max_score:
            max_score = score
            header_idx = i

    if max_score < 3: 
        return {"error": "Không tìm thấy dòng tiêu đề hợp lệ trong file Excel."}
    
    df = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
    df.columns = df_raw.iloc[header_idx].astype(str).str.strip().str.lower()

    # Tạo thư mục tạm (tự động xóa sau khi xong việc để tiết kiệm RAM)
    with tempfile.TemporaryDirectory() as temp_dir:
        for index, row in df.iterrows():
            bks_raw = str(lay_gia_tri(row, ['bks', 'biển'])).strip()
            daidien_raw = str(lay_gia_tri(row, ['daidien', 'đại diện', 'tên']))
            if not bks_raw or bks_raw.lower() == 'nan': continue

            doc = DocxTemplate(io.BytesIO(word_data))
            
            so_raw = lay_gia_tri(row, ['so', 'stt', 'số'])
            try: so_str = str(int(float(so_raw))).zfill(2)
            except: so_str = "00"

            cmt = xu_ly_cmt_cccd(lay_gia_tri(row, ['cmt', 'cccd', 'cmnd']))
            ngaycap = xu_ly_ngay_excel(lay_gia_tri(row, ['ngaycap', 'ngày cấp', 'capngay']))
            trongtai = format_number(lay_gia_tri(row, ['trongtai', 'trọng tải', 'tải']))
            tien_thang = format_number(lay_gia_tri(row, ['tien', 'tiền', 'giá']))
            so_thang = format_number(lay_gia_tri(row, ['thang', 'tháng']))
            tong_tien = format_number(lay_gia_tri(row, ['tong', 'tổng', 'thành']))

            bang_subdoc = doc.new_subdoc()
            bang = bang_subdoc.add_table(rows=3, cols=5)
            ke_vien_full(bang)
            autofit_to_window(bang)

            hdr = bang.rows[0].cells
            format_cell(hdr[0], 'Từ tháng', bold=True)
            format_cell(hdr[1], 'Đến tháng', bold=True)
            format_cell(hdr[2], 'Số tháng', bold=True)
            format_cell(hdr[3], 'Giá trị thuê / tháng', bold=True)
            format_cell(hdr[4], 'Thành tiền', bold=True)

            r1 = bang.rows[1].cells
            format_cell(r1[0], '01')
            format_cell(r1[1], '12')
            format_cell(r1[2], so_thang)
            format_cell(r1[3], tien_thang, align='right')
            format_cell(r1[4], tong_tien, align='right')

            r2 = bang.rows[2].cells
            r2[0].merge(r2[3])
            format_cell(r2[0], 'Tổng cộng:', bold=True)
            format_cell(r2[4], tong_tien, bold=True, align='right')

            context = {
                'so': so_str,
                'daidien': daidien_raw if daidien_raw.lower() != 'nan' else "",
                'diachi': str(lay_gia_tri(row, ['diachi', 'địa chỉ'])) if str(lay_gia_tri(row, ['diachi', 'địa chỉ'])).lower() != 'nan' else "",
                'cmt': cmt,
                'ngaycap': ngaycap,
                'capngay': ngaycap,
                'congan': str(lay_gia_tri(row, ['congan', 'công an', 'nơi'])) if str(lay_gia_tri(row, ['congan', 'công an', 'nơi'])).lower() != 'nan' else "",
                'bks': bks_raw,
                'oto': str(lay_gia_tri(row, ['oto', 'ô tô', 'loại'])) if str(lay_gia_tri(row, ['oto', 'ô tô', 'loại'])).lower() != 'nan' else "",
                'trongtai': trongtai,
                'bang_thanh_toan': bang_subdoc
            }

            file_name = f"{so_str}_{bks_raw}.docx"
            output_path = os.path.join(temp_dir, file_name)
            doc.render(context)
            doc.save(output_path)

        # Gói gém lại gửi về Frontend
        return tao_file_zip(temp_dir, "HopDongXeTai_ThanhPham.zip")

# =====================================================================
# API 2: XUẤT HỢP ĐỒNG GẠO TỔNG HỢP THEO NHÓM
# =====================================================================
def ve_vien_ngang(table):
    tblPr = table._tbl.tblPr
    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is not None: tblBorders.clear()
    else:
        tblBorders = OxmlElement('w:tblBorders')
        tblPr.append(tblBorders)
    for border_name in ['insideH']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4') 
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
        tblBorders.append(border)

@app.post("/api/hop-dong-gao")
async def tao_hop_dong_gao(excel_file: UploadFile = File(...), word_template: UploadFile = File(...)):
    excel_data = await excel_file.read()
    word_data = await word_template.read()
    
    df = pd.read_excel(io.BytesIO(excel_data))
    df = df.dropna(subset=['sohd']) 
    nhom_hop_dong = df.groupby('sohd')

    with tempfile.TemporaryDirectory() as temp_dir:
        for sohd, group in nhom_hop_dong:
            doc = DocxTemplate(io.BytesIO(word_data))
            
            first_row = group.iloc[0]
            so_hd_str = str(sohd).replace('.0', '').zfill(2)
            ngay = str(first_row['ngay']).replace('.0', '').zfill(2)
            thang = str(first_row['thang']).replace('.0', '').zfill(2)
            nam = str(first_row['nam']).replace('.0', '')
            
            try:
                ngay_ky_obj = datetime(int(nam), int(thang), int(ngay))
                ngay_giao_obj = ngay_ky_obj + timedelta(days=2)
                giaohang = ngay_giao_obj.strftime("%d")
                thanggiaohang = ngay_giao_obj.strftime("%m")
                namgiaohang = ngay_giao_obj.strftime("%Y")
            except Exception:
                giaohang, thanggiaohang, namgiaohang = "...", "...", "..."
                
            so_mat_hang = len(group)

            bang_1_subdoc = doc.new_subdoc()
            bang_1 = bang_1_subdoc.add_table(rows=so_mat_hang + 1, cols=4)
            ve_vien_ngang(bang_1)
            autofit_to_window(bang_1)
            
            hdr1 = bang_1.rows[0].cells
            format_cell(hdr1[0], 'Tên hàng', bold=True)
            format_cell(hdr1[1], 'Số lượng', bold=True)
            format_cell(hdr1[2], 'Đơn giá', bold=True)
            format_cell(hdr1[3], 'Mặt hàng', bold=True)
            
            for i, (_, row) in enumerate(group.iterrows()):
                r1 = bang_1.rows[i + 1].cells
                format_cell(r1[0], str(row['tenhang']) if pd.notna(row['tenhang']) else "")
                format_cell(r1[1], f"{format_number(row['soluong'])} kg")
                format_cell(r1[2], f"{format_number(row['dongia'])} đ/kg", align='right') 
                format_cell(r1[3], str(row['mathang']) if pd.notna(row['mathang']) else "")

            bang_subdoc = doc.new_subdoc()
            bang = bang_subdoc.add_table(rows=so_mat_hang + 3, cols=5)
            bang.style = 'Table Grid'
            autofit_to_window(bang) 
            
            hdr_cells = bang.rows[0].cells
            format_cell(hdr_cells[0], 'STT', bold=True)
            format_cell(hdr_cells[1], 'Tên hàng hóa', bold=True)
            format_cell(hdr_cells[2], 'Số lượng (Kg)', bold=True)
            format_cell(hdr_cells[3], 'Đơn Giá', bold=True)
            format_cell(hdr_cells[4], 'Thành Tiền', bold=True)
            
            for i, (_, row) in enumerate(group.iterrows()):
                row_cells = bang.rows[i + 1].cells
                format_cell(row_cells[0], str(i + 1))
                format_cell(row_cells[1], str(row['tenhang']) if pd.notna(row['tenhang']) else "")
                format_cell(row_cells[2], format_number(row['soluong']))
                format_cell(row_cells[3], format_number(row['dongia']), align='right') 
                format_cell(row_cells[4], format_number(row['thanhtien']), align='right') 
                
            thue_row = bang.rows[so_mat_hang + 1].cells
            thue_row[0].merge(thue_row[3]) 
            thue_txt = f"Thuế Suất GTGT: {first_row['thuesuat'] if pd.notna(first_row['thuesuat']) else ''}"
            format_cell(thue_row[0], thue_txt, bold=True)
            format_cell(thue_row[4], format_number(first_row['thue']), bold=True, align='right') 
            
            tong_row = bang.rows[so_mat_hang + 2].cells
            tong_row[0].merge(tong_row[3])
            format_cell(tong_row[0], "Tổng tiền thanh toán", bold=True)
            format_cell(tong_row[4], format_number(first_row['tongtien']), bold=True, align='right') 
                
            context = {
                'sohd': so_hd_str,
                'ngay': ngay, 'thang': thang, 'nam': nam,
                'giaohang': giaohang, 'thanggiaohang': thanggiaohang, 'namgiaohang': namgiaohang,
                'benmua': str(first_row['benmua']) if pd.notna(first_row['benmua']) else "",
                'diachi': str(first_row['diachi']) if pd.notna(first_row['diachi']) else "",
                'mst': xu_ly_mst(first_row['mst']),
                'gioitinh': str(first_row['gioitinh']) if pd.notna(first_row['gioitinh']) else "",
                'daidien': str(first_row['daidien']) if pd.notna(first_row['daidien']) else "",
                'bang_dieu_1': bang_1_subdoc,
                'bang_hang_hoa': bang_subdoc
            }

            file_name = f"{so_hd_str}. Hợp đồng gạo - {ngay}{thang}.docx"
            output_path = os.path.join(temp_dir, file_name)
            doc.render(context)
            doc.save(output_path)

        # Gói gém lại gửi về Frontend
        return tao_file_zip(temp_dir, "HopDongGao_ThanhPham.zip")
