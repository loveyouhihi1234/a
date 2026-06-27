import os
import sys
import subprocess
import warnings
from datetime import datetime, timedelta

# Tắt còi báo động (cằn nhằn) của openpyxl cho giao diện chạy sạch đẹp
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# =====================================================================
# KHỐI LỆNH TỰ ĐỘNG TRANG BỊ VŨ KHÍ (PIP INSTALL)
# =====================================================================
def auto_install_vukhi(package_name, install_name=None):
    if install_name is None:
        install_name = package_name
    try:
        __import__(package_name)
    except ImportError:
        print(f"🛠 Tớ phát hiện máy cậu đang thiếu '{package_name}'...")
        print(f"⚙️ Đang tự động cài đặt luôn cho cậu. Đợi tớ 10 giây nhé!")
        subprocess.check_call([sys.executable, "-m", "pip", "install", install_name])
        print(f"✅ Trang bị {package_name} thành công!\n")

# Tự động tải các "đồ nghề" thiết yếu
auto_install_vukhi('docxtpl')
auto_install_vukhi('pandas')
auto_install_vukhi('docx', 'python-docx')
auto_install_vukhi('openpyxl')
auto_install_vukhi('jinja2')

import pandas as pd
import jinja2
from docxtpl import DocxTemplate
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# =====================================================================
# CẤU HÌNH ĐƯỜNG DẪN FILE
# =====================================================================
DATA_FILE = '2_thong_tin.xlsx'
WORD_TEMPLATE = '2_mau_hop_dong_xe_tai.docx'
OUTPUT_DIR = ''

# =====================================================================
# CÁC HÀM XỬ LÝ DỮ LIỆU THÔNG MINH
# =====================================================================

def xu_ly_cmt_cccd(val):
    """Radar chống nhiễu: Phục hồi số 0 bị mất cho CMT (9 số) và CCCD (12 số)."""
    if pd.isna(val) or str(val).strip() == "" or str(val).lower() == 'nan':
        return ""

    s = str(val).split('.')[0].strip()
    if len(s) == 8 or len(s) == 11:
        s = "0" + s
    return s

def xu_ly_ngay_excel(val):
    """Cỗ máy thời gian: Chuyển đổi số seri hoặc định dạng Date tự động về ngày chuẩn."""
    if pd.isna(val) or str(val).strip() == "" or str(val).lower() == 'nan':
        return ""

    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%d/%m/%Y")

    if isinstance(val, str):
        try:
            return pd.to_datetime(val).strftime("%d/%m/%Y")
        except:
            return str(val).strip()

    try:
        delta = timedelta(days=float(val))
        date_obj = datetime(1899, 12, 30) + delta
        return date_obj.strftime("%d/%m/%Y")
    except Exception:
        return str(val)

def format_number(val):
    try:
        if pd.isna(val) or val == "" or str(val).lower() == 'nan':
            return ""
        return str(int(float(val)))
    except ValueError:
        return str(val)

def format_money(val):
    try:
        if pd.isna(val) or val == "" or str(val).lower() == 'nan':
            return ""
        return f"{float(val):,.0f}".replace(',', '.')
    except ValueError:
        return str(val)

def lay_gia_tri(row, keys):
    """Hàm bọc lót siêu cấp BẤT TỬ: Tìm giá trị theo nhiều tên cột khác nhau"""
    # Ưu tiên 1: Tìm chính xác tên cột
    for k in keys:
        for col in row.index:
            if k == str(col).strip():
                return row[col]
    # Ưu tiên 2: Tìm tương đối (chỉ cần chứa từ khóa)
    for k in keys:
        for col in row.index:
            if k in str(col):
                return row[col]
    return ""

# =====================================================================
# BỘ CÔNG CỤ VẼ BẢNG
# =====================================================================

def format_cell(cell, text, bold=False, align='center'):
    cell.text = str(text)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)

    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == 'left':
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    for run in p.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = bold

def ke_vien_full(table):
    tblPr = table._tbl.tblPr
    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is None:
        tblBorders = OxmlElement('w:tblBorders')
        tblPr.append(tblBorders)
    else:
        tblBorders.clear()

    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
        tblBorders.append(border)

def autofit_to_window(table):
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')

# =====================================================================
# DÂY CHUYỀN SẢN XUẤT CHÍNH
# =====================================================================
def san_xuat_hop_dong():

    print("🚀 Khởi động dây chuyền tạo hợp đồng xe tải tự động ...")

    try:
        df_raw = pd.read_excel(DATA_FILE, engine='openpyxl', header=None)

        # =====================================================
        # MẮT THẦN RADAR FUZZY TỐI THƯỢNG (Chấm điểm AI)
        # =====================================================
        header_idx = -1
        max_score = 0

        # Danh sách các "từ khóa vàng" thường xuất hiện trên dòng tiêu đề
        keywords = ['so', 'stt', 'số', 'bks', 'biển', 'daidien', 'đại diện', 'xe', 'oto', 'ô tô', 'trongtai', 'tải', 'cmt', 'cccd', 'ngay', 'tien', 'tiền', 'thang', 'tháng', 'tong', 'tổng']

        for i, row in df_raw.iterrows():
            # Gom tất cả giá trị trong dòng thành 1 chuỗi chữ thường
            row_text = " ".join([str(val).lower() for val in row.values if pd.notna(val)])

            # Chấm điểm xem dòng này chứa bao nhiêu từ khóa
            score = sum(1 for k in keywords if k in row_text)

            # Nếu điểm cao nhất, đánh dấu đây là dòng tiêu đề (Header)
            if score > max_score:
                max_score = score
                header_idx = i

        # Phải có ít nhất 3 từ khóa trùng khớp thì mới an tâm đó là dòng tiêu đề
        if max_score < 3:
            print(f"\n❌ LỖI: Tớ quét lòi mắt cả {len(df_raw)} dòng mà không thấy dòng nào giống tiêu đề cả!")
            print(f"💡 Bí kíp kiểm tra:")
            print("1. Có thể dữ liệu của cậu đang nằm ở Sheet 2 hoặc Sheet 3 (chương trình mặc định đọc Sheet đầu tiên).")
            print("2. Hoặc file Excel này bị hỏng định dạng. Cậu thử Copy vùng dữ liệu dán sang một file Excel mới tinh rồi lưu lại xem sao nhé!")
            return

        print(f"🎯 Mắt thần AI đã khóa mục tiêu ở dòng số {header_idx + 1} (với {max_score} điểm trùng khớp)!")

        # Ép chuẩn cột
        df = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
        df.columns = df_raw.iloc[header_idx].astype(str).str.strip().str.lower()
    except FileNotFoundError:
        print(f"\n❌ Úi, tớ không tìm thấy file '{DATA_FILE}' đâu cả!")
        return
    except Exception as e:
        print(f"\n❌ Lỗi đọc file Data rồi cậu ơi: {e}")
        return
    # =========================================================
    # TẠO TÊN FOLDER ĐỘNG THEO SỐ HỢP ĐỒNG
    # =========================================================

    try:
        danh_sach_so = []

        for _, row in df.iterrows():
            so_raw = lay_gia_tri(row, ['so', 'stt', 'số'])

            try:
                so_int = int(float(so_raw))
                danh_sach_so.append(so_int)
            except:
                pass

        danh_sach_so = sorted(danh_sach_so)

        so_dau = str(danh_sach_so[0]).zfill(2)
        so_cuoi = str(danh_sach_so[-1]).zfill(2)

        output_dir = f"Hop_Dong_Thue_Xe_Tai_So_{so_dau}_Den_{so_cuoi}"

    except:
        output_dir = "Hop_Dong_Thue_Xe_Tai"

    # Tạo folder
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    print("✅ Đang nạp nhiên liệu, bắt đầu tạo hợp đồng...\n")
    thanh_cong = 0

    for index, row in df.iterrows():
        try:
            # Rút trích dữ liệu an toàn bằng hàm bọc lót (Bất chấp tên cột bị sai lệch đôi chút)
            bks_raw = str(lay_gia_tri(row, ['bks', 'biển'])).strip()
            daidien_raw = str(lay_gia_tri(row, ['daidien', 'đại diện', 'tên']))

            # Nếu dòng không có BKS hoặc Đại diện, lướt qua luôn vì đó là dòng trống
            if not bks_raw or bks_raw.lower() == 'nan':
                continue

            doc = DocxTemplate(WORD_TEMPLATE)

            # Chuẩn hóa biến
            so_raw = lay_gia_tri(row, ['so', 'stt', 'số'])
            try:
                so_str = str(int(float(so_raw))).zfill(2)
            except:
                so_str = "00"

            cmt = xu_ly_cmt_cccd(lay_gia_tri(row, ['cmt', 'cccd', 'cmnd']))
            ngaycap = xu_ly_ngay_excel(lay_gia_tri(row, ['ngaycap', 'ngày cấp', 'capngay']))
            trongtai = format_number(lay_gia_tri(row, ['trongtai', 'trọng tải', 'tải']))

            tien_thang = format_money(lay_gia_tri(row, ['tien', 'tiền', 'giá']))
            so_thang = format_number(lay_gia_tri(row, ['thang', 'tháng']))
            tong_tien = format_money(lay_gia_tri(row, ['tong', 'tổng', 'thành']))

            # =========================================================
            # VẼ BẢNG TIỀN THUÊ XE
            # =========================================================
            bang_subdoc = doc.new_subdoc()
            bang = bang_subdoc.add_table(rows=3, cols=5)
            ke_vien_full(bang)
            autofit_to_window(bang)

            hdr = bang.rows[0].cells
            format_cell(hdr[0], 'Từ tháng', bold=True, align='center')
            format_cell(hdr[1], 'Đến tháng', bold=True, align='center')
            format_cell(hdr[2], 'Số tháng', bold=True, align='center')
            format_cell(hdr[3], 'Giá trị thuê / tháng', bold=True, align='center')
            format_cell(hdr[4], 'Thành tiền', bold=True, align='center')

            r1 = bang.rows[1].cells
            format_cell(r1[0], '01', align='center')
            format_cell(r1[1], '12', align='center')
            format_cell(r1[2], so_thang, align='center')
            format_cell(r1[3], tien_thang, align='right')
            format_cell(r1[4], tong_tien, align='right')

            r2 = bang.rows[2].cells
            r2[0].merge(r2[3])
            format_cell(r2[0], 'Tổng cộng:', bold=True, align='center')
            format_cell(r2[4], tong_tien, bold=True, align='right')

            # =========================================================
            # ĐÓNG GÓI VÀO WORD
            # =========================================================
            context = {
                'so': so_str,
                'daidien': daidien_raw if daidien_raw.lower() != 'nan' else "",
                'diachi': str(lay_gia_tri(row, ['diachi', 'địa chỉ'])) if str(lay_gia_tri(row, ['diachi', 'địa chỉ'])).lower() != 'nan' else "",
                'cmt': cmt,
                # Tớ nạp cả 2 biến này để chống bẫy, Word cậu gõ kiểu gì cũng ăn!
                'ngaycap': ngaycap,
                'capngay': ngaycap,
                'congan': str(lay_gia_tri(row, ['congan', 'công an', 'nơi'])) if str(lay_gia_tri(row, ['congan', 'công an', 'nơi'])).lower() != 'nan' else "",
                'bks': bks_raw,
                'oto': str(lay_gia_tri(row, ['oto', 'ô tô', 'loại'])) if str(lay_gia_tri(row, ['oto', 'ô tô', 'loại'])).lower() != 'nan' else "",
                'trongtai': trongtai,
                'bang_thanh_toan': bang_subdoc
            }

            # Lưu file
            file_name = f"{so_str}_{bks_raw}.docx"
            output_path = os.path.join(output_dir, file_name)

            doc.render(context)
            doc.save(output_path)

            print(f"✅ Đã xuất: {file_name} (Chủ xe: {context['daidien']})")
            thanh_cong += 1

        except jinja2.exceptions.TemplateSyntaxError as e:
            print(f"\n❌ BÁO ĐỘNG ĐỎ: Lỗi gõ thẻ bên trong file Word '{WORD_TEMPLATE}'!")
            print(f"Chi tiết: {e}")
            print(f"💡 Cậu mở file Word ra, kiểm tra lại các ngoặc kép xem có bị cách trống không nhé (VD: '{{ diachi }}').")
            break

        except Exception as e:
            print(f"⚠️ Bỏ qua 1 dòng do lỗi: {e}")

    if thanh_cong > 0:
        print(f"\n🎉 XONG! Đã tạo thành công {thanh_cong} bản hợp đồng hoàn hảo. Mở '{output_dir}' để kiểm tra!")
    else:
        print(f"\n😢 Máy chạy xong mà không ra hợp đồng nào. Cậu check lại data xem có trống không nhé!")

if __name__ == "__main__":
    san_xuat_hop_dong()
