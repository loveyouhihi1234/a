#NoEnv
#SingleInstance Force
SetWorkingDir %A_ScriptDir%

; VŨ KHÍ TỐI THƯỢNG: Tốc độ gõ phím siêu rùa (Nghỉ 75ms giữa mỗi phím, giữ phím 75ms)
SetKeyDelay, 75, 75 

; --- TẠO BẢNG THEO DÕI (GUI OVERLAY - GUI 1) ---
Gui, 1:+AlwaysOnTop -Caption +Border +ToolWindow +LastFound
WinSet, Transparent, 220
Gui, 1:Color, 1A1A1A
Gui, 1:Font, s11 c00FF00 bold, Segoe UI
Gui, 1:Add, Text, w300 vGuiStatus, ⏸️ Đang ngủ... (Nhấn F8 để gọi Bot)
Gui, 1:Font, s10 cWhite norm
Gui, 1:Add, Text, w300 vGuiInfo1, HĐ: --- | CTy: --- | Số CT: ---
Gui, 1:Add, Text, w300 vGuiInfo2, Tiến độ: ---
Gui, 1:Add, Text, w300 vGuiInfo3, KH: --- | Số HĐ: ---

yPos := A_ScreenHeight - 226
Gui, 1:Show, x02 y%yPos% NoActivate, Tracker
return 

; ====================================================================
; BẤM F8 -> HIỆN BẢNG CẤU HÌNH PRO VIP
; ====================================================================
F8::
FilePath := A_ScriptDir . "\4_data_hoadon.xlsx"

if !FileExist(FilePath) {
    MsgBox, 16, Lỗi, Không tìm thấy file 4_data_hoadon.xlsx ở cùng thư mục!
    return
}

Gui, 2:New, +AlwaysOnTop -MinimizeBox -MaximizeBox, Bảng Cấu Hình Hệ Thống Bot
Gui, 2:Color, White
Gui, 2:Font, s10, Segoe UI

; --- KHU VỰC 1: THÔNG TIN CHUNG ---
Gui, 2:Add, GroupBox, x10 y10 w160 h70, 1. Loại Công Ty
Gui, 2:Add, Radio, x20 y35 vOptCongTy Checked, C.Ty May
Gui, 2:Add, Radio, x95 y35, Khác

Gui, 2:Add, GroupBox, x180 y10 w180 h70, 2. Chế Độ Chạy
Gui, 2:Add, Radio, x190 y35 vOptLoaiHD Checked, MUA VÀO
Gui, 2:Add, Radio, x280 y35, BÁN RA

Gui, 2:Add, GroupBox, x370 y10 w120 h70, 3. Số c.từ
Gui, 2:Add, Text, x380 y35, Cắt mấy số:
Gui, 2:Add, Edit, x455 y30 w25 vKyTuSoCtu Number, 4

; --- KHU VỰC 2: TÙY CHỈNH MUA VÀO ---
Gui, 2:Add, GroupBox, x10 y90 w480 h90 cBlue, TÙY CHỈNH HÓA ĐƠN MUA VÀO (Bỏ trống = Mặc định)
Gui, 2:Add, Text, x20 y115, TK Chi phí xe:
Gui, 2:Add, Radio, x115 y115 vOptTaiKhoanMua Checked, Dùng 1541
Gui, 2:Add, Radio, x205 y115, Dùng 154

Gui, 2:Add, Text, x20 y150, Ép TK Nợ:
Gui, 2:Add, Edit, x90 y145 w80 vTKN_Mua
Gui, 2:Add, Text, x200 y150, Ép TK Có:
Gui, 2:Add, Edit, x260 y145 w80 vTKC_Mua
Gui, 2:Add, Text, x350 y145 cGray, (>5Tr tự sang 331)

; --- KHU VỰC 3: TÙY CHỈNH BÁN RA ---
Gui, 2:Add, GroupBox, x10 y190 w480 h60 cRed, 4. TÀI KHOẢN HÓA ĐƠN BÁN RA
Gui, 2:Add, Text, x20 y220, TK Nợ (Khách nợ):
Gui, 2:Add, Edit, x140 y215 w80 vTKN_Ban, 131   
Gui, 2:Add, Text, x240 y220, TK Có (Doanh thu):
Gui, 2:Add, Edit, x360 y215 w80 vTKC_Ban, 5112  

; --- NÚT CHẠY ---
Gui, 2:Font, bold s11
Gui, 2:Add, Button, x10 y265 w480 h40 gBatDauChay, 🚀 LƯU CẤU HÌNH VÀ BẮT ĐẦU CHẠY
Gui, 2:Show, w500 h315
return

; ====================================================================
; NHẤN NÚT CHẠY -> BẮT ĐẦU VÀO VIỆC
; ====================================================================
BatDauChay:
Gui, 2:Submit  
Gui, 2:Destroy 

; Xử lý dữ liệu cấu hình
LoaiCongTy := (OptCongTy == 1) ? "May" : "Khac"
LoaiHoaDon := (OptLoaiHD == 1) ? "MuaVao" : "BanRa"
TK_ChiPhi := (OptTaiKhoanMua == 1) ? "1541" : "154"

if (KyTuSoCtu = "" or KyTuSoCtu < 1) 
    KyTuSoCtu := 4 

GuiControl, 1:, GuiInfo1, HĐ: %LoaiHoaDon% | CTy: %LoaiCongTy% | CT: %KyTuSoCtu% số
GuiControl, 1:+cFFFF00, GuiStatus 
GuiControl, 1:, GuiStatus, ⏳ Đang nạp dữ liệu vào RAM...

; --- KIẾN TRÚC MỚI: ĐỌC 1 LẦN VÀO RAM RỒI ĐÓNG EXCEL LUÔN ---
xl := ComObjCreate("Excel.Application")
xl.Visible := False
wb := xl.Workbooks.Open(FilePath)
ws := wb.Sheets(1)

lastRow := ws.Cells(ws.Rows.Count, 1).End(-4162).Row 
TotalHoadon := lastRow - 1

TongTienTheoNgay := {}
DanhSachHoaDon := [] ; Mảng lưu toàn bộ dữ liệu

Loop, % TotalHoadon
{
    r := A_Index + 1
    ng := ws.Cells(r, 1).Text
    sohd := ws.Cells(r, 2).Text
    dt := ws.Cells(r, 3).Text
    
    tien_val := ws.Cells(r, 4).Value
    thue_val := ws.Cells(r, 6).Value
    tien_nhap := (tien_val != "") ? Floor(tien_val) : 0
    thue_nhap := (thue_val != "") ? Floor(thue_val) : 0
    
    vat := ws.Cells(r, 5).Text
    dg := ws.Cells(r, 7).Text

    ; Cất vào mảng trí nhớ của bot
    DanhSachHoaDon.Push({Ngay: ng, SoHD: sohd, DoiTuong: dt, Tien: tien_nhap, Thue: thue_nhap, VAT: vat, DienGiai: dg})
    
    ; Cộng dồn tổng tiền
    if (LoaiHoaDon == "MuaVao") {
        key := ng . "_" . dt
        if !TongTienTheoNgay.HasKey(key)
            TongTienTheoNgay[key] := 0
        TongTienTheoNgay[key] += (tien_nhap + thue_nhap)
    }
}

; Đóng Excel ngay lập tức để giải phóng bộ nhớ và chống lỗi COM
wb.Close(False)
xl.Quit()

; --- ĐẾM NGƯỢC THÔNG MINH ---
GuiControl, 1:+cFF8C00, GuiStatus 
GuiControl, 1:, GuiStatus, ⚠️ Click chuột vào 'Ngày c.từ' ngay! (3s)
Sleep, 1000
GuiControl, 1:, GuiStatus, ⚠️ Click chuột vào 'Ngày c.từ' ngay! (2s)
Sleep, 1000
GuiControl, 1:, GuiStatus, ⚠️ Click chuột vào 'Ngày c.từ' ngay! (1s)
Sleep, 1000

GuiControl, 1:+c00FF00, GuiStatus 
GuiControl, 1:, GuiStatus, ▶️ ĐANG GÕ DỮ LIỆU...

; --- KỊCH BẢN GÕ PHÍM (LẤY TỪ RAM) ---
Loop, % TotalHoadon
{
    ; Bốc dữ liệu từ trong mảng ra
    hd := DanhSachHoaDon[A_Index]

    ngay_hd := hd.Ngay
    so_hd := hd.SoHD
    doi_tuong := hd.DoiTuong
    tien := hd.Tien
    thue := hd.Thue
    vat := hd.VAT
    dien_giai_unicode := hd.DienGiai

    ; Xử lý đệm Số c.từ linh hoạt
    pad_str := "0000000000"
    so_ctu := pad_str . so_hd
    StringRight, so_ctu, so_ctu, %KyTuSoCtu%

    StringUpper, vat, vat
    StringLower, dg_lower, dien_giai_unicode 
    muc_thue := SubStr(vat, 2, 2)

    ; --- PHÂN TÍCH LOGIC TÀI KHOẢN ---
    if (LoaiHoaDon == "BanRa") {
        ; LOGIC BÁN RA
        tk_no := (TKN_Ban != "") ? TKN_Ban : "131"
        tk_co := (TKC_Ban != "") ? TKC_Ban : "5112"
        so_enter_sau_tien := 3
    } else {
        ; LOGIC MUA VÀO
        so_enter_sau_tien := 4
        
        ; Xét TK Có
        if (TKC_Mua != "") {
            tk_co := TKC_Mua
        } else {
            tong_trong_ngay := TongTienTheoNgay.HasKey(ngay_hd . "_" . doi_tuong) ? TongTienTheoNgay[ngay_hd . "_" . doi_tuong] : 0
            tk_co := (tong_trong_ngay > 5000000) ? "331" : "1111"
        }

        ; Xét TK Nợ
        if (TKN_Mua != "") {
            tk_no := TKN_Mua
        } else {
            if (InStr(dg_lower, "phí ngân hàng") or InStr(dg_lower, "phi ngan hang")) {
                tk_no := "642"
            } else if (InStr(dg_lower, "xăng")) {
                tk_no := (LoaiCongTy = "May") ? TK_ChiPhi : "642"
            } else if (InStr(dg_lower, "dầu")) {
                tk_no := TK_ChiPhi
            } else {
                tk_no := TK_ChiPhi 
            }
        }
    }

    ; Xác định Enter sau ô Thuế
    if (muc_thue = "00") {
        so_enter_sau_thue := 6
    } else {
        so_enter_sau_thue := 8
    }

    GuiControl, 1:, GuiInfo2, Tiến độ: %A_Index% / %TotalHoadon%
    GuiControl, 1:, GuiInfo3, KH: %doi_tuong% | Số HĐ: %so_hd%
    
    ; --- BẮT ĐẦU GÕ BÀN PHÍM ---
    Send, %ngay_hd%{Enter}
    Sleep, 500 
    
    Send, %so_ctu%{Enter 3}
    Sleep, 500 
    
    Send, %doi_tuong%{Enter 4}
    Sleep, 500 
    
    ; Copy & Paste Diễn giải
    Clipboard := "" 
    Clipboard := dien_giai_unicode
    Sleep, 100 
    Send, ^v
    Sleep, 200
    Send, {Enter}
    Sleep, 500
    
    Send, %tk_no%{Enter}
    Sleep, 500
    
    Send, %tk_co%{Enter}
    Sleep, 500
    
    Send, %tien%{Enter %so_enter_sau_tien%}
    Sleep, 500
    
    Send, %vat%{Enter}
    Sleep, 500
    
    Send, %thue%{Enter %so_enter_sau_thue%}
    Sleep, 500
    
    Send, %so_hd%{Enter}
    Sleep, 500
    
    Send, %ngay_hd%{Enter 3}
    
    Sleep, 2500 
}

; --- KẾT THÚC ---
GuiControl, 1:, GuiStatus, ✅ HOÀN THÀNH TOÀN BỘ!
MsgBox, 64, Xong Việc, Đã xử lý êm ru %TotalHoadon% hóa đơn. Mời sếp check lại Bravo!
return

; ====================================================================
; ĐIỀU KHIỂN F9 / F10
; ====================================================================
F9::
Pause, Toggle, 1
if A_IsPaused
{
    GuiControl, 1:+cFF0000, GuiStatus 
    GuiControl, 1:, GuiStatus, ⏸️ ĐÃ TẠM DỪNG (F9 để tiếp tục)
}
else
{
    GuiControl, 1:+c00FF00, GuiStatus
    GuiControl, 1:, GuiStatus, ▶️ ĐANG GÕ DỮ LIỆU...
}
return

F10::
ExitApp
