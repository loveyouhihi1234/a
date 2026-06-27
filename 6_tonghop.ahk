#NoEnv
#SingleInstance Force
SendMode Input
SetBatchLines, -1
SetWorkingDir %A_ScriptDir%

Global CurrentOutput := "NK"
Global IsUIVisible := false

; === 1. BẢNG ĐIỀU KHIỂN CHÍNH (THIẾT KẾ GRID 3x3 CHUYÊN NGHIỆP) ===
Gui, Main:+AlwaysOnTop -MaximizeBox -MinimizeBox
Gui, Main:Font, s11, Segoe UI

; Tạo khung viền bọc ngoài cho gọn
Gui, Main:Add, GroupBox, x15 y10 w320 h135, Chọn lệnh nhanh (Phím Right)

; --- Hàng 1 ---
Gui, Main:Add, Radio, vChoice gUpdateChoice x30 y35 w80 Checked, NK
Gui, Main:Add, Radio, gUpdateChoice x130 y35 w80, r00
Gui, Main:Add, Radio, gUpdateChoice x230 y35 w80, r08

; --- Hàng 2 ---
Gui, Main:Add, Radio, gUpdateChoice x30 y65 w80, r10
Gui, Main:Add, Radio, gUpdateChoice x130 y65 w80, v00
Gui, Main:Add, Radio, gUpdateChoice x230 y65 w80, v08

; --- Hàng 3 ---
Gui, Main:Add, Radio, gUpdateChoice x30 y95 w80, v10
Gui, Main:Add, Radio, gUpdateChoice x130 y95 w80, XK
Gui, Main:Add, Radio, gUpdateChoice x230 y95 w80, Khác:

; Ô nhập liệu Tùy chỉnh (Được căn lề phẳng phiu)
Gui, Main:Add, Text, x15 y160 w80, Tùy chỉnh:
Gui, Main:Font, s11 bold, Consolas
Gui, Main:Add, Edit, vCustomStr gUpdateChoice x95 y155 w240 c1A73E8, xuatkho

; Dòng hướng dẫn mờ ở đáy (Đã sửa F1 thành F8)
Gui, Main:Font, s9 italic norm, Segoe UI
Gui, Main:Add, Text, x15 y195 w320 Center cGray, ( Bấm F8 hoặc Enter để ẩn bảng này đi )

Gui, Main:Add, Button, Hidden Default gHideMain, OK


; === 2. BẢNG HUD LƠ LỬNG (THIẾT KẾ HIỆN ĐẠI CÓ DẢI LED) ===
Gui, OSD:+AlwaysOnTop -Caption +ToolWindow +LastFound +E0x20
Gui, OSD:Color, 1E1E1E ; Đen nhám hiện đại

; Dải "LED" dọc ở cạnh trái (Progress bar giả làm viền)
Gui, OSD:Add, Progress, x0 y0 w4 h65 vOsdLine c43D08A, 100

; Text Trạng thái
Gui, OSD:Font, s10 bold c43D08A, Segoe UI
Gui, OSD:Add, Text, x15 y10 w200 vOsdStatus, ▶ ĐANG HOẠT ĐỘNG

; Text Giá trị
Gui, OSD:Font, s10 norm cWhite, Segoe UI
Gui, OSD:Add, Text, x15 y35 w200 vOsdValue, Lệnh: [ NK ]

; Chỉnh độ mờ
WinSet, Transparent, 225

; Đặt vị trí
SysGet, WorkArea, MonitorWorkArea
OSD_X := WorkAreaLeft + 20
OSD_Y := WorkAreaBottom - 85
Gui, OSD:Show, NoActivate x%OSD_X% y%OSD_Y% w220 h65, HUD


return ; === KẾT THÚC KHỞI TẠO ===


; === 3. XỬ LÝ DỮ LIỆU ===
UpdateChoice:
Gui, Main:Submit, NoHide
Switch Choice
{
    Case 1: CurrentOutput := "NK"
    Case 2: CurrentOutput := "r00"
    Case 3: CurrentOutput := "r08"
    Case 4: CurrentOutput := "r10"
    Case 5: CurrentOutput := "v00"
    Case 6: CurrentOutput := "v08"
    Case 7: CurrentOutput := "v10"
    Case 8: CurrentOutput := "XK"
    Case 9: CurrentOutput := CustomStr
}
GuiControl, OSD:, OsdValue, Lệnh: [ %CurrentOutput% ]
return

HideMain:
Gui, Main:Hide
IsUIVisible := false
return


; === 4. HỆ THỐNG PHÍM TẮT ĐIỀU KHIỂN (ĐÃ SỬA F8, F9, F10) ===

F8::
Suspend, Permit
if (IsUIVisible) {
    Gui, Main:Hide
    IsUIVisible := false
} else {
    Gui, Main:Show, AutoSize xCenter yCenter, Bảng Điều Khiển Nhanh
    IsUIVisible := true
}
return

F9::
Suspend
if (A_IsSuspended) {
    ; Chuyển dải LED và chữ sang Đỏ
    GuiControl, OSD:+cFF4C4C, OsdLine
    Gui, OSD:Font, cFF4C4C
    GuiControl, OSD:Font, OsdStatus
    ; Đã cập nhật text báo F9 thay vì F2
    GuiControl, OSD:, OsdStatus, ⏸ TẠM DỪNG (F9) 
} else {
    ; Chuyển lại sang Xanh
    GuiControl, OSD:+c43D08A, OsdLine
    Gui, OSD:Font, c43D08A
    GuiControl, OSD:Font, OsdStatus
    GuiControl, OSD:, OsdStatus, ▶ ĐANG HOẠT ĐỘNG
}
return

F10::
Suspend, Permit
ExitApp


; === 5. PHÍM TẮT THỰC THI (RIGHT ARROW) ===
*Right::
Gui, Main:Submit, NoHide 
if (Choice == 9) {
    CurrentOutput := CustomStr
}
SendInput %CurrentOutput%
return