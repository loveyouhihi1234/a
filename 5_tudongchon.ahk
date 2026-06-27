; =========== Tự động chọn để in ===========

isPaused := false

F8::
InputBox, n, Nhập số lần lặp, Nhập số lần muốn thực hiện:
if ErrorLevel
    return

Loop %n%
{
    ; Nếu đang tạm dừng thì chờ
    while (isPaused)
        Sleep, 10

    Send, {Space}

    ; Kiểm tra lại ngay sau khi gửi phím
    while (isPaused)
        Sleep, 10

    Sleep, 100

    while (isPaused)
        Sleep, 10

    Send, {Down}

    while (isPaused)
        Sleep, 10

    Sleep, 100
}
return


F9::
isPaused := !isPaused

if (isPaused)
    TrayTip, AutoHotkey, Đã tạm dừng, 1
else
    TrayTip, AutoHotkey, Tiếp tục chạy, 1
return


F10::
ExitApp
return