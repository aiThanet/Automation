; --- Import Libraries ---
#SingleInstance Force
#Requires AutoHotkey v2.0
#Include _JXON.ahk ; Ensure you have the JSON library available in your script directory

; --- Configuration ---
SetTitleMatchMode 3 ; Exact match for window titles
SetControlDelay 0

; IMPORTANT: Replace these placeholders with the actual values you find using "Active Window Info".
CustField := "Edit70" ; Customer ID field
SOWindowTitle := "ขายเชื่อ"
WinspeedClass := "ahk_class FNWND380"
global response_obj := {} ; Initialize response object

^l::
{
    if not WinExist(WinspeedClass) {
        MsgBox("โปรแกรม Winspeed ไม่ได้ถูกเปิดใช้งาน", "Error", "Iconx")
        return ; Exit the script if the window isn't found
    }

    HWNDs := WinGetList(WinspeedClass)
    SOHwnd := ''
    for hwnd in HWNDs {
        winspeedTitle := WinGetTitle(hwnd)
        if (winspeedTitle = SOWindowTitle) {
            WinActivate(hwnd)
            SOHwnd := hwnd
            break
        }
    }

    if (SOHwnd = '') {
        MsgBox("หน้าต่าง " . SOWindowTitle . " ไม่ได้ถูกเปิดใช้งาน", "Error", "Iconx")
        return
    }

    WinActivate(SOHwnd)

    SelectedFile := FileSelect(3, , "Open a file", "Excel (*.xlsx)")
    if (!selectedFile) {
        MsgBox("กรุณาเลือกไฟล์", "Error", "Iconx")
        return
    }

    Boundary := "---------------------------" . A_TickCount . Random(1000, 9999) ; Generate a unique boundary

    ; Construct the multipart/form-data body
    ; This is a simplified example. For very large files, reading in chunks might be better.
    FileContent := FileRead(SelectedFile, "RAW") ; Read file content as raw binary data

    ; Define the form fields
    ; The 'file' field is for the actual file content
    Body := "--" . Boundary . "`r`n"
        . "Content-Disposition: form-data; name='file'; filename='test`r`n'"
        . "Content-Type: application/octet-stream`r`n" ; Or specific MIME type if known (e.g., image/jpeg)
        . "`r`n"
        . FileContent . "`r`n"
        . "--" . Boundary . "--`r`n" ; Closing boundary

    URL := "https://mathongapi.jpn.local/cost/preupload?version=V1"

    HttpRequest := ComObject("WinHttp.WinHttpRequest.5.1")
    HttpRequest.Open("POST", URL, True) ; True for asynchronous, False for synchronous

    HttpRequest.SetRequestHeader("accept", "application/json")
    HttpRequest.SetRequestHeader("Content-Type", "multipart/form-data; boundary=" . Boundary)
    HttpRequest.SetRequestHeader("Content-Length", StrLen(Body)) ; Set content length

    ; Send the request
    HttpRequest.Send(Body)

    ; Wait for the response (for synchronous request, this isn't strictly needed)
    ; For asynchronous, you'd typically use HttpRequest.WaitForResponse() or an event handler
    ; Since this is a simple example, we'll block. In a real GUI, consider a separate thread.
    HttpRequest.WaitForResponse()

    StatusCode := HttpRequest.Status
    ResponseText := HttpRequest.ResponseText

}
