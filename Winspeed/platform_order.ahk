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
orders := []
orderIdx := 0
SOHwnd := ""

^l::
{
    global orders, orderIdx, SOHwnd

    checkWinActive()

    SelectedFile := FileSelect(3, , "Open a file", "Excel (*.xlsx)")
    if (!selectedFile) {
        MsgBox("กรุณาเลือกไฟล์", "Error", "Iconx")
        return
    }

    ; Prepare the HTTP request
    apiUrl := "http://localhost:3100/ecom/upload/bigseller"

    ; Create COM object for HTTP request
    req := ComObject("WinHttp.WinHttpRequest.5.1")
    req.Open("POST", apiUrl, false)

    objParam := { file: [SelectedFile] }
    CreateFormData(&PostData, &hdr_ContentType, objParam)

    req.SetRequestHeader("Content-Type", hdr_ContentType)
    req.Send(PostData)

    response := req.ResponseBody
    stream := ComObject("ADODB.Stream")
    stream.Type := 1  ; Binary
    stream.Open()
    stream.Write(response)
    stream.Position := 0
    stream.Type := 2  ; Text
    stream.Charset := "utf-8"
    resText := stream.ReadText()
    stream.Close()

    response_obj := jxon_load(&resText)
    total_orders := response_obj.Length
    orders := response_obj
    orderIdx := 1

    MsgBox("จำนวนออเดอร์ทั้งหมด: " . total_orders . "`n`n กด Ctrl+N เพื่อเริ่มใส่ข้อมูลออเดอร์...")
}

^n::
{
    global orders, orderIdx, SOHwnd

    checkWinActive()

    if (orders.Length = 0) {
        MsgBox("ไม่มีออเดอร์ให้ดำเนินการ", "Error", "Iconx")
        return
    }
    if (orderIdx > orders.Length) {
        MsgBox("ดำเนินการเสร็จสิ้นแล้ว", "Info", "Iconx")
        return
    }

    order := orders[orderIdx]
    total_orders := orders.Length

    order_id := order["order_id"]
    cust_id := order["cust_id"]
    platform := order["platform"]
    store := order["store"]
    items := order["items"]

    MsgBox("Order ที่ " . orderIdx . " / " . total_orders . "`nOrder ID: " . order_id . "`n`nแพลตฟอร์ม: " .
        platform .
        "`nร้าน: " . store . "`n`nจำนวนรายการสินค้า: " . items.Length)

    ControlFocus("Edit70", SOHwnd)
    ControlSend(cust_id . "{Enter}", "Edit70", SOHwnd) ; Customer ID
    Sleep(200)

    ControlClick "x588 y50", SOHwnd, , , , "NA" ; Run Bill Code
    Sleep(200)

    ControlFocus "Edit10", SOHwnd
    Sleep(500)

    loop items.Length {
        index := A_Index
        item := items[index]

        valueToSend := item["goodcode"]

        if (valueToSend = "") {
            MsgBox("รายการสินค้า: " . item["sku"] . "`nจำนวน: " . item["quantity"] .
                "`n`nกด F5 เพื่อใส่ข้อมูลต่อไป...", "ไม่พบรหัสสินค้า", "Iconx")
            KeyWait "F5", "D"
        } else {
            SendText(valueToSend)
            Sleep(100)

            loop 5 {
                Send("{Enter}")
                Sleep(100)
            }

            SendText(item["qty"])
            Sleep(100)

            loop 4
                Send("{Enter}")
        }
    }

    ControlClick "x160 y404", SOHwnd, , , , "NA" ; Click Description
    Sleep(200)

    ControlFocus("Edit38", SOHwnd)
    ControlSendText(platform . ":" . order_id, "Edit38", SOHwnd) ; Customer ID
    Sleep(200)

    ControlClick "x40 y404", SOHwnd, , , , "NA" ; Click Description
    Sleep(200)

    MsgBox("ดำเนินการสำเร็จ `n`nกด Ctrl+N เพื่อใส่ข้อมูลออเดอร์ถัดไป...")

    orderIdx := orderIdx + 1

}

checkWinActive() {
    global SOHwnd
    if not WinExist(WinspeedClass) {
        MsgBox("โปรแกรม Winspeed ไม่ได้ถูกเปิดใช้งาน", "Error", "Iconx")
        return ; Exit the script if the window isn't found
    }

    HWNDs := WinGetList(WinspeedClass)
    SOHwnd := ""
    for hwnd in HWNDs {
        winspeedTitle := WinGetTitle(hwnd)
        if (winspeedTitle = SOWindowTitle) {
            WinActivate(hwnd)
            SOHwnd := hwnd
            break
        }
    }

    if (SOHwnd = "") {
        MsgBox("หน้าต่าง " . SOWindowTitle . " ไม่ได้ถูกเปิดใช้งาน", "Error", "Iconx")
        return
    }

    WinActivate(SOHwnd)
}

class CreateFormData {

    __New(&retData, &retHeader, objParam) {

        local CRLF := "`r`n", i, k, v, str, pvData
        ; Create a random Boundary
        local Boundary := CreateFormData.RandomBoundary()
        local BoundaryLine := "------------------------------" . Boundary

        ; Create an IStream backed with movable memory.
        hData := DllCall("GlobalAlloc", "uint", 0x2, "uptr", 0, "ptr")
        DllCall("ole32\CreateStreamOnHGlobal", "ptr", hData, "int", False, "ptr*", &pStream := 0, "uint")
        CreateFormData.pStream := pStream

        ; Loop input paramters
        for k, v in objParam.OwnProps() {
            if IsObject(v) {
                for i, FileName in v {
                    str := BoundaryLine . CRLF
                        . 'Content-Disposition: form-data; name="' . k . '"; filename="' . FileName . '"' . CRLF
                        . 'Content-Type: ' . CreateFormData.MimeType(FileName) . CRLF . CRLF

                    CreateFormData.StrPutUTF8(str)
                    CreateFormData.LoadFromFile(Filename)
                    CreateFormData.StrPutUTF8(CRLF)

                }
            } else {
                str := BoundaryLine . CRLF
                    . 'Content-Disposition: form-data; name="' . k '"' . CRLF . CRLF
                    . v . CRLF
                CreateFormData.StrPutUTF8(str)
            }
        }

        CreateFormData.StrPutUTF8(BoundaryLine . "--" . CRLF)

        CreateFormData.pStream := ObjRelease(pStream) ; Should be 0.
        pData := DllCall("GlobalLock", "ptr", hData, "ptr")
        size := DllCall("GlobalSize", "ptr", pData, "uptr")

        ; Create a bytearray and copy data in to it.
        retData := ComObjArray(0x11, size) ; Create SAFEARRAY = VT_ARRAY|VT_UI1
        pvData := NumGet(ComObjValue(retData), 8 + A_PtrSize, "ptr")
        DllCall("RtlMoveMemory", "Ptr", pvData, "Ptr", pData, "Ptr", size)

        DllCall("GlobalUnlock", "ptr", hData)
        DllCall("GlobalFree", "Ptr", hData, "Ptr")                   ; free global memory

        retHeader := "multipart/form-data; boundary=----------------------------" . Boundary
    }

    static StrPutUTF8(str) {
        buf := Buffer(StrPut(str, "UTF-8") - 1) ; remove null terminator
        StrPut(str, buf, buf.size, "UTF-8")
        DllCall("shlwapi\IStream_Write", "ptr", CreateFormData.pStream, "ptr", buf.Ptr, "uint", buf.Size, "uint")
    }

    static LoadFromFile(filepath) {
        DllCall("shlwapi\SHCreateStreamOnFileEx"
            , "wstr", filepath
            , "uint", 0x0             ; STGM_READ
            , "uint", 0x80            ; FILE_ATTRIBUTE_NORMAL
            , "int", False            ; fCreate is ignored when STGM_CREATE is set.
            , "ptr", 0               ; pstmTemplate (reserved)
            , "ptr*", &pFileStream := 0
            , "uint")
        DllCall("shlwapi\IStream_Size", "ptr", pFileStream, "uint64*", &size := 0, "uint")
        DllCall("shlwapi\IStream_Copy", "ptr", pFileStream, "ptr", CreateFormData.pStream, "uint", size, "uint")
        ObjRelease(pFileStream)
    }

    static RandomBoundary() {
        str := "0|1|2|3|4|5|6|7|8|9|a|b|c|d|e|f|g|h|i|j|k|l|m|n|o|p|q|r|s|t|u|v|w|x|y|z"
        Sort str, 'D| Random'
        str := StrReplace(str, "|")
        return SubStr(str, 1, 12)
    }

    static MimeType(FileName) {
        n := FileOpen(FileName, "r").ReadUInt()
        return (n = 0x474E5089) ? "iกmage/png"
            : (n = 0x38464947) ? "image/gif"
                : (n & 0xFFFF = 0x4D42) ? "image/bmp"
                    : (n & 0xFFFF = 0xD8FF) ? "image/jpeg"
                        : (n & 0xFFFF = 0x4949) ? "image/tiff"
                            : (n & 0xFFFF = 0x4D4D) ? "image/tiff"
                                : "application/octet-stream"
    }

}
