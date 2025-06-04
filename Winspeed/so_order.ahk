; --- Import Libraries ---
#Requires AutoHotkey v2.0
#Include _JXON.ahk ; Ensure you have the JSON library available in your script directory

; --- Configuration ---
SetTitleMatchMode 3 ; Exact match for window titles
SetControlDelay 0

; IMPORTANT: Replace these placeholders with the actual values you find using "Active Window Info".
DDL := ""
MAXITEMPERPAGE := 14
SOWindowTitle := "ใบสั่งขาย"
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

    orderId := InputBox("กรุณาใส่ตัวเลข18หลัก", "ใส่ Order ID", "w200 h100")
    ; Check if orderId is an 18-digit number
    if !RegExMatch(orderId.Value, "^\d{18}$") {
        MsgBox("Order ID ไม่ถูกต้อง. กรุณาใส่ตัวเลข18หลัก", "Error", "Iconx")
        return
    }

    url := "http://mathongapi.jpn.local/linebot/order?linebot_order_id=" . orderId.Value
    whr := ComObject("WinHttp.WinHttpRequest.5.1")
    whr.Open("GET", url)
    whr.Send()

    response := whr.ResponseText
    response_obj := jxon_load(&response)

    items := response_obj["items"]
    cust_id := response_obj["cust_id"]
    TotalItems := items.Length
    TotalPage := Ceil(TotalItems / MAXITEMPERPAGE)

    if (TotalPage = 0) {
        MsgBox("No Data")
        return
    }

    ListPage := ""
    loop TotalPage {
        ListPage .= "Page" A_Index
        if (A_Index < TotalPage)
            ListPage .= "|"
    }

    MsgBox("Order ID: " . response_obj["linebot_order_id"] . "`nอัพเดต: " . SubStr(response_obj["rec_updated_when"], 1,
        19) . "`nจำนวนหน้าทั้งหมด: " . TotalPage . " หน้า")

    ; === GUI Setup ===
    gui1 := Gui("+AlwaysOnTop", "Select Page")
    gui1.MarginX := 20
    gui1.MarginY := 20

    ddlControl := gui1.Add("DropDownList", "w200 h60 r10 vDDL", StrSplit(ListPage, "|"))
    btnConfirm := gui1.Add("Button", "w200 h40", "Confirm")

    ddlControl.OnEvent("Change", Submit_All)
    btnConfirm.OnEvent("Click", (*) => Do_AddItem(gui1, SOHwnd, items, cust_id))

    gui1.Show()

    return
}

; === Event Handlers ===
Submit_All(ctrl, info) {
    global DDL := ctrl.Text
}

Do_AddItem(guiWindow, SOHwnd, items, cust_id) {

    if (DDL = "") {
        MsgBox "Please select page"
        return
    }

    ; Extract page number (assumes format "Page X")
    ppage := SubStr(DDL, 5) ; skip "Page " (5 chars)
    minItem := ((ppage - 1) * MAXITEMPERPAGE) + 1
    maxItem := (ppage * MAXITEMPERPAGE)

    guiWindow.Destroy()

    ControlFocus("Edit46", SOHwnd)
    ControlSend(cust_id . "{Enter}", "Edit46", SOHwnd) ; Customer ID
    Sleep(200)

    ControlClick "x588 y50", SOHwnd, , , , "NA" ; Run Bill Code
    Sleep(200)

    ControlFocus "Edit12", SOHwnd
    Sleep(500)

    loop items.Length {
        index := A_Index
        if (index >= minItem && index <= maxItem) {
            element := items[index]

            valueToSend := (element["good_id_correction"] != "") ? element["good_id_correction"] : element[
                "good_code_correction"]
            SendText(valueToSend)

            loop 5
                Send("{Enter}")

            SendText(element["quantity"])

            loop 1
                Send("{Enter}")
        }
    }
}
