; --- Import Libraries ---
#SingleInstance Force
#Requires AutoHotkey v2.0
#Include _JXON.ahk ; Ensure you have the JSON library available in your script directory

; --- Configuration ---
SetTitleMatchMode 3 ; Exact match for window titles
SetControlDelay 0

MAXITEMPERPAGE := 14
SOWindowTitle := "ใบสั่งขาย"
WinspeedClass := "ahk_class FNWND380"
API_URL := "http://mathongapi.jpn.local/linebot/order?linebot_order_id="
DONE_INI := A_ScriptDir "\so_order_done.ini" ; remembers which pages were typed per order
CODE_CONTROL := "Edit12"   ; SO grid: product code cell
QTY_CONTROL := "PBEDIT807" ; SO grid: quantity cell (next are PBEDIT808 price, Edit18 discount)
MAX_ENTER := 8             ; max Enter presses to reach the next field before giving up
FIELD_DELAY := 30          ; ms settle time after the cursor moved to a new field
COLS := [30, 100, 150, 50] ; product list column widths (px): #, รหัสสินค้า, ชื่อสินค้า, จำนวน — window width follows
FIELD_TIMEOUT := 3000      ; ms max wait for WINSPEED to move the cursor after one Enter (product lookup)
global typing := false
global abortTyping := false
global pageGui := "" ; the open order window, if any

; Esc stops typing only while a page is being typed
#HotIf typing
Esc:: {
    global abortTyping := true
}
#HotIf

^l::
{
    ; Only one order window at a time: bring the open one to front (close it to load another order)
    if typing
        return
    if pageGui {
        pageGui.Show()
        return
    }

    if not WinExist(WinspeedClass) {
        MsgBox("โปรแกรม Winspeed ไม่ได้ถูกเปิดใช้งาน", "Error", "Iconx")
        return ; Exit the script if the window isn't found
    }

    SOHwnd := FindSOWindow()
    if (SOHwnd = '') {
        MsgBox("หน้าต่าง " . SOWindowTitle . " ไม่ได้ถูกเปิดใช้งาน", "Error", "Iconx")
        return
    }

    WinActivate(SOHwnd)

    ; Prefill from clipboard (backoffice "copy order id" button)
    clip := Trim(A_Clipboard)
    orderId := InputBox("กรุณาใส่ตัวเลข15-20หลัก", "ใส่ Order ID", "w200 h100", RegExMatch(clip, "^\d{15,20}$") ? clip :
        "")
    if (orderId.Result != "OK")
        return
    id := Trim(orderId.Value)
    if !RegExMatch(id, "^\d{15,20}$") {
        MsgBox("Order ID ไม่ถูกต้อง. กรุณาใส่ตัวเลข15-20หลัก", "Error", "Iconx")
        return
    }

    order := FetchOrder(id)
    if (order = "" || !ValidateOrder(order))
        return

    ShowPageGui(order, id, SOHwnd)
}

FindSOWindow() {
    for hwnd in WinGetList(WinspeedClass)
        if (WinGetTitle(hwnd) = SOWindowTitle)
            return hwnd
    return ""
}

FetchOrder(id) {
    loading := Gui("+AlwaysOnTop -SysMenu +ToolWindow", "Loading")
    loading.Add("Text", "w260", "กำลังดึงข้อมูล Order " . id . " ...")
    bar := loading.Add("Progress", "w260 h16 +0x8") ; PBS_MARQUEE
    SendMessage(0x40A, 1, 30, bar) ; PBM_SETMARQUEE: start animation
    loading.Show()
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.SetTimeouts(5000, 5000, 10000, 15000)
        whr.Open("GET", API_URL . id, true) ; async so the loading bar keeps animating
        whr.Send()
        while !whr.WaitForResponse(0)
            Sleep(50)
    } catch as e {
        loading.Destroy()
        MsgBox("เชื่อมต่อ API ไม่ได้`n" . e.Message, "Error", "Iconx")
        return ""
    }
    loading.Destroy()
    ; Decode raw bytes as UTF-8: API sends no charset, so ResponseText garbles Thai
    text := ""
    try {
        body := whr.ResponseBody
        pData := NumGet(ComObjValue(body), 8 + A_PtrSize, "Ptr") ; SAFEARRAY.pvData
        text := Trim(StrGet(pData, body.MaxIndex() + 1, "UTF-8"))
    }
    if (whr.Status != 200 || text = "" || text = "null") {
        MsgBox("ไม่พบ Order ID: " . id . " (HTTP " . whr.Status . ")", "Error", "Iconx")
        return ""
    }
    return jxon_load(&text)
}

ItemCode(el) {
    return (el.Get("good_id_correction", "") != "") ? el["good_id_correction"] : el.Get("good_code_correction", "")
}

; Stop before typing anything if the order has bad rows or is already linked to an SO
ValidateOrder(order) {
    items := order["items"]
    if (items.Length = 0) {
        MsgBox("No Data")
        return false
    }

    bad := ""
    for i, el in items {
        qty := el.Get("quantity", "")
        if (ItemCode(el) = "" || !IsNumber(qty) || qty <= 0)
            bad .= i . ". " . el.Get("good_code", "") . "  จำนวน=" . qty . "`n"
    }
    if (bad != "") {
        MsgBox("มีสินค้าไม่มีรหัส หรือจำนวนไม่ถูกต้อง กรุณาแก้ใน backoffice ก่อน:`n`n" . bad, "Error", "Iconx")
        return false
    }

    links := order.Get("link_so", [])
    if (links.Length > 0) {
        docs := ""
        for so in links
            docs .= so["docuno"] . " "
        if (MsgBox("Order นี้ผูกกับ SO แล้ว: " . docs . "`nต้องการคีย์อีกครั้งหรือไม่?", "Warning", "YesNo Icon!") !=
        "Yes")
            return false
    }
    return true
}

ShowPageGui(order, id, SOHwnd) {
    items := order["items"]
    totalPage := Ceil(items.Length / MAXITEMPERPAGE)

    g := Gui("+AlwaysOnTop", "Order " . id)
    g.MarginX := 15
    g.MarginY := 15
    ; Width = columns + 4px ListView border (no vertical scrollbar: page has max 14 rows = r14).
    ; Columns are pixels, GUI units are DPI-scaled.
    colsW := COLS[1] + COLS[2] + COLS[3] + COLS[4]
    W := Round((colsW + 4) * 96 / A_ScreenDPI)
    g.Add("Text", "w" . W, "ลูกค้า: " . order.Get("cust_code", "") . " " . order.Get("cust_name", "") . "    สถานะ: " .
    order.Get("status", "")
    . "`nอัพเดต: " . SubStr(order["rec_updated_when"], 1, 19) . "    สินค้า: " . items.Length . " รายการ / " .
    totalPage . " หน้า")
    pages := g.Add("ListView", "w" . W . " r5 -Multi NoSortHdr", ["หน้า", "รายการที่", "สถานะ"])
    btn := g.Add("Button", "w" . W . " h40 Default", "คีย์หน้านี้   (กด Esc เพื่อหยุดระหว่างคีย์)")
    toggle := g.Add("Button", "w" . W . " Left", "▼ รายการสินค้า")
    preview := g.Add("ListView", "w" . W . " r14 NoSortHdr", ["#", "รหัสสินค้า", "ชื่อสินค้า", "จำนวน"])

    ; Show at top-right of the work area (pixels, excludes taskbar)
    Place() {
        g.Show("AutoSize")
        WinGetPos(, , &gw, , g.Hwnd)
        MonitorGetWorkArea(, , &top, &right)
        WinMove(right - gw, top, , , g.Hwnd)
    }

    OnToggle(*) {
        show := !preview.Visible
        preview.Visible := show
        toggle.Text := (show ? "▼" : "▶") . " รายการสินค้า"
        Place()
    }

    Refresh(selectRow) {
        pages.Delete()
        loop totalPage {
            first := (A_Index - 1) * MAXITEMPERPAGE + 1
            last := Min(A_Index * MAXITEMPERPAGE, items.Length)
            done := IniRead(DONE_INI, id, "p" . A_Index, "")
            status := (done = "") ? "" : (SubStr(done, 1, 4) = "ERR ") ? "✗ คีย์ไม่สำเร็จ " . SubStr(done, 5) :
                "✓ คีย์แล้ว " . done
            pages.Add(A_Index = selectRow ? "Select Focus" : "", A_Index, first . " - " . last, status)
        }
        pages.ModifyCol(1, 40), pages.ModifyCol(2, 80), pages.ModifyCol(3, colsW - 40 - 80)
        ShowPreview(selectRow)
    }

    ShowPreview(page) {
        preview.Delete()
        first := (page - 1) * MAXITEMPERPAGE + 1
        last := Min(page * MAXITEMPERPAGE, items.Length)
        loop last - first + 1 {
            i := first + A_Index - 1
            el := items[i]
            preview.Add("", i, el.Get("winspeed_good_code", ItemCode(el)), el.Get("winspeed_good_name", ""), el[
                "quantity"])
        }
        loop 4
            preview.ModifyCol(A_Index, COLS[A_Index])
    }

    OnConfirm(*) {
        page := pages.GetNext()
        if (page = 0) {
            MsgBox("Please select page")
            return
        }
        prev := IniRead(DONE_INI, id, "p" . page, "")
        if (prev != "") {
            msg := (SubStr(prev, 1, 4) = "ERR ")
                ? "หน้า " . page . " คีย์ไม่สำเร็จครั้งก่อน (" . SubStr(prev, 5) .
                ")`nกรุณาลบรายการที่คีย์ค้างใน Winspeed ก่อน`n`nต้องการคีย์ใหม่หรือไม่?"
                : "หน้า " . page . " คีย์ไปแล้ว ต้องการคีย์ซ้ำหรือไม่?"
            if (MsgBox(msg, "Warning", "YesNo Icon! Owner" . g.Hwnd) != "Yes")
                return
        }

        ; WINSPEED may have been closed/reopened since Ctrl+L: old window handle is gone
        if !WinExist("ahk_id " . SOHwnd) {
            SOHwnd := FindSOWindow()
            if (SOHwnd = "") {
                MsgBox("ไม่พบหน้าต่าง " . SOWindowTitle . " กรุณาเปิด Winspeed และหน้า " . SOWindowTitle . " ก่อน", "Error",
                    "Iconx Owner" . g.Hwnd)
                return
            }
        }

        g.Hide()
        r := TypePage(SOHwnd, order["cust_id"], items, page)

        if (r.typed = r.expected) {
            IniWrite(FormatTime(, "dd/MM HH:mm"), DONE_INI, id, "p" . page)
            MsgBox("หน้า " . page . ": คีย์ครบ " . r.typed . " รายการ, จำนวนรวม " . r.qty
                . "`nกรุณาตรวจยอดใน Winspeed ก่อนบันทึก", "Done", "Iconi")
            Refresh(Min(page + 1, totalPage))
        } else {
            next := r.first + r.typed
            IniWrite("ERR " . FormatTime(, "dd/MM HH:mm") . " หยุดที่ #" . next, DONE_INI, id, "p" . page)
            MsgBox("หยุดคีย์ที่รายการ #" . next . " (" . ItemCode(items[next]) . ")"
            . "`nคีย์ไปแล้ว " . r.typed . "/" . r.expected . " รายการ"
            . "`nสาเหตุ: " . r.reason
            . "`n`nบันทึกสถานะหน้านี้เป็น คีย์ไม่สำเร็จ", "Stopped", "Icon!")
            Refresh(page)
        }
        Place()
    }

    pages.OnEvent("ItemSelect", (ctrl, row, selected) => selected ? ShowPreview(row) : "")
    btn.OnEvent("Click", OnConfirm)
    toggle.OnEvent("Click", OnToggle)
    g.OnEvent("Close", OnClose)
    Refresh(1)
    global pageGui := g
    Place()

    OnClose(*) {
        global pageGui := ""
        g.Destroy()
    }
}

FocusClass(SOHwnd) {
    try return ControlGetClassNN(ControlGetFocus(SOHwnd))
    return ""
}

; Press Enter until focus reaches target field. Works for both permission levels
; (after quantity: edit-permission users pass price/discount, no-permission users go straight to next row).
; Never presses Enter when SO window is not active (popup = stop, don't confirm it).
MoveTo(SOHwnd, target) {
    loop MAX_ENTER {
        if !WinActive("ahk_id " . SOHwnd)
            return false
        before := FocusClass(SOHwnd)
        Send("{Enter}")
        ; Wait until WINSPEED has processed this Enter (cursor left the field) before the next one.
        ; Pressing again too early queues Enters and skips rows (same ClassNN on every row).
        start := A_TickCount
        loop {
            now := FocusClass(SOHwnd)
            if (now != "" && now != before)
                break
            if (A_TickCount - start > FIELD_TIMEOUT)
                return false
            Sleep(20)
        }
        Sleep(FIELD_DELAY) ; let the new field settle
        if (FocusClass(SOHwnd) = target)
            return true
    }
    return false
}

; Types one page into the SO window. Stops if Esc is pressed, the SO window loses focus,
; or the cursor is not in the expected field.
TypePage(SOHwnd, cust_id, items, page) {
    global typing, abortTyping
    first := (page - 1) * MAXITEMPERPAGE + 1
    last := Min(page * MAXITEMPERPAGE, items.Length)
    r := { typed: 0, expected: last - first + 1, qty: 0, first: first, reason: "" }

    WinActivate(SOHwnd)
    if !WinWaitActive("ahk_id " . SOHwnd, , 2) {
        r.reason := "เปิดหน้าต่าง ใบสั่งขาย ไม่ได้"
        return r
    }

    typing := true
    abortTyping := false

    ControlFocus("Edit46", SOHwnd)
    ControlSend(cust_id . "{Enter}", "Edit46", SOHwnd) ; Customer ID
    Sleep(200)

    ControlClick "x588 y50", SOHwnd, , , , "NA" ; Run Bill Code
    Sleep(200)

    ControlFocus CODE_CONTROL, SOHwnd
    Sleep(500)

    loop r.expected {
        element := items[first + A_Index - 1]

        if abortTyping {
            r.reason := "กด Esc"
            break
        }
        if !WinActive("ahk_id " . SOHwnd) {
            r.reason := "หน้าต่าง ใบสั่งขาย ไม่ได้ active (มี popup หรือคลิกที่อื่น)"
            break
        }
        if (FocusClass(SOHwnd) != CODE_CONTROL) {
            r.reason := "เคอร์เซอร์ไม่อยู่ที่ช่องรหัสสินค้า"
            break
        }

        SendText(ItemCode(element))
        if !MoveTo(SOHwnd, QTY_CONTROL) {
            r.reason := "ไปไม่ถึงช่องจำนวน (ไม่พบสินค้า / มี popup)"
            break
        }

        SendText(element["quantity"])
        if !MoveTo(SOHwnd, CODE_CONTROL) {
            r.reason := "ไปไม่ถึงช่องรหัสสินค้าบรรทัดถัดไป"
            r.typed++ ; quantity of this item was typed
            r.qty += element["quantity"]
            break
        }

        r.typed++
        r.qty += element["quantity"]
    }

    typing := false
    return r
}
