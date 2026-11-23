; --- Import Libraries ---
#SingleInstance Off ; a mathong-so:// link starts a 2nd copy; it hands the order ID to the running copy and exits
#Requires AutoHotkey v2.0
#Include _JXON.ahk ; Ensure you have the JSON library available in your script directory

; --- Configuration ---
SetTitleMatchMode 3 ; Exact match for window titles
SetControlDelay 0

MAXITEMPERPAGE := 14
SOWindowTitle := "ใบสั่งขาย"
WinspeedClass := "ahk_class FNWND380"
API_URL := "http://mathongapi.jpn.local/linebot/order?linebot_order_id="  ; mathong-system order
LIFF_API_URL := "http://mathongapi.jpn.local/firebase/liff/order?id="       ; LINE OA (LIFF) order
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
global pageOrderId := "" ; order ID shown in pageGui

; --- Startup: backoffice button opens mathong-so://<order id> ---
linkId := A_Args.Length ? ParseProtocolId(A_Args[1]) : ""
if (A_Args.Length && linkId = "")
    ExitApp ; malformed/hostile link: do nothing
OnMessage(0x4A, OnCopyData) ; WM_COPYDATA from a 2nd copy (registered first: a link may arrive right away)
; Named mutex decides which copy stays: a window scan alone lets 2 copies started together both exit.
DllCall("CreateMutex", "Ptr", 0, "Int", 0, "Str", "Local\mathong-so-order:" . StrReplace(A_ScriptFullPath, "\", "/"))
if (A_LastError = 183) { ; ERROR_ALREADY_EXISTS → another copy is the main one
    loop 30 { ; its main window may still be starting
        if (other := FindOtherCopy()) {
            SendOrderId(other, linkId) ; running copy opens it (or just shows its window)
            break
        }
        Sleep(100)
    }
    ExitApp
}
RegisterProtocol()
if (linkId != "")
    SetTimer(() => OpenOrder(linkId), -1)

; Esc stops typing only while a page is being typed
#HotIf typing
Esc:: {
    global abortTyping := true
}
#HotIf

^l:: OpenOrder()

; Strict: only 15-20 digits, so a hostile link can't pass anything else in
; Returns the order key: "<id>" = mathong-system (linebot) order, "liff/<id>" = LINE OA order
ParseProtocolId(arg) {
    return RegExMatch(arg, "i)^mathong-so:(?://)?((?:liff/)?\d{15,20})/?$", &m) ? StrLower(m[1]) : ""
}

IsOrderKey(key) {
    return RegExMatch(key, "^(liff/)?\d{15,20}$")
}

ProtocolCommand() {
    return '"' . A_ScriptFullPath . '" "%1"'
}

; Per-user (HKCU, no admin). Only the compiled exe registers, so testing the .ahk never hijacks the link.
; Re-registers when the exe was moved.
RegisterProtocol() {
    if !A_IsCompiled
        return
    key := "HKCU\Software\Classes\mathong-so"
    if (RegRead(key . "\shell\open\command", , "") = ProtocolCommand())
        return
    RegWrite("URL:Mathong SO", "REG_SZ", key)
    RegWrite("", "REG_SZ", key, "URL Protocol")
    RegWrite(ProtocolCommand(), "REG_SZ", key . "\shell\open\command")
}

; Another running copy of this script/exe. Matches the hidden main window title (starts with the
; script path), not the process path: uncompiled scripts all run as AutoHotkey64.exe.
FindOtherCopy() {
    found := 0
    DetectHiddenWindows(true)
    for hwnd in WinGetList("ahk_class AutoHotkey")
        if (hwnd != A_ScriptHwnd && IsCopyTitle(WinGetTitle(hwnd)))
            found := hwnd
    DetectHiddenWindows(false)
    return found
}

; Main window title: compiled exe = bare path; .ahk script = "path - AutoHotkey v2.x"
IsCopyTitle(title) {
    return title = A_ScriptFullPath || InStr(title, A_ScriptFullPath . " - ") = 1
}

SendOrderId(hwnd, id) {
    cds := Buffer(3 * A_PtrSize, 0) ; COPYDATASTRUCT
    NumPut("UPtr", 0, "UInt", (StrLen(id) + 1) * 2, cds)
    NumPut("Ptr", StrPtr(id), cds, 2 * A_PtrSize)
    DetectHiddenWindows(true) ; target is the running copy's hidden main window
    try SendMessage(0x4A, 0, cds, , "ahk_id " . hwnd, , , , 3000)
}

OnCopyData(wParam, lParam, msg, hwnd) {
    ptr := NumGet(lParam, 2 * A_PtrSize, "Ptr"), size := NumGet(lParam, A_PtrSize, "UInt")
    id := ptr ? StrGet(ptr, size // 2) : "" ; positive length = max chars, stops at the null
    if IsOrderKey(id) ; untrusted input: only a valid order key may reach OpenOrder/the API
        SetTimer(() => OpenOrder(id), -1) ; don't keep the sender waiting
    else if pageGui ; plain 2nd launch: just bring the order window to front, never the Ctrl+L InputBox
        pageGui.Show()
    return true
}

; id = "" → Ctrl+L flow (ask for the ID); id given → from a mathong-so:// link (skip the InputBox)
OpenOrder(id := "") {
    ; One open at a time: a 2nd link/Ctrl+L while the 1st is still fetching or showing a dialog is ignored
    static busy := false
    if (typing || busy)
        return
    busy := true
    try OpenOrderOnce(id)
    finally busy := false
}

OpenOrderOnce(id) {
    ; Only one order window at a time: bring the open one to front (close it to load another order)
    if pageGui {
        pageGui.Show()
        if (id != "" && id != pageOrderId)
            MsgBox("มี Order " . pageOrderId . " เปิดอยู่ กรุณาปิดก่อน", "Warning", "Icon! Owner" . pageGui.Hwnd)
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

    if (id = "") {
        id := AskOrderKey()
        if (id = "")
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

; Ctrl+L: pick the source first (mathong-system and LINE OA IDs look alike), then the ID.
; Returns the order key ("<id>" or "liff/<id>"), "" if cancelled.
AskOrderKey() {
    ; Prefill from clipboard (backoffice "copy order id" button)
    clip := Trim(A_Clipboard)
    d := Gui("+AlwaysOnTop -MinimizeBox", "ใส่ Order ID")
    d.Add("Text", , "Order มาจาก:")
    mathong := d.Add("Radio", "Checked", "Mathong System")
    d.Add("Radio", "x+20", "LINE OA (LIFF)")
    d.Add("Text", "xm", "Order ID (ตัวเลข 15-20 หลัก):")
    edit := d.Add("Edit", "xm w220", RegExMatch(clip, "^\d{15,20}$") ? clip : "")
    ok := d.Add("Button", "xm w105 Default", "OK")
    cancel := d.Add("Button", "x+10 w105", "Cancel")
    result := ""
    ok.OnEvent("Click", (*) => (result := (mathong.Value ? "" : "liff/") . Trim(edit.Value), d.Destroy()))
    cancel.OnEvent("Click", (*) => d.Destroy())
    d.OnEvent("Close", (*) => d.Destroy())
    d.OnEvent("Escape", (*) => d.Destroy())
    d.Show()
    WinWaitClose(d.Hwnd)
    if (result != "" && !IsOrderKey(result)) {
        MsgBox("Order ID ไม่ถูกต้อง. กรุณาใส่ตัวเลข15-20หลัก", "Error", "Iconx")
        return ""
    }
    return result
}

; LINE OA (LIFF) order → same shape as a mathong-system order, so checks/page window/typing are shared
NormalizeLiffOrder(raw) {
    items := []
    for el in raw.Get("items", [])
        items.Push(Map(
            "good_id_correction", el.Get("GoodID", ""),
            "good_code_correction", el.Get("GoodCode", ""),
            "good_code", el.Get("GoodCode", ""),
            "quantity", el.Get("itemQty", ""),
            "winspeed_good_code", el.Get("GoodCode", ""),
            "winspeed_good_name", el.Get("GoodMarketName", "")))
    return Map("items", items, "link_so", [],
        ; user selects the WINSPEED customer by hand (cust_id "" → TypePage skips it); show the LINE name instead
        "cust_id", "", "cust_code", "(LINE)", "cust_name", raw.Get("line_display_name", "") . " — เลือกลูกค้าใน Winspeed เอง",
        "status", raw.Get("isConfirm", false) ? "ยืนยันแล้ว" : "ยังไม่ยืนยัน",
        "rec_updated_when", raw.Get("updatedAt", ""))
}

FetchOrder(key) {
    isLiff := SubStr(key, 1, 5) = "liff/"
    id := isLiff ? SubStr(key, 6) : key
    url := isLiff ? LIFF_API_URL . id : API_URL . id
    loading := Gui("+AlwaysOnTop -SysMenu +ToolWindow", "Loading")
    loading.Add("Text", "w260", "กำลังดึงข้อมูล Order " . id . " ...")
    bar := loading.Add("Progress", "w260 h16 +0x8") ; PBS_MARQUEE
    SendMessage(0x40A, 1, 30, bar) ; PBM_SETMARQUEE: start animation
    loading.Show()
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.SetTimeouts(5000, 5000, 10000, 15000)
        whr.Open("GET", url, true) ; async so the loading bar keeps animating
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
    order := jxon_load(&text)
    return isLiff ? NormalizeLiffOrder(order) : order
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
    g.Add("Text", "w" . W, (SubStr(id, 1, 5) = "liff/" ? "[LINE OA] " : "[Mathong System] ") . "ลูกค้า: " . order.Get("cust_code", "") . " " . order.Get("cust_name", "") . "    สถานะ: " .
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
    global pageGui := g, pageOrderId := id
    Place()

    OnClose(*) {
        global pageGui, pageOrderId
        if (pageGui = g) ; an older window must not clear the newer one
            pageGui := "", pageOrderId := ""
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

    if (cust_id != "") { ; LINE OA order: user selects the customer in WINSPEED by hand
        ControlFocus("Edit46", SOHwnd)
        ControlSend(cust_id . "{Enter}", "Edit46", SOHwnd) ; Customer ID
        Sleep(200)
    }

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
