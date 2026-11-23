; --- Import Libraries ---
#SingleInstance Force
#Requires AutoHotkey v2.0
#Include _JXON.ahk ; Ensure you have the JSON library available in your script directory
#Include _theme.ahk ; colours, ThemeGui, AddPrimary, Enter → orange button

; --- Configuration ---
SetTitleMatchMode 3 ; Exact match for window titles
SetControlDelay 0

; IMPORTANT: Replace these placeholders with the actual values you find using "Active Window Info".
CustField := "Edit70" ; Customer ID field
SOWindowTitle := "ขายเชื่อ"
WinspeedClass := "ahk_class FNWND380"
orders := []
SOHwnd := ""
CODE_CONTROL := "Edit10" ; ขายเชื่อ grid: product code cell (same ClassNN on every row)
CODE_TO_QTY := 5         ; Enter presses from product code to quantity
MAX_ENTER := 8           ; max Enter presses from quantity back to the next row's code before giving up
FIELD_DELAY := 30        ; ms settle time after the cursor moved to a new field
FIELD_TIMEOUT := 3000    ; ms max wait for WINSPEED to move the cursor after one Enter (product lookup)

apiUrl := "http://mathongapi.jpn.local/ecom/upload/bigseller"
; apiUrl := "http://localhost:3100/ecom/upload/bigseller"
ITEMS_URL := "http://mathongapi.jpn.local/item/all" ; WINSPEED product list (code check + search in the edit dialog)
global typing := false
global abortTyping := false
global orderStatus := []   ; per order: "" = not typed, "HH:mm" = typed, "ERR ..." = stopped
global ordersGui := ""     ; the order list window, if open
global keySelected := ""   ; its "key the selected order" action (Ctrl+N runs it too)
global uploadName := ""    ; uploaded file name (window title)
global products := []      ; [code, name] of every usable WINSPEED product, for search
global validCodes := Map() ; code → name; empty = list not loaded (then only an empty code counts as wrong)
global editDlgHwnd := 0    ; the open edit dialog, if any
global editDown := ""      ; its ↓ action: from the code box into the search results

; Esc stops typing only while an order is being typed
#HotIf typing
Esc:: {
    global abortTyping := true
}
#HotIf

; Edit dialog: ↓ in the code box jumps into the search results
#HotIf editDlgHwnd && WinActive("ahk_id " . editDlgHwnd)
Down:: editDown()
#HotIf

FocusClass(SOHwnd) {
    try return ControlGetClassNN(ControlGetFocus(SOHwnd))
    return ""
}

; One Enter, then wait until WINSPEED has moved the cursor to another field before the next one:
; pressing again too early queues Enters and skips rows. Never presses Enter when ขายเชื่อ is not
; active (WINSPEED popup = stop, don't confirm it).
NextField(SOHwnd) {
    if (abortTyping || !WinActive("ahk_id " . SOHwnd))
        return false
    before := FocusClass(SOHwnd)
    Send("{Enter}")
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
    return true
}

; Enter until focus reaches target. Works for both permission levels
; (after quantity: edit-permission users pass price/discount, no-permission users go straight to next row).
MoveTo(SOHwnd, target) {
    loop MAX_ENTER {
        if !NextField(SOHwnd)
            return false
        if (FocusClass(SOHwnd) = target)
            return true
    }
    return false
}

^l::
{
    global orders, SOHwnd

    if (typing || !checkWinActive())
        return

    SelectedFile := FileSelect(3, , "Open a file", "Excel (*.xlsx)")
    if (!selectedFile) {
        MsgBox("กรุณาเลือกไฟล์", "Error", "Iconx")
        return
    }

    ; Create COM object for HTTP request
    req := ComObject("WinHttp.WinHttpRequest.5.1")
    req.SetTimeouts(5000, 5000, 30000, 60000) ; resolve, connect, send, receive (ms)
    req.Open("POST", apiUrl, true) ; async so the loading bar keeps animating

    objParam := { file: [SelectedFile] }
    CreateFormData(&PostData, &hdr_ContentType, objParam)

    req.SetRequestHeader("Content-Type", hdr_ContentType)
    global T := Theme() ; pick up a Windows light/dark switch
    SplitPath(SelectedFile, &fileName)
    loading := ShowLoading("กำลังอัปโหลด " . fileName . " ...")
    try {
        req.Send(PostData)
        while !req.WaitForResponse(0)
            Sleep(50)
    } catch as e {
        loading.Destroy()
        MsgBox("เชื่อมต่อ API ไม่ได้`n" . e.Message, "Error", "Iconx")
        return
    }
    loading.Destroy()
    if (req.Status != 200) {
        MsgBox("API ตอบกลับผิดพลาด (HTTP " . req.Status . ")", "Error", "Iconx")
        return
    }

    resText := Utf8Body(req)

    try response_obj := jxon_load(&resText)
    catch {
        MsgBox("อ่านข้อมูลจาก API ไม่ได้ (ไม่ใช่ JSON)", "Error", "Iconx")
        return
    }
    if (response_obj.Length = 0) {
        MsgBox("ไม่มีออเดอร์ในไฟล์", "Error", "Iconx")
        return
    }
    orders := response_obj
    global orderStatus := [], uploadName := fileName
    loading := ShowLoading("กำลังโหลดรหัสสินค้าจาก WINSPEED ...")
    loaded := FetchProducts()
    loading.Destroy()
    if !loaded
        MsgBox("โหลดรายการรหัสสินค้าไม่ได้`nแก้รหัสได้ แต่ตรวจไม่ได้ว่ามีใน WINSPEED หรือไม่", "Warning", "Icon!")
    loop orders.Length
        orderStatus.Push("")
    ShowOrdersGui()
}

; Ctrl+N = the window's orange button (key the selected order); reopens the list if it was closed
^n::
{
    if typing
        return
    if (orders.Length = 0) {
        MsgBox("ไม่มีออเดอร์ให้ดำเนินการ กด Ctrl+L เพื่อเลือกไฟล์", "Error", "Iconx")
        return
    }
    if !ordersGui
        ShowOrdersGui()
    else
        keySelected()
}

LogisticMsg(order) {
    delivery_method := order["delivery_method"]
    return "Logistic: " . order["logistic"] . (delivery_method = "" ? "" : (" (" . delivery_method . ")"))
}

; Response body as UTF-8 (the API sends no charset, so ResponseText garbles Thai)
Utf8Body(req) {
    stream := ComObject("ADODB.Stream")
    stream.Type := 1  ; Binary
    stream.Open()
    stream.Write(req.ResponseBody)
    stream.Position := 0
    stream.Type := 2  ; Text
    stream.Charset := "utf-8"
    text := stream.ReadText()
    stream.Close()
    return text
}

; WINSPEED products from mathongapi /item/all: every code except Inactive "I" = the same codes the bigseller
; upload accepts. Regex scan instead of a JSON parse: 2.5 MB of flat records, only 3 fields needed.
FetchProducts() {
    global products := [], validCodes := Map()
    req := ComObject("WinHttp.WinHttpRequest.5.1")
    req.SetTimeouts(5000, 5000, 10000, 30000)
    try {
        req.Open("GET", ITEMS_URL, true) ; async so the loading bar keeps animating
        req.Send()
        while !req.WaitForResponse(0)
            Sleep(50)
    } catch
        return false
    if (req.Status != 200)
        return false
    text := Utf8Body(req)
    pos := 1
    while (pos := RegExMatch(text, '"GoodCode": "([^"]*)"[^{}]*?"GoodName1": (?:"((?:[^"\\]|\\.)*)"|null)'
        . '[^{}]*?"Inactive": (?:"(\w*)"|null)', &m, pos)) {
        pos += m.Len
        if (m[3] = "I")
            continue
        code := StrUpper(m[1])
        products.Push([code, StrReplace(StrReplace(m[2], '\"', '"'), "\\", "\")])
        validCodes[code] := products[products.Length][2]
    }
    return products.Length > 0
}

; Wrong code = empty, or not in WINSPEED (when the product list loaded)
IsBadCode(code) {
    return code = "" || (validCodes.Count && !validCodes.Has(StrUpper(code)))
}

BadItems(order) {
    n := 0
    for item in order["items"]
        n += IsBadCode(item["goodcode"])
    return n
}

; Types one order into ขายเชื่อ. Stops on Esc, a popup / other window, or the cursor not where expected.
; Returns { stopped: item # (0 = all typed), reason }
KeyOrder(order) {
    global typing, abortTyping
    items := order["items"]
    r := { stopped: 0, reason: "" }

    WinActivate(SOHwnd)
    if !WinWaitActive("ahk_id " . SOHwnd, , 2) {
        r.stopped := 1, r.reason := "เปิดหน้าต่าง " . SOWindowTitle . " ไม่ได้"
        return r
    }

    typing := true
    abortTyping := false
    ; ControlFocus("Edit70", SOHwnd)
    ControlClick "x100 y70", SOHwnd, , , , "NA"
    Sleep(200)
    SendInput(order["cust_id"] . "{Enter}")
    Sleep(200)

    SendInput("ON")
    Sleep(200)

    ControlFocus CODE_CONTROL, SOHwnd
    Sleep(500)

    loop items.Length {
        item := items[A_Index]
        r.stopped := A_Index ; until this row is complete
        if abortTyping
            break
        if !WinActive("ahk_id " . SOHwnd) {
            r.reason := "หน้าต่าง " . SOWindowTitle . " ไม่ได้ active (มี popup หรือคลิกที่อื่น)"
            break
        }
        if (FocusClass(SOHwnd) != CODE_CONTROL) {
            r.reason := "เคอร์เซอร์ไม่อยู่ที่ช่องรหัสสินค้า"
            break
        }

        SendText(item["goodcode"])
        moved := true
        loop CODE_TO_QTY
            if !(moved := NextField(SOHwnd))
                break
        if !moved {
            r.reason := "ไปไม่ถึงช่องจำนวน (ไม่พบสินค้า / มี popup)"
            break
        }

        SendText(item["qty"])
        if !MoveTo(SOHwnd, CODE_CONTROL) {
            r.reason := "ไปไม่ถึงช่องรหัสสินค้าบรรทัดถัดไป"
            break
        }
        r.stopped := 0
    }
    if abortTyping
        r.reason := "กด Esc"
    typing := false
    if r.stopped
        return r

    ControlClick "x160 y404", SOHwnd, , , , "NA" ; Click Description
    Sleep(500)

    ControlFocus("Edit38", SOHwnd)
    Sleep(200)

    SendText(order["platform"] . ":" . order["order_id"])
    Sleep(200)

    Send("{Enter}")
    Sleep(200)

    SendText(LogisticMsg(order))
    Sleep(200)

    Send("{Enter}")
    Sleep(200)

    SendText("Confirm time: " . order["confirm_time"])
    Sleep(200)

    ControlClick "x40 y404", SOHwnd, , , , "NA" ; Click Description
    Sleep(200)
    return r
}

checkWinActive() {
    global SOHwnd
    if not WinExist(WinspeedClass) {
        MsgBox("โปรแกรม Winspeed ไม่ได้ถูกเปิดใช้งาน", "Error", "Iconx")
        return false ; Exit the script if the window isn't found
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
        return false
    }

    WinActivate(SOHwnd)

    return true
}

; Order list window (top-right, same look as so_order): selected order details, all orders with status,
; orange "key" button (Enter / Ctrl+N). After a clean run it moves to the next order by itself.
ShowOrdersGui() {
    global ordersGui, keySelected
    if ordersGui ; new upload replaces the old list
        enterActions.Delete(ordersGui.Hwnd), ordersGui.Destroy(), ordersGui := ""
    global T := Theme() ; pick up a Windows light/dark switch

    g := Gui("+AlwaysOnTop", "BigSeller · " . uploadName)
    ThemeGui(g)
    M := 12, W := 400
    g.MarginX := M, g.MarginY := M

    ; Header band (dark rail): platform · store, order ID, count / logistic / confirm time
    rail := " Background" . T["rail"]
    g.SetFont("s10 bold c" . T["railMeta"])
    shop := g.Add("Text", "x" . M . " y" . M . " w" . W . rail)
    g.SetFont("s13 bold c" . T["railText"])
    orderId := g.Add("Text", "x" . M . " y+4 w" . W . rail)
    g.SetFont("s10 norm c" . T["railMeta"])
    meta := g.Add("Text", "x" . M . " y+4 w" . W . " r3" . rail)
    meta.GetPos(, &my, , &mh)
    bandH := my + mh + M
    ; Added last + WS_CLIPSIBLINGS so the band never paints over the labels above
    g.Add("Text", "x0 y0 w" . (W + 2 * M) . " h" . bandH . " +0x4000000" . rail)

    g.SetFont("s11 norm c" . T["ink"])
    lv := g.Add("ListView", "x" . M . " y" . (bandH + M) . " w" . W . " r8 -Multi NoSortHdr" . LVOpts(),
        ["#", "Order ID", "แพลตฟอร์ม", "สถานะ"])
    DarkSkin(lv)
    lv.ModifyCol(1, "Integer")
    kinds := [] ; per order: "ok" | "bad" | "idle" → colour of the สถานะ cell
    ColorStatusColumn(lv, 4, kinds)
    btn := AddPrimary(g, "x" . M . " y+12 w" . W . " h44", "", OnKey)
    g.SetFont("s10 c" . T["ink2"])
    g.Add("Text", "x" . M . " y+4 w" . W . " Center", "Enter / Ctrl+N = คีย์   ·   Esc = หยุดระหว่างคีย์")
    ; Items of the selected order, last so collapsing leaves no gap
    g.SetFont("s11 c" . T["blue"])
    toggle := g.Add("Text", "x" . M . " y+10 w" . W)
    g.SetFont("c" . T["ink"])
    itemsLv := g.Add("ListView", "x" . M . " y+4 w" . W . " r8 -Multi NoSortHdr" . LVOpts(),
        ["#", "SKU", "รหัสสินค้า", "จำนวน"])
    DarkSkin(itemsLv)
    itemsLv.ModifyCol(1, "Integer"), itemsLv.ModifyCol(4, "Integer") ; numbers right-aligned
    itemKinds := [] ; per item: "bad" = whole row red (code empty / not in WINSPEED)
    ColorStatusColumn(itemsLv, 0, itemKinds)
    g.SetFont("s10 c" . T["ink2"])
    itemsHint := g.Add("Text", "x" . M . " y+4 w" . W, "ดับเบิลคลิกรายการเพื่อแก้รหัสสินค้า / จำนวน")
    g.MarginX := 0 ; AutoSize adds MarginX after the right-most control: the full-width band already ends at the edge

    ; Top-right of the work area (pixels, excludes taskbar): WINSPEED stays visible beside it
    Place() {
        g.Show("AutoSize")
        WinGetPos(, , &gw, , g.Hwnd)
        MonitorGetWorkArea(, , &top, &right)
        WinMove(right - gw, top, , , g.Hwnd)
    }

    Refresh(selectRow) {
        lv.Delete()
        kinds.Length := 0
        for i, o in orders {
            st := orderStatus[i]
            bad := st = "" ? BadItems(o) : 0
            kinds.Push(bad ? "badrow" : SubStr(st, 1, 4) = "ERR " ? "bad" : st = "" ? "idle" : "ok")
            text := bad ? "▲ รหัสไม่พบ " . bad . " รายการ" : st = "" ? "○ ยังไม่คีย์"
                : SubStr(st, 1, 4) = "ERR " ? "✗ ไม่สำเร็จ " . SubStr(st, 5) : "✓ คีย์แล้ว " . st
            lv.Add(i = selectRow ? "Select Focus Vis" : "", i, o["order_id"], o["platform"], text)
        }
        ; สถานะ fills what is left (no sideways scroll); minus a vertical scrollbar when rows overflow r8
        lv.ModifyCol(1, "AutoHdr"), lv.ModifyCol(2, "AutoHdr"), lv.ModifyCol(3, "AutoHdr")
        ControlGetPos(, , &lw, , lv)
        used := 4 + (orders.Length > 8 ? SysGet(2) : 0) ; border + SM_CXVSCROLL
        loop 3
            used += SendMessage(0x101D, A_Index - 1, 0, lv) ; LVM_GETCOLUMNWIDTH
        lv.ModifyCol(4, lw - used)
        ShowOrder(selectRow)
    }

    ShowOrder(i) {
        shown := i
        o := orders[i]
        shop.Text := o["platform"] . "  ·  " . o["store"]
        orderId.Text := o["order_id"]
        meta.Text := "Order ที่ " . i . " / " . orders.Length . "   ·   สินค้า " . o["items"].Length . " รายการ`n"
            . LogisticMsg(o) . "`nConfirm time: " . o["confirm_time"]
        bad := BadItems(o)
        btn.Text := bad ? "แก้รหัสสินค้าสีแดงก่อน  (" . bad . " รายการ)"
            : (orderStatus[i] = "" ? "คีย์" : "คีย์ซ้ำ") . " Order " . i . " / " . orders.Length
        SetPrimaryEnabled(btn, !bad)
        if bad ; show the red rows right away
            showItems := true
        FillItems()
    }

    FillItems() {
        o := orders[shown]
        bad := BadItems(o)
        itemsLv.Delete()
        itemKinds.Length := 0
        for n, it in o["items"] {
            itemKinds.Push(IsBadCode(it["goodcode"]) ? "bad" : "")
            itemsLv.Add(, n, it["sku"], it["goodcode"] = "" ? "— ไม่พบ —" : it["goodcode"], it["qty"])
        }
        ; SKU fills what is left (no sideways scroll); minus a vertical scrollbar when rows overflow r8
        itemsLv.ModifyCol(1, "AutoHdr"), itemsLv.ModifyCol(3, "AutoHdr"), itemsLv.ModifyCol(4, Round(70 * A_ScreenDPI / 96))
        ControlGetPos(, , &lw, , itemsLv)
        used := 4 + (o["items"].Length > 8 ? SysGet(2) : 0)
        for c in [1, 3, 4]
            used += SendMessage(0x101D, c - 1, 0, itemsLv) ; LVM_GETCOLUMNWIDTH
        itemsLv.ModifyCol(2, lw - used)
        toggle.Text := (showItems ? "▼" : "▶") . " รายการสินค้า (" . o["items"].Length . ")"
            . (bad ? "   ▲ รหัสไม่พบ " . bad : "")
        itemsLv.Visible := itemsHint.Visible := showItems
        Place()
    }

    OnToggle(*) {
        showItems := !showItems
        FillItems()
    }

    OnEditItem(ctrl, row) {
        if (row = 0 || typing)
            return
        if EditItem(g, orders[shown]["items"][row])
            Refresh(shown) ; status, button and red rows follow the fix
    }

    OnKey(*) {
        row := lv.GetNext()
        if (row = 0) {
            MsgBox("กรุณาเลือก Order", "Warning", "Icon! Owner" . g.Hwnd)
            return
        }
        if BadItems(orders[row]) { ; button is greyed out: red rows must be fixed first
            showItems := true
            FillItems()
            return
        }
        prev := orderStatus[row]
        if (prev != "" && MsgBox("Order " . row . " " . (SubStr(prev, 1, 4) = "ERR "
            ? "คีย์ไม่สำเร็จครั้งก่อน`nกรุณาลบรายการที่คีย์ค้างใน Winspeed ก่อน"
            : "คีย์ไปแล้ว") . "`n`nต้องการคีย์ใหม่หรือไม่?", "Warning", "YesNo Icon! Owner" . g.Hwnd) != "Yes")
            return
        if !checkWinActive() ; WINSPEED may have been closed/reopened: finds ขายเชื่อ again
            return

        g.Hide()
        order := orders[row]
        r := KeyOrder(order)
        if r.stopped {
            orderStatus[row] := "ERR " . FormatTime(, "HH:mm") . " หยุดที่ #" . r.stopped
            MsgBox("หยุดคีย์ที่รายการ #" . r.stopped . " (" . order["items"][r.stopped]["goodcode"] . ")"
                . "`nสาเหตุ: " . r.reason
                . "`n`nลบรายการที่คีย์ค้างใน Winspeed แล้วกดคีย์ Order นี้ใหม่", "Stopped", "Icon!")
            Refresh(row)
            return
        }
        orderStatus[row] := FormatTime(, "HH:mm")
        Refresh(Min(row + 1, orders.Length))
    }

    OnClose(*) {
        global ordersGui, keySelected
        if (ordersGui = g) ; an older window must not clear the newer one
            ordersGui := "", keySelected := ""
        enterActions.Delete(g.Hwnd)
        g.Destroy()
    }

    lv.OnEvent("ItemSelect", (ctrl, row, selected) => selected ? ShowOrder(row) : "")
    toggle.OnEvent("Click", OnToggle)
    itemsLv.OnEvent("DoubleClick", OnEditItem)
    g.OnEvent("Close", OnClose)
    ordersGui := g, keySelected := OnKey
    shown := 1, showItems := false
    first := 1 ; first order not typed yet
    while (first < orders.Length && orderStatus[first] != "")
        first++
    Refresh(first)
}

; Edit one item's WINSPEED code / quantity. Search by code or name; only a code that exists in WINSPEED
; can be saved (when the product list loaded). Opens left of the order window. Returns true when saved.
EditItem(owner, item) {
    global editDlgHwnd, editDown
    d := Gui("+AlwaysOnTop -MinimizeBox +Owner" . owner.Hwnd, "แก้ไขรายการ")
    ThemeGui(d)
    d.MarginX := 16, d.MarginY := 14
    W := 460, BW := (W - 10) // 2
    d.SetFont("s10 bold c" . T["ink2"])
    d.Add("Text", , "SKU")
    d.SetFont("s11 norm c" . T["ink"])
    d.Add("Text", "xm y+2 w" . W, item["sku"])
    d.SetFont("s10 bold c" . T["ink2"])
    d.Add("Text", "xm y+12", "รหัสสินค้า  (พิมพ์รหัสหรือชื่อเพื่อค้นหา · ↓ เลือก)")
    d.SetFont("s11 norm c" . T["ink"])
    codeEd := d.Add("Edit", "xm y+4 w" . W . " Uppercase Background" . T["paper"], item["goodcode"])
    DarkSkin(codeEd, "DarkMode_CFD")
    ; Results: ✓ marks the code now in the box; double-click / Enter chooses a row
    list := d.Add("ListView", "xm y+4 w" . W . " r6 -Multi NoSortHdr" . LVOpts(), ["", "รหัสสินค้า", "ชื่อสินค้า"])
    DarkSkin(list)
    d.SetFont("s10 c" . T["ink2"])
    d.Add("Text", "xm y+4 w" . W, "ดับเบิลคลิก หรือ Enter เพื่อเลือกรหัสนี้")
    d.SetFont("s10 c" . T["ink"])
    info := d.Add("Text", "xm y+4 w" . W)
    d.SetFont("s10 bold c" . T["ink2"])
    d.Add("Text", "xm y+12", "จำนวน")
    d.SetFont("s11 norm c" . T["ink"])
    qtyEd := d.Add("Edit", "xm y+4 w120 Number Background" . T["paper"], item["qty"])
    DarkSkin(qtyEd, "DarkMode_CFD")
    cancel := d.Add("Button", "xm y+16 w" . BW . " h40", "ยกเลิก (Esc)")
    DarkSkin(cancel)
    saved := false
    AddPrimary(d, "x+10 yp w" . BW . " h40", "บันทึก (Enter)", OnSave)

    ; Matches for what was typed: code prefix first, then code / name containing it (max 50)
    Filter(*) {
        q := Trim(codeEd.Value)
        list.Delete()
        if (q != "") {
            ; Closest first: exact code, code starts with, code contains, name contains;
            ; then shortest code, then A→Z. Sort key "<rank><len 3 digits><code>`t<index>".
            keys := ""
            for i, p in products {
                pos := InStr(p[1], q)
                rank := p[1] = q ? 0 : pos = 1 ? 1 : pos ? 2 : InStr(p[2], q) ? 3 : -1
                if (rank >= 0)
                    keys .= rank . Format("{:03}", StrLen(p[1])) . p[1] . "`t" . i . "`n"
            }
            keys := Sort(RTrim(keys, "`n"))
            current := StrUpper(q)
            if (keys != "")
                loop parse keys, "`n" {
                    if (A_Index > 50)
                        break
                    p := products[Integer(StrSplit(A_LoopField, "`t")[2])]
                    list.Add(, p[1] = current ? "✓" : "", p[1], p[2])
                }
        }
        ; ✓ and code fit their content; name takes the rest (minus a scrollbar when rows overflow r6)
        list.ModifyCol(1, Round(28 * A_ScreenDPI / 96)), list.ModifyCol(2, "AutoHdr")
        ControlGetPos(, , &lw, , list)
        used := 4 + (list.GetCount() > 6 ? SysGet(2) : 0)
        loop 2
            used += SendMessage(0x101D, A_Index - 1, 0, list) ; LVM_GETCOLUMNWIDTH
        list.ModifyCol(3, lw - used)
        ShowInfo()
    }

    ShowInfo() {
        code := StrUpper(Trim(codeEd.Value))
        ok := code != "" && validCodes.Has(code)
        info.SetFont("c" . (ok ? T["ok"] : code != "" && !validCodes.Count ? T["warn"] : T["bad"]))
        info.Text := code = "" ? "✗ กรุณาใส่รหัสสินค้า"
            : ok ? "✓ เลือก " . code . "  ·  " . validCodes[code]
            : !validCodes.Count ? "⚠ โหลดรายการรหัสไม่ได้ — ตรวจรหัสไม่ได้"
            : "✗ ไม่พบรหัสนี้ใน WINSPEED"
    }

    Pick(*) {
        if !(row := list.GetNext())
            return
        code := list.GetText(row, 2)
        codeEd.Value := code ; setting Value raises no Change event: results stay, only the ✓ moves
        loop list.GetCount()
            list.Modify(A_Index, , list.GetText(A_Index, 2) = code ? "✓" : "")
        ShowInfo()
        qtyEd.Focus()
        SendMessage(0xB1, 0, -1, qtyEd) ; EM_SETSEL: select all, typing replaces
    }

    OnSave(*) {
        focused := 0
        try focused := ControlGetFocus(d)
        if (focused = list.Hwnd && list.GetNext()) { ; Enter in the results = pick
            Pick()
            return
        }
        code := StrUpper(Trim(codeEd.Value))
        if IsBadCode(code) {
            ShowInfo()
            codeEd.Focus()
            return
        }
        qty := Trim(qtyEd.Value)
        if !(IsInteger(qty) && qty > 0) {
            info.SetFont("c" . T["bad"])
            info.Text := "✗ จำนวนต้องมากกว่า 0"
            qtyEd.Focus()
            return
        }
        item["goodcode"] := code, item["qty"] := Integer(qty)
        saved := true
        d.Destroy()
    }

    Down() {
        focused := 0
        try focused := ControlGetFocus(d)
        if (focused = codeEd.Hwnd) {
            list.Focus()
            if list.GetCount()
                list.Modify(1, "Select Focus")
            return
        }
        Send("{Down}")
    }

    codeEd.OnEvent("Change", Filter)
    list.OnEvent("DoubleClick", Pick)
    cancel.OnEvent("Click", (*) => d.Destroy())
    d.OnEvent("Close", (*) => d.Destroy())
    d.OnEvent("Escape", (*) => d.Destroy())
    Filter()

    hwnd := d.Hwnd
    editDlgHwnd := hwnd, editDown := Down
    owner.Opt("+Disabled") ; modal: no keying / other edits while open
    d.Show("Hide")
    WinGetPos(&ox, &oy, , , owner.Hwnd)
    WinGetPos(, , &dw, , hwnd)
    WinMove(Max(0, ox - dw), oy, , , hwnd) ; left of the order window, WINSPEED stays visible
    d.Show()
    codeEd.Focus()
    SendMessage(0xB1, 0, -1, codeEd) ; EM_SETSEL: select all
    WinWaitClose(hwnd)
    editDlgHwnd := 0, editDown := ""
    enterActions.Delete(hwnd)
    owner.Opt("-Disabled")
    WinActivate(owner.Hwnd)
    return saved
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
