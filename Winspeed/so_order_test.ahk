; Self-check for so_order.ahk protocol handling. Run: AutoHotkey64.exe /ErrorStdOut so_order_test.ahk
#Warn VarUnset, StdOut ; warnings to stdout, not a blocking dialog
#Include %A_ScriptDir%\so_order.ahk
global SOWindowTitle := "NO-SUCH-WINDOW-xyz" ; never touch real WINSPEED
OnError((e, *) => (FileAppend("ERROR: " . e.Message . " " . e.Extra . " line " . e.Line . "`n", "*"), ExitApp(2)))

MY_DIALOGS := "ahk_class #32770 ahk_pid " . DllCall("GetCurrentProcessId") ; only this test's dialogs, never WINSPEED's

CountAndCloseDialogs() {
    global dialogsSeen := WinGetList(MY_DIALOGS).Length
    while WinExist(MY_DIALOGS)
        WinClose(MY_DIALOGS), Sleep(100)
}

MakeCds(str) {
    static keep := []
    s := Buffer((StrLen(str) + 1) * 2), StrPut(str, s)
    cds := Buffer(3 * A_PtrSize, 0)
    NumPut("UPtr", 0, "UInt", s.Size, cds)
    NumPut("Ptr", s.Ptr, cds, 2 * A_PtrSize)
    keep.Push(s, cds)
    return cds.Ptr
}

Check(cond, name) {
    if !cond {
        FileAppend("FAIL: " . name . "`n", "*")
        ExitApp(1)
    }
}

; test_parse_protocol_id
Check(ParseProtocolId("mathong-so://202605206421375678") = "202605206421375678", "plain")
Check(ParseProtocolId("mathong-so://202605206421375678/") = "202605206421375678", "trailing slash")
Check(ParseProtocolId("MATHONG-SO:202605206421375678") = "202605206421375678", "no slashes, case")
for badLink in ["mathong-so://abc", "mathong-so://123", "mathong-so://202605206421375678;calc",
    "mathong-so://202605206421375678/x", "http://202605206421375678", ""]
    Check(ParseProtocolId(badLink) = "", "reject " . badLink)

; test_parse_liff_link: LINE OA orders use mathong-so://liff/<id> (IDs overlap with mathong-system, so source is explicit)
Check(ParseProtocolId("mathong-so://liff/261123722687489") = "liff/261123722687489", "liff link")
Check(ParseProtocolId("mathong-so://liff/2611214812182550/") = "liff/2611214812182550", "liff link trailing slash")
for badLink in ["mathong-so://liff/", "mathong-so://liff/abc", "mathong-so://lif/261123722687489",
    "mathong-so://liff/261123722687489/x", "mathong-so://liff//261123722687489"]
    Check(ParseProtocolId(badLink) = "", "reject " . badLink)

; test_is_order_key
Check(IsOrderKey("202605206421375678") && IsOrderKey("liff/261123722687489"), "valid keys")
Check(!IsOrderKey("") && !IsOrderKey("liff/") && !IsOrderKey("x/261123722687489") && !IsOrderKey("123"), "invalid keys")

; test_normalize_liff_order: LIFF fields → same shape as a mathong-system order
liffRaw := Map("id", "261123722687489", "isConfirm", true, "updatedAt", "2026-11-23 04:13:47.504000+00:00",
    "line_display_name", "สมชาย LINE",
    "items", [Map("GoodID", "5341", "GoodCode", "HG22-04", "GoodMarketName", "เครื่องทดสอบ", "itemQty", 20)])
n := NormalizeLiffOrder(liffRaw)
it := n["items"][1]
Check(it["good_id_correction"] = "5341" && it["good_code_correction"] = "HG22-04" && it["quantity"] = 20, "liff item ids/qty")
Check(it["winspeed_good_code"] = "HG22-04" && it["winspeed_good_name"] = "เครื่องทดสอบ", "liff item display")
; LIFF without linked customer: user picks it in WINSPEED (cust_id "" → not typed); LINE name kept separately
Check(n["cust_id"] = "" && n["line_name"] = "สมชาย LINE", "liff unlinked: no cust_id, LINE name")
; LIFF with linked customer (API resolved shopName "CODE : NAME" → WINSPEED CustID): typed like mathong-system
liffRaw["cust_id"] := 9041, liffRaw["cust_code"] := "N0200-AR", liffRaw["cust_name"] := "บริษัท ทดสอบ จำกัด"
n2 := NormalizeLiffOrder(liffRaw)
Check(n2["cust_id"] = 9041 && n2["cust_code"] = "N0200-AR" && n2["cust_name"] = "บริษัท ทดสอบ จำกัด", "liff linked customer")
Check(n2["line_name"] = "สมชาย LINE", "liff linked keeps LINE name")
Check(n["rec_updated_when"] = "2026-11-23 04:13:47.504000+00:00" && n["link_so"].Length = 0, "liff updated/link_so")

; test_is_copy_title: compiled exe main window title is the bare path; .ahk is "path - AutoHotkey v2..."
Check(IsCopyTitle(A_ScriptFullPath), "compiled title (bare path)")
Check(IsCopyTitle(A_ScriptFullPath . " - AutoHotkey v2.0.19"), "script title")
Check(!IsCopyTitle(A_ScriptFullPath . "x"), "other file with same prefix")
Check(!IsCopyTitle("C:\other\so_order.exe"), "other path")

; test_protocol_command
Check(ProtocolCommand() = '"' . A_ScriptFullPath . '" "%1"', "command uses current path")

; test_open_order_ignored_while_typing
global typing := true
OpenOrder("202605206421375678")
Check(!WinExist("Order 202605206421375678"), "no window while typing")
global typing := false

; test_open_order_other_id_keeps_existing
testOrder := Map("items", [Map("good_id_correction", 1, "quantity", 1)], "cust_id", 1, "status", "t",
    "rec_updated_when", "2026-11-23 10:00:00")
ShowPageGui(testOrder, "111111111111111", 0)
SetTimer(() => WinClose(MY_DIALOGS), -500) ; dismiss the "close first" MsgBox
OpenOrder("222222222222222")
Check(pageOrderId = "111111111111111" && WinGetList("Order 111111111111111").Length = 1, "existing kept")
Check(!WinExist("Order 222222222222222"), "no second window")
WinClose("Order 111111111111111")
Sleep(200)

; test_onclose_clears_only_own_window
ShowPageGui(testOrder, "444444444444444", 0)
ShowPageGui(testOrder, "555555555555555", 0)
WinClose("Order 444444444444444")
Sleep(200)
Check(pageOrderId = "555555555555555" && pageGui != "", "closing older window keeps newer one tracked")
WinClose("Order 555555555555555")
Sleep(200)

; test_open_order_reentry_ignored: 2nd open while the 1st is still busy (here: blocked in its MsgBox)
global WinspeedClass := "ahk_class NO-SUCH-CLASS" ; OpenOrder stops at "Winspeed not open" MsgBox
global dialogsSeen := 0
SetTimer(() => OpenOrder("333333333333333"), -300)
SetTimer(CountAndCloseDialogs, -900)
OpenOrder("333333333333333")
Check(dialogsSeen = 1, "re-entrant open ignored (dialogs seen: " . dialogsSeen . ")")

; test_copydata_empty_or_bad_id_does_nothing (no InputBox, no fetch, no MsgBox)
global dialogsSeen := 0
OnCopyData(0, MakeCds(""), 0x4A, 0)
OnCopyData(0, MakeCds("123;calc"), 0x4A, 0)
SetTimer(CountAndCloseDialogs, -600)
Sleep(900)
Check(dialogsSeen = 0, "empty/bad id via WM_COPYDATA opens nothing (dialogs seen: " . dialogsSeen . ")")

; test_copydata_valid_id_opens: valid ID reaches OpenOrder (stops at fake "Winspeed not open" MsgBox)
global dialogsSeen := 0
OnCopyData(0, MakeCds("202605206421375678"), 0x4A, 0)
SetTimer(CountAndCloseDialogs, -600)
Sleep(900)
Check(dialogsSeen = 1, "valid id via WM_COPYDATA opens order (dialogs seen: " . dialogsSeen . ")")

FileAppend("ALL PASS`n", "*")
ExitApp(0)
