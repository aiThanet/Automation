; Shared UI theme for the Winspeed scripts (so_order.ahk, platform_order.ahk)
; Colours from mathong-backoffice app/themes (LightTheme.ts / DarkTheme.ts). Dark follows Windows app mode.
global enterActions := Map() ; window hwnd → function run by Enter (coloured "buttons" are Text controls, not Default buttons)
global T := Theme()

; Enter = the window's orange button (replaces a Default button). Only while one of our themed windows is
; active: the scripts' own Send("{Enter}") into WINSPEED never reaches it (other window, and sent keys skip hotkeys).
#HotIf enterActions.Has(WinActive("A"))
Enter::
NumpadEnter:: enterActions[WinActive("A")]()
#HotIf

Theme() {
    dark := false
    try dark := RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme") = 0
    return dark
        ? Map("dark", true, "paper", "162130", "ink", "E7EDF5", "ink2", "9FAEC2", "rail", "0A1018",
            "railText", "E7EDF5", "railMeta", "9FAEC2", "blue", "7FA8F5", "blueWash", "1D2C48",
            "accent", "FB8A4C", "onAccent", "0D131C", "teal", "5CC6D6", "tealWash", "12343A",
            "ok", "5BD28B", "okWash", "123323", "warn", "F4B54C", "warnWash", "3A2A0E", "bad", "FF8B7E",
            "badWash", "3D1714", "line", "2A384B")
        : Map("dark", false, "paper", "FFFFFF", "ink", "0F1A2A", "ink2", "4E5B6E", "rail", "1D2B3F",
            "railText", "E7EDF5", "railMeta", "B9C6D8", "blue", "1556C9", "blueWash", "E3ECFB",
            "accent", "C2410C", "onAccent", "FFFFFF", "teal", "0F6E7F", "tealWash", "E0F1F4",
            "ok", "127A3E", "okWash", "DDF3E5", "warn", "9A5B00", "warnWash", "FFF0D6", "bad", "B42318",
            "badWash", "FDE3E0", "line", "CFD7E2")
}

; Window colours + font (Leelawadee UI = Thai UI font on Win10/11; Tahoma kept if it is missing)
ThemeGui(g) {
    g.BackColor := T["paper"]
    g.SetFont("s11 norm c" . T["ink"], "Tahoma")
    g.SetFont(, "Leelawadee UI")
    if T["dark"] ; dark title bar (DWMWA_USE_IMMERSIVE_DARK_MODE)
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", g.Hwnd, "UInt", 20, "Int*", 1, "UInt", 4)
}

; Dark scrollbars/header/edit/button skins; light mode keeps the normal Windows look
DarkSkin(ctrl, name := "DarkMode_Explorer") {
    if T["dark"]
        DllCall("uxtheme\SetWindowTheme", "Ptr", ctrl.Hwnd, "Str", name, "Ptr", 0)
}

; The one orange action of a window. Enter runs it while that window is active.
AddPrimary(g, opts, text, fn) {
    g.SetFont("bold c" . T["onAccent"])
    b := g.Add("Text", opts . " +Center +0x200 Background" . T["accent"], text) ; SS_CENTERIMAGE: centred vertically
    g.SetFont("norm c" . T["ink"])
    b.OnEvent("Click", fn)
    enterActions[g.Hwnd] := fn
    return b
}

; Grey out / restore an AddPrimary button. Its click handler must also check: a Text control has no disabled state.
SetPrimaryEnabled(b, on) {
    b.Opt("Background" . (on ? T["accent"] : T["line"]))
    b.SetFont("c" . (on ? T["onAccent"] : T["ink2"]))
    b.Redraw()
}

BGR(hex) {
    n := Integer("0x" . hex)
    return ((n & 0xFF) << 16) | (n & 0xFF00) | (n >> 16)
}

; Column header stays light: every dark header theme draws black (unreadable) header text
LVOpts() {
    return " +LV0x20 Background" . T["paper"] . " c" . T["ink"] ; LVS_EX_FULLROWSELECT
}

; Colour one ListView column per row: kinds[row] = "ok" (green) | "bad" (red) | "idle" (grey).
; col = 0: the whole row, and "bad" rows also get a red background.
; "badrow": always the whole row red + red background, whatever col is.
; kinds is read at paint time, so the caller just updates the array. NM_CUSTOMDRAW, NMLVCUSTOMDRAW 64/32-bit.
ColorStatusColumn(lv, col, kinds) {
    Draw(ctrl, lParam) {
        P := A_PtrSize
        stage := NumGet(lParam, 3 * P, "UInt")
        if (stage = 1) ; CDDS_PREPAINT → ask for each item
            return 0x20
        ; ITEMPREPAINT: colours here too (column 0 may be drawn with them, else a previous row's colour leaks in),
        ; then ask for each sub-item
        if (stage = 0x10001 || stage = 0x30001) { ; CDDS_ITEMPREPAINT [| CDDS_SUBITEM]
            row := NumGet(lParam, 5 * P + 16, "UPtr") + 1
            sub := stage = 0x10001 ? 0 : NumGet(lParam, P = 8 ? 88 : 56, "Int")
            k := row <= kinds.Length ? kinds[row] : ""
            whole := col = 0 || k = "badrow"
            if !(whole || sub = col - 1)
                k := ""
            red := k = "bad" || k = "badrow"
            NumPut("UInt", BGR(k = "ok" ? T["ok"] : red ? T["bad"] : k = "idle" ? T["ink2"] : T["ink"]),
                lParam, P = 8 ? 80 : 48) ; clrText
            NumPut("UInt", BGR(red && whole ? T["badWash"] : T["paper"]), lParam, P = 8 ? 84 : 52) ; clrTextBk
            return stage = 0x10001 ? 0x22 : 0x2 ; (CDRF_NOTIFYSUBITEMDRAW |) CDRF_NEWFONT
        }
        return 0
    }
    lv.OnNotify(-12, Draw)
}

; Indeterminate "loading" window; caller destroys it
ShowLoading(text) {
    g := Gui("+AlwaysOnTop -SysMenu +ToolWindow", "Loading")
    ThemeGui(g)
    g.MarginX := 16, g.MarginY := 14
    g.Add("Text", "w320", text)
    bar := g.Add("Progress", "w320 h8 y+10 +0x8") ; PBS_MARQUEE
    SendMessage(0x40A, 1, 30, bar) ; PBM_SETMARQUEE: start animation
    g.Show()
    return g
}
