# SO Order Protocol Button Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A button in mathong-backoffice (`/sale/external-order/mathong-system`) opens `so_order.exe` with that order already loaded — same result as Ctrl+L → paste order ID → OK — so the user only picks a page and clicks คีย์บิลนี้.

**Architecture:** `so_order.exe` registers a per-user URL protocol `mathong-so://<linebot_order_id>` (HKCU, no admin). The backoffice button is a plain link to that URL. When Windows launches a second copy of the exe with the link, that copy hands the ID to the already-running copy via `WM_COPYDATA` and exits, so an in-progress typing session is never killed. Typing into WINSPEED still starts only on the user's คีย์บิลนี้ click.

**Tech Stack:** AutoHotkey v2 (compiled with Ahk2Exe, base `AutoHotkey64.exe`), Nuxt 3 + Vuetify (mathong-backoffice).

**Spec:** No separate doc — agreed in conversation on 2026-11-23:
- Button does Ctrl+L + fill order ID + OK automatically; it does **not** click คีย์บิลนี้.
- If an order window is already open, bring it to front (do not replace it).
- If WINSPEED / ใบสั่งขาย is not open, show the same error as Ctrl+L.

## Global Constraints

- Protocol name: `mathong-so`. Link format: `mathong-so://<id>`, `<id>` = 15–20 digits (same rule as the Ctrl+L InputBox: `^\d{15,20}$`).
- Registry: `HKCU\Software\Classes\mathong-so` only. Never HKLM, never needs admin.
- Register the protocol only when `A_IsCompiled` (so running/testing the `.ahk` never points the protocol at a dev script).
- A link must never start typing into WINSPEED. It may only load the order and show the page window.
- Ctrl+L keeps working exactly as today.
- mathong-backoffice has unrelated uncommitted changes (`app/components/repair/*`, `shared/types/repair.ts`, `shared/utils/ui.ts`) — do not stage them.
- Commit only after the user approves; Automation is on `master` → create branch `feat/so-order-protocol` first.

## Review Focus

1. **Link clicked while a page is being typed** → ignored (no new window, typing not interrupted). Pinned in Task 1 test `test_open_order_ignored_while_typing`.
2. **Link for a different order while one order window is open** → existing window comes to front and a message says which order is open and to close it first; nothing is replaced. Pinned in Task 1 test `test_open_order_other_id_keeps_existing`.
3. **Malformed / hostile link** (`mathong-so://abc`, `mathong-so://123;calc`, extra path, too short) → rejected, nothing fetched. Pinned in Task 1 test `test_parse_protocol_id`.
4. **Browser appends a trailing slash** (`mathong-so://202605206421375678/`) → accepted. Pinned in `test_parse_protocol_id`.
5. **Exe moved to another folder** → next start re-registers the protocol with the new path. Pinned in Task 1 test `test_protocol_command`.

---

### Task 1: so_order.ahk — accept order ID from a `mathong-so://` link

**Files:**
- Modify: `Winspeed/so_order.ahk` (top-of-file directives, `^l::` hotkey body, `ShowPageGui` close handler)
- Create: `Winspeed/so_order_test.ahk` (assert-based self-check; `#Include`s `so_order.ahk`)

**Interfaces:**
- Produces:
  - `ParseProtocolId(arg) => String` — returns the ID if `arg` matches `i)^mathong-so:(//)?(\d{15,20})/?$`, else `""`.
  - `ProtocolCommand() => String` — `'"' A_ScriptFullPath '" "%1"'`.
  - `RegisterProtocol()` — if `A_IsCompiled` and `HKCU\Software\Classes\mathong-so\shell\open\command` default value ≠ `ProtocolCommand()`, write: key default `URL:Mathong SO`, value `URL Protocol` = `""`, command default = `ProtocolCommand()`.
  - `OpenOrder(id := "")` — the current `^l::` body moved into a function. Order of checks: `typing` → return; `pageGui` open → `pageGui.Show()`, and if `id != ""` and `id != pageOrderId` show `MsgBox("มี Order " pageOrderId " เปิดอยู่ กรุณาปิดก่อน")`, return; WINSPEED/ใบสั่งขาย checks (unchanged messages); if `id = ""` show the existing InputBox, else skip it; then `FetchOrder` → `ValidateOrder` → `ShowPageGui`.
  - New global `pageOrderId` — set with `pageGui` in `ShowPageGui`, cleared in `OnClose`.
  - `^l::OpenOrder()`.
  - Startup hand-off: replace `#SingleInstance Force` with `#SingleInstance Off`. At startup, find another running copy (`DetectHiddenWindows true`, windows of `ahk_class AutoHotkey` whose title starts with `A_ScriptFullPath " - "`, excluding `A_ScriptHwnd`). Match on the hidden main-window title, not `WinGetProcessPath` — for an uncompiled script the process is `AutoHotkey64.exe`, so every other AHK script (e.g. `platform_order`) would match. If found: send the parsed ID (may be `""`) with `WM_COPYDATA` (0x4A) and `ExitApp`. Otherwise `RegisterProtocol()`, `OnMessage(0x4A, OnCopyData)`, and if the parsed ID ≠ `""` call `OpenOrder(id)`.
  - `OnCopyData(wParam, lParam, msg, hwnd)` — reads the string from COPYDATASTRUCT, calls `SetTimer(() => OpenOrder(id), -1)` (so the sender is not blocked), returns `true`. Empty string → just `pageGui.Show()` if open.

- [ ] **Step 1: Write the failing test** `Winspeed/so_order_test.ahk`

```ahk
#Include %A_ScriptDir%\so_order.ahk
global SOWindowTitle := "NO-SUCH-WINDOW-xyz" ; never touch real WINSPEED
Check(cond, name) {
    if !cond
        throw Error("FAIL: " name)
}
; test_parse_protocol_id
Check(ParseProtocolId("mathong-so://202605206421375678") = "202605206421375678", "plain")
Check(ParseProtocolId("mathong-so://202605206421375678/") = "202605206421375678", "trailing slash")
Check(ParseProtocolId("MATHONG-SO:202605206421375678") = "202605206421375678", "no slashes, case")
for bad in ["mathong-so://abc", "mathong-so://123", "mathong-so://202605206421375678;calc",
            "mathong-so://202605206421375678/x", "http://202605206421375678", ""]
    Check(ParseProtocolId(bad) = "", "reject " bad)
; test_protocol_command
Check(ProtocolCommand() = '"' A_ScriptFullPath '" "%1"', "command uses current path")
; test_open_order_ignored_while_typing
global typing := true
OpenOrder("202605206421375678")
Check(!WinExist("Order 202605206421375678"), "no window while typing")
global typing := false
; test_open_order_other_id_keeps_existing
order := Map("items", [Map("good_id_correction", 1, "quantity", 1)], "cust_id", 1, "status", "t", "rec_updated_when", "2026-11-23 10:00:00")
ShowPageGui(order, "111111111111111", 0)
SetTimer(() => WinClose("ahk_class #32770"), -500) ; dismiss the "close first" MsgBox
OpenOrder("222222222222222")
Check(pageOrderId = "111111111111111" && WinGetList("Order 111111111111111").Length = 1, "existing kept")
Check(!WinExist("Order 222222222222222"), "no second window")
FileAppend("ALL PASS`n", "*")
ExitApp
```

- [ ] **Step 2: Run it, expect failure**

Run: `MSYS_NO_PATHCONV=1 "/c/Program Files/AutoHotkey/v2/AutoHotkey64.exe" /ErrorStdOut Winspeed/so_order_test.ahk`
Expected: error `This function does not exist` / `ParseProtocolId`.

- [ ] **Step 3: Implement the Interfaces above in `Winspeed/so_order.ahk`**

Keep every existing message, constant and behavior of Ctrl+L. COPYDATASTRUCT: `dwData` (Ptr), `cbData` (UInt), `lpData` (Ptr); send UTF-16 `(StrLen(id)+1)*2` bytes; read with `StrGet(NumGet(lParam, 2*A_PtrSize, "Ptr"))`.

- [ ] **Step 4: Run test + syntax check, expect pass**

Run: `MSYS_NO_PATHCONV=1 "/c/Program Files/AutoHotkey/v2/AutoHotkey64.exe" /ErrorStdOut /validate Winspeed/so_order.ahk` → exit 0
Run: the Step 2 command → prints `ALL PASS`

- [ ] **Step 5: Compile and verify end to end (manual, with WINSPEED ใบสั่งขาย open)**

1. Compile (user runs, see "Compile" note at bottom) → `Successfully compiled`.
2. Start `so_order.exe` once. `reg query "HKCU\Software\Classes\mathong-so\shell\open\command"` → value is `"…\Winspeed\so_order.exe" "%1"`.
3. `cmd //c start mathong-so://202605206421375678` → order window opens top-right with that order; no InputBox; nothing typed.
4. Run step 3 again → still exactly one order window, brought to front.
5. Close it, Ctrl+L → InputBox appears as before.
6. Task Manager shows one `so_order.exe` after steps 3–5.

- [ ] **Step 6: Commit (after user approval)**

```bash
git checkout -b feat/so-order-protocol
git add Winspeed/so_order.ahk Winspeed/so_order_test.ahk Winspeed/so_order.exe
git commit -m "feat(so_order): open order from mathong-so:// link"
```

---

### Task 2: Backoffice button "เปิดใน Winspeed"

**Files:**
- Modify: `mathong-backoffice/app/components/sale/external-order/MathongSystemTable.vue:44-65` (inside `<template v-slot:item.linebot_order_id>`, after the share-link `v-btn`)

**Interfaces:**
- Consumes: Task 1's link format `mathong-so://<linebot_order_id>`.

- [ ] **Step 1: Add the button**

A third `v-btn`, matching the two siblings (`icon size="x-small" class="mx-1"`), `color="success"`, `aria-label="เปิดใน Winspeed"`, `:href="`mathong-so://${item.linebot_order_id}`"`, icon `<IconMdiOpenInApp class="text-subtitle-1" />` (unplugin-icons, no import), tooltip `เปิดใน Winspeed` (`<v-tooltip activator="parent" location="top">`). No script changes.

- [ ] **Step 2: Typecheck**

Run (in mathong-backoffice): `npm run typecheck` → no new errors in `MathongSystemTable.vue`.

- [ ] **Step 3: Manual check (dev server, PC from Task 1 Step 5)**

`npm run dev` → open `/sale/external-order/mathong-system` → click the new button on a row → browser asks "Open so_order.exe?" (tick Always allow) → order window opens with that row's order ID in its title.

- [ ] **Step 4: Commit (after user approval)** — stage only the component:

```bash
git add app/components/sale/external-order/MathongSystemTable.vue
git commit -m "feat(external-order): open mathong-system order in Winspeed SO script"
```

---

## Rollout notes (not tasks)

- **Compile:** Ahk2Exe is not installed system-wide. The temporary copy is in the session scratchpad; the user must run it (auto mode blocks running a downloaded exe). Long commands wrap when pasted, so use a short helper script as before.
- **Each PC:** run the new `so_order.exe` once → protocol registered. Moving the exe → run once from the new place.
- **Deploy order:** exe to PCs first, then backoffice. A button clicked on a PC without the new exe just does nothing / browser says no app.
