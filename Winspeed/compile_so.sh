#!/bin/sh
# Compile so_order.ahk -> so_order.exe (exit the running so_order.exe from the tray first)
# Ahk2Exe: Automation/tools/Ahk2Exe.exe (not in git; official release github.com/AutoHotkey/Ahk2Exe)
DIR="$(cd "$(dirname "$0")" && pwd)"
MSYS_NO_PATHCONV=1 "$DIR/../tools/Ahk2Exe.exe" \
  /in "$(cygpath -w "$DIR/so_order.ahk")" \
  /out "$(cygpath -w "$DIR/so_order.exe")" \
  /base "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
echo "exit=$?"
ls -la "$DIR/so_order.exe"
