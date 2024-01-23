#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Include JSON.ahk

^i::
InputBox, UserInput, Order ID, Please enter a order ID., , 200, 100
if ErrorLevel ;*[Untitled1]
	MsgBox, CANCEL was pressed.
else
	API = http://mathongapi.jpn.local/firebase/liff/order?id=%UserInput%

oWhr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
oWhr.Open("GET", API, false)
oWhr.Send()

response := JSON.Load(oWhr.ResponseText)


MsgBox, % "Order: " response["id"] "`nอัพเดต: " response["updatedAt"]
for index, element in response["items"]
{
			Send, % element["GoodCode"]
			Loop, 5{
				Send, {Enter}
			}
			Send, % element["itemQty"]
			Loop, 4{
				Send, {Enter}
			}
		}
			
		