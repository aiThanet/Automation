#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Include JSON.ahk

^i::
InputBox, UserInput, Order Number, Please enter a order number., , 200, 100
if ErrorLevel ;*[Untitled1]
	MsgBox, CANCEL was pressed.
else
	API = https://us-central1-mathong-b6742.cloudfunctions.net/widgets/order/%UserInput%

oWhr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
oWhr.Open("GET", API, false)
oWhr.Send()

response := JSON.Load(oWhr.ResponseText)
MsgBox, % "ร้าน: " response["customer_name"] "`nเวลา: " response["readable_time"]
for index, element in response["orders"]
{
			Send, % element["id"]
			Loop, 5{
				Send, {Enter}
			}
			Send, % element["qty"]
			Loop, 4{
				Send, {Enter}
			}
		}
			
		