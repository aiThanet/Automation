#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Include JSON.ahk


^k::
Gui, Destroy
InputBox, UserInput, ChatGPT Order ID, Please enter a order ID., , 200, 100
if ErrorLevel ;*[Untitled1]
{
	MsgBox, CANCEL was pressed.
	return
}
else
	API = http://mathongapi.jpn.local/chatgpt/order?order_id=%UserInput%&is_combine=true

oWhr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
oWhr.Open("GET", API, false)
oWhr.Send()

response := JSON.Load(oWhr.ResponseText)

MAXITEMPERPAGE = 14
TotalPage := 0
ListPage := ""

for index, element in response["items"]
    if(Mod(index, MAXITEMPERPAGE) = 1)
	{
		TotalPage++
	}

p := 1
Loop {
    if (p < TotalPage)
        ListPage .= "Page" p "|"
	else if (p = TotalPage)
		ListPage .= "Page" p
    else
		break
	p++
}

if (TotalPage = 0)
{
	MsgBox, No Data
	return
}

MsgBox, % "Order: " response["order_id"] "`nอัพเดต: " SubStr(response["rec_updated_when"],1,19) " (+7hr)`nจำนวนหน้าทั้งหมด: " TotalPage " หน้า"

clipboard := response["cust_code"]

Gui, Margin, +20, +20

Gui, Add, DropDownList, xm ym w200 h60 r10 vDDL gSubmit_All, % ListPage
Gui, Add, Button, xm w200 h40 gDo_AddItem, Confirm

Gui, Show, , Select Page
return

Submit_All:
	Gui, Submit, NoHide
	return

Do_AddItem:
	if(DDL = "")
	{
		MsgBox, % "Please select page"
		return
	}

	ppage := substr(DDL,5)
	minItem := ((ppage-1) * (MAXITEMPERPAGE)) + 1
	maxItem := ((ppage) * (MAXITEMPERPAGE))
	Gui, Destroy

	for index, element in response["items"]
	{
		if(index >= minItem and index <= maxItem)
		{
			Send, % element["good_code_correction"]
			Loop, 5{
				Send, {Enter}
			}
			Send, % element["quantity"]
			Loop, 4{
				Send, {Enter}
			}
		}
	}