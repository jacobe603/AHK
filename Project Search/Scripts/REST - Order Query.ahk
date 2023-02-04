#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance force
SetTitleMatchMode 2
SetWorkingDir %A_ScriptDir%
#Include Library\JSON.AHK

    ComINIPath=..\Data\Search.ini
    IniRead,AuthKey,%ComINIPath%,Auth,AuthKey
    IniRead,JobNumber,%ComINIPath%,CurrentProject,JobNumber
    IniRead,FolderName,%ComINIPath%,CurrentProject,FolderName


    oHTTP:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
    ;oHTTP.Option(6) := false
    If !IsObject(oHTTP)
    {
      MsgBox,4112,Fatal Error,Unable to create HTTP object
      ExitApp
    }
    ResponseFileDownload:="c:\temp\DocparserDownload.txt"
    FileDelete,%ResponseFileDownload%
    URL:="https://objm.svl.com/svl-services/rest/line-item?orderNo=" . JobNumber
    oHTTP.Open("GET",URL)
    oHTTP.SetRequestHeader("Authorization", AuthKey)
    ;OnError("LogError")
    oHTTP.Send()
    Response:=oHTTP.ResponseText
    FileAppend,%Response%,%ResponseFileDownload%
    FileRead json_String, c:\temp\DocparserDownload.txt
    parsed := JSON.Load(json_String)

    ResponseFileDownload1:="c:\temp\DocparserDownload1.txt"
    FileDelete,%ResponseFileDownload1%
    URL:="https://objm.svl.com/svl-services/rest/order/order-book/" . JobNumber
    oHTTP.Open("GET",URL)
    oHTTP.SetRequestHeader("Authorization", AuthKey)
    OnError("LogError")

    oHTTP.Send()
    Response:=oHTTP.ResponseText
    FileAppend,%Response%,%ResponseFileDownload1%

    FileRead json_String1, c:\temp\DocparserDownload1.txt
    parsed1 := JSON.Load(json_String1)
    customer := parsed1.custName
    totalVol := 0
    ;data := parsed.result[1].equipDesc
    index := 1
        ; Create the ListView:
        Gui, Add, ListView, r20 w1000 gMyListView altsubmit grid, Line No|Mfg Code|Mfg|Item|Qty|Volume|Cost|Customer PO|Rep PO|Date Added|Ship Date
        for index, value in parsed.result
        {
            totalVol += value.lnVolume
            LV_Add("",value.lineNo,value.mfgCode,value.mfgName,value.equipdesc,value.lnQty, "$" . value.lnVolume,"$" . value.lnCost, value.custPO, value.repTrackingNo, SubStr(value.dateAdded,1,10), SubStr(value.curShipDate,1,10))
        }
            LV_ModifyCol()  ; Auto-size each column to fit its contents.

        ; Display the window and return. The script will be notified whenever the user double clicks a row.
        Gui, Add, Button, gOrderBook w115, Open Order
        Gui, Add, Button, gCommSplit w115, Comm Splits
        Gui, Add, Button, Hidden Default, OK
        LV_ModifyCol(5, "Integer")
        LV_ModifyCol(1, "Integer")
        dollar := Format("{1:.2f}",totalVol)

        Gui, show,,Order Detail - %FolderName%  - %customer% - TOTAL VOLUME: $%dollar%
        Gui, +Resize

        return

MyListView:
return

ButtonOK:
            Controlget, SelectedItems, List,Selected, SysListView321, Order Detail
            X =
            D =
            Loop, Parse, SelectedItems, `n  ; Rows are delimited by linefeeds (`n).
                {
                    y =
                    RowNumber := A_Index
                    Loop, Parse, A_LoopField, %A_Tab%  ; Fields (columns) in each row are delimited by tabs (A_Tab).
                    {
                        y .= A_LoopField "|"
                        If A_Index = 6
                        {
                            text := SubStr(A_LoopField,2)
                            number := text*1
                            D += number
                        }
                    }
                    x .= y "`n"

                }
            calcTotal := "$" . format("{:.2f}",D)
            msgbox,0,Volume of Selected Lines, %calcTotal%
            clipboard := x
return



GUISize:
	width:=A_GuiWidth-15
	height:=A_GuiHeight-50
    buttonx := 20
    buttonxc := 150
    buttony := A_GuiHeight-30
	guicontrol, move, MyListView, w%width% h%height%
    guicontrol, move, Open Order, x%buttonx% y%buttony%
    guicontrol, move, Comm Splits, x%buttonxc% y%buttony%
return


OrderBook:
    Run chrome.exe "https://objm.svl.com/#/dashboard/order-inquiry-no-dollars/%JobNumber%?newWindow=true" " --new-window "   ; ADDED - Open new Chrome window and go to Order Book
return

CommSplit:
run, REST - Commission Splits.ahk
return


LogError(exception) {

    msgbox, % "Error on line " exception.Line ": " exception.Message "`n"
    ExitApp
    ;global ComINIPath
    ;run, chrome.exe "https://objm.svl.com/#/dashboard/order-entry"
    ;sleep, 2000
    ;send, {F12}
    ;sleep, 1000
    ;send, {F5}
    ;InputBox, authkeynew, Authorization Key
   ; sleep, 1000
    ;winset, AlwaysOnTop, , Authorization Key
    ;IniWrite,"%authkeynew%",%ComINIPath%,Auth,AuthKey
    ;ExitApp
}
return


GuiClose:  ; Indicate that the script should exit automatically when the window is closed.
GuiEscape:
Gui, Destroy
ExitApp
;~ ExitApp