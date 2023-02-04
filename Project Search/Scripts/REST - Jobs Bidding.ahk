#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; Working in scrpit
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance force
SetTitleMatchMode 2 ;A window's title can contain WinTitle anywhere inside it to be a match.
SetWorkingDir %A_ScriptDir%
#Include Library\Class_LV_Colors.ahk
#Include Library\JSON.AHK
ComINIPath = ..\Data\Search.ini

menu, Tray, Icon, icons\svl.png
IniRead,AuthKey,%ComINIPath%,Auth,AuthKey
IniRead,SalesCode,%ComINIPath%,Sales,SalesCode


    oHTTP:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
    ;oHTTP.Option(6) := false
    If !IsObject(oHTTP)
    {
        MsgBox,4112,Fatal Error,Unable to create HTTP object
        ExitApp
    }
    ResponseFileDownload:="c:\temp\DocparserDownload.txt"
    FileDelete,%ResponseFileDownload%
    URL:="https://objm.svl.com/svl-services/rest/job/bid"
    JSONString = {"todaysJobs":false,"myBids":false}
    oHTTP.Open("POST",URL)
    oHTTP.SetRequestHeader("Authorization", AuthKey)
    oHTTP.SetRequestHeader("Content-Type", "application/json")
    OnError("LogError")
    oHTTP.Send(JSONString)
    Response:=oHTTP.ResponseText
    FileAppend,%Response%,%ResponseFileDownload%

    FileRead json_String, c:\temp\DocparserDownload.txt
    parsed := JSON.Load(json_String)


index := 1
    ; Create the ListView:
    Gui, Font, s10, Roboto
    Gui, Add, ListView, r20 w1200 gMyListView vMyListView altsubmit grid hwndHLV, Job Num|Bid Date|Bid Time|Job Name|City|State|Engineer|Sales Team|Location ; added hwndHLV, maybe look at changing the HLV to something else
    for index, value in parsed.result
        if InStr(value.steam,SalesCode)
        {
        LV_Add("",value.jobNo,SubStr(value.bidDate,1,10),value.dueTime,value.name,value.city, value.state, value.engineer, value.steam, value.location)
        }

    LV_ModifyCol()  ; Auto-size each column to fit its contents.
    LV_ModifyCol(2, "Sort")

    ; Create a new instance of LV_Colors
    ; Set the colors for selected rows
    CLV := New LV_Colors(HLV)
    CLV.SelectionColors(0xffebe4,0xff5722)
    CLV.AlternateRows(0xfafafa)
    If !IsObject(CLV) {
    MsgBox, 0, ERROR, Couldn't create a new LV_Colors object!
    ExitApp
    }

; Highlight jobs bidding today
Loop % LV_GetCount()
{
    LV_GetText(RetrievedText, A_Index,2)
    RetrievedText := StrReplace(RetrievedText,"-")
    todayDate := subStr(A_now,1,8)
    if InStr(RetrievedText, todayDate)
        CLV.Row(A_Index, 0x00ff00)  ; Select each row whose first field contains the filter-text.
}

; Highlight jobs that are on the bid board
Loop % LV_GetCount()
{
    LV_GetText(RetrievedText, A_Index,9)
    if InStr(RetrievedText, "Bid Board")
        CLV.Row(A_Index, 0xa0a0ff)  ; Select each row whose first field contains the filter-text.
}


    LV_ModifyCol(2, "Sort")
;------------------------------------
    ;CLV.Clear(1, 1)
    ;CLV.AlternateRows(0xfafafa, 0x000000)
    Gui, Add, Button, gRefresh, Refresh
    Gui, show,H300,JOBS BIDDING FOR %SalesCode%
    Gui, +Resize

MyListView:
if (A_GuiEvent = "DoubleClick")
{
    LV_GetText(jobNum, A_EventInfo, 1)  ; Get the text from the row's first field.
    GoToJobFolder(jobNum)
}
   If A_GuiEvent = RightClick  ; IF USER RIGHT CLICKS A SEARCH RESULT, Copy Contents for TodoTXT
    {   ;msgbox, Right click
        GuiControlGet, FocusedControl, FocusV
        FocusedRow := LV_GetNext("", "Focused")
    ;   If no file is selected, do nothing.
        if FocusedRow = 0
            return
        LV_GetText(jobNo, FocusedRow, 1)
        LV_GetText(jobName, FocusedRow, 4)
        LV_GetText(bidDate, FocusedRow, 2)
            ;IniWrite,%jobno%,%ComINIPath%,CurrentProject,JobNumber
            ;IniWrite,%jobName%,%ComINIPath%,CurrentProject,JobName
            ;Run, RC.ahk
            clipboard = +%jobNo% %jobName% due:%bidDate% @quote
    }
return

GUISize:
	width:=A_GuiWidth-15
	height:=A_GuiHeight-50
    buttonx := 20
    buttony := A_GuiHeight-30
	guicontrol, move, MyListView, w%width% h%height%
    guicontrol, move, Refresh, x%buttonx% y%buttony%
return

Refresh:
Reload
Return


GoToJobFolder(match)     ; Function that sorts through folder names on S: drive to find the match and then opens the folder.
{
    count := ""
    Loop, Files, S:\Projects\%match%*, D
        {
            count .= 1
            jobpath := A_LoopFileFullPath
            Clipboard := A_LoopFileName
            Break
        }
            if !count {
            msgbox,0, File Error, Job NOT Found, 2
        }
        Run, %jobpath%
    Return
}


LogError(exception) {
    ;Global ComINIPath
    ;run, chrome.exe "https://objm.svl.com/#/dashboard/order-entry"
    ;sleep, 2000
   ; send, {F12}
    ;sleep, 1000
    ;send, {F5}
   ; InputBox, authkeynew, Authorization Key
   ; sleep, 1000
    ;winset, AlwaysOnTop, , Authorization Key
    ;IniWrite,"%authkeynew%",%ComINIPath%,Auth,AuthKey
    run, "REST - GET AUTH.ahk"
    ExitApp
}
return

GUIClose:
GUIEscape:
ExitApp

