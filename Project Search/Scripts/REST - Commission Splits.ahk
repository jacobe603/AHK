#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance force
SetTitleMatchMode 2
SetWorkingDir %A_ScriptDir%
#Include Library\JSON.AHK


    ;Ini File for RC Menu Variable Transfer
    ComINIPath=..\Data\Search.ini
    IniRead,AuthKey,%ComINIPath%,Auth,AuthKey
    IniRead,JobNumber,%ComINIPath%,CurrentProject,JobNumber
    IniRead,FolderName,%ComINIPath%,CurrentProject,FolderName

oHTTP:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
    If !IsObject(oHTTP)
    {
      MsgBox,4112,Fatal Error,Unable to create HTTP object
      ExitApp
    }
    ResponseFileDownload:="c:\temp\DocparserDownload2.txt"
    FileDelete,%ResponseFileDownload%
    URL:="https://objm.svl.com/svl-services/rest/split/no-security?orderNo=" . jobNumber
    oHTTP.Open("GET",URL)
    oHTTP.SetRequestHeader("Authorization", AuthKey)
    oHTTP.Send()
    Response:=oHTTP.ResponseText
    FileAppend,%Response%,%ResponseFileDownload%
    FileRead json_String, %ResponseFileDownload%
    parsed := JSON.Load(json_String)
    data := parsed.result[1].equipDesc

index := 1
    ; Create the ListView:
    Gui, Add, ListView, r10 w300 gMyListView, Comm Type|Salesman|Split

    for index, value in parsed.result
        LV_Add("",value.commType,value.slsCode,value.slsSplit)
        ;LV_ModifyCol()  ; Auto-size each column to fit its contents.

    ; Display the window and return. The script will be notified whenever the user double clicks a row.
    Gui, Show,,Splits - %FolderName%
    return

    MyListView:
    if (A_GuiEvent = "DoubleClick")
    {
        LV_GetText(RowText, A_EventInfo)  ; Get the text from the row's first field.
        ToolTip You double-clicked row number %A_EventInfo%. Text: "%RowText%"
    }
    return

GuiClose:  ; Indicate that the script should exit automatically when the window is closed.
GuiEscape:
Gui, Destroy
ExitApp
;~ ExitApp