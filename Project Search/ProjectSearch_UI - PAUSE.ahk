;Project Search UI - This needs to be running to intercept hotkey and launch Project Search
;Created by Jacob Erbele - updated 8/17/2022
#NoEnv
#SingleInstance Force
SendMode Input

SetTitleMatchMode,2
CoordMode, Mouse, Screen

Menu, Tray, Icon, shell32.dll, 210

ComINIPath=%A_ScriptDir%\Data\Search.ini
LibPath=%A_ScriptDir%\Scripts\Library\
IniWrite,%ComINIPath%,%ComINIPath%,Paths,DataPath
IniWrite,%LibPath%,%ComINIPath%,Paths,LibPath
IniRead,SalesName,%ComINIPath%,Sales,SalesName
IniRead,SalesCode,%ComINIPath%,Sales,SalesCode
IniRead,AllSales,%ComINIPath%,Sales,AllSales

SearchAppPath = %A_ScriptDir%\Scripts\ProjectSearch_v4.ahk

If (SalesName = "" or SalesCode = "")
    {
    Gui, Add, Text,, Select Your Name & Initials
    Gui, Add, DropDownList, vList1, %AllSales%
    Gui, Add, Edit, vSalesCode, Intials
    Gui, Add, Button,,Submit
    Gui, Show
    }

MsgBox, Project Search v4 is now running in the background. Press PAUSE to activate.
return

ButtonSubmit:
    Gui, submit
    Gui, destroy
    IniWrite,%List1%,%ComINIPath%,Sales,SalesName
    IniWrite,%SalesCode%,%ComINIPath%,Sales,SalesCode
    StringSplit, Name, List1,`,
    msgbox Thanks,%Name2%!  Search.ini has been updated.
Return

F23::
PAUSE::
        MouseGetPos, xpos
        ;MsgBox, % xpos
        Run, %SearchAppPath%
        WinWait,Project Search
        if xpos > 3840
        WinMove,,,3916,894

Return




;~If script needs to be run at an elevated level- this takes care of that
;if (! A_IsAdmin){ ;http://ahkscript.org/docs/Variables.htm#IsAdmin
;	Run *RunAs "%A_ScriptFullPath%"  ; Requires v1.0.92.01+
;	ExitApp
;}
