#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force

;Go to Job Folder based on Window Title (CTRL + TILDE)
^`::

match := ""
WinGetTitle, wintitle, A
if wintitle = Inbox - JacobE@svl.com - Outlook
{
    oApp := ComObjActive("Outlook.Application")	; Get Outlook
    oExp := oApp.ActiveExplorer					; Get the ActiveExplorer.
    oSel := oExp.Selection						; Get the selection.
    oItem := oSel.Item(1)						; Get a selected item.
    wintitle := oItem.subject
}
if wintitle = Project Tracker JE.xlsm - Excel
{ 
    xl := ComObjActive("Excel.Application")
    JNCell := xl.cells(xl.ActiveCell.Row,1).Text
    GoToJobFolder(JNCell)
    return
}

RegExMatch(wintitle, "(\d{6})", match)
if match2
   msgbox, There were multiple matches
if match1 
{
    Loop, Files, S:\Projects\%match%*, D
    {
    jobpath := A_LoopFileFullPath
    Run, %jobpath%
    Clipboard := A_LoopFileName
    Break
    }
    Return
}
else 
{

    Clipboard := ""
    send, ^+{Left}
    send, ^c
    ClipWait, 1
    JobNo := Clipboard
    if RegExMatch(JobNo, "(\d{6})") = 0 {
        msgbox Not a match
        return
    }
    GoToJobFolder(JobNo)
return
}






;Go to Order Book based on Window Title or number to left of cursor (CTRL + SHIFT + TILDE)
^+`::
WinGetTitle, wintitle, A
RegExMatch(wintitle, "(\d{6})", match)
if match2
   msgbox, There were multiple matches
if match1 
{
    jobno := match
    Run chrome.exe "https://objm.svl.com/#/dashboard/order-entry/%jobno%?newWindow=true" " --new-window "   ; ADDED - Open new Chrome window and go to Order Book
}
else 
{
    send, ^+{Left}
    send, ^c
    jobno := Clipboard
    Run chrome.exe "https://objm.svl.com/#/dashboard/order-entry/%jobno%?newWindow=true" " --new-window "   ; ADDED - Open new Chrome window and go to Order Book
}
return



GoToJobFolder(match)
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
        msgbox Job NOT Found
        }
        Run, %jobpath%
    Return
}