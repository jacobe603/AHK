#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force


match := ""
match1 := ""
match2 := ""
WinGetTitle, wintitle, A
WinGet, process, ProcessName, A

if process = OUTLOOK.EXE
    {
    oApp := ComObjActive("Outlook.Application")	; Get Outlook
    oExp := oApp.ActiveExplorer					; Get the ActiveExplorer.
    oSel := oExp.Selection						; Get the selection.
    oItem := oSel.Item(1)						; Get a selected item.
    wintitle := oItem.subject                   ; Get Subject Line
    ;MsgBox, % oItem.subject
    }

if wintitle = Project Tracker JE.xlsm - Excel   ; Use first column of Project Tracker spread sheet. Update to file name of tracker.
    {
    xl := ComObjActive("Excel.Application")
    JNCell := xl.cells(xl.ActiveCell.Row,1).Text
    GoToOrderBook(JNCell)
    return
    }

RegExMatch(wintitle, "(\d{6,7})", match)          ; Looks for any 6 digit number.  If there is more than 1, it fails.

if match2
   msgbox, There were multiple matches
if match1
    {
        GoToOrderBook(match)
        Return
    }
    else
        {
            Clipboard := ""                             ; Last effort to get number.  Copies the text to the left of the cursor.
            send, ^+{Left}
            send, ^c
            ClipWait, 1
            JobNo := Clipboard
            if RegExMatch(JobNo, "(\d{6})") = 0
                {
                msgbox,0,Match Error, Not a match,2
                return
                }
            GoToOrderBook(JobNo)
            return
        }


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

GoToOrderBook(jobno)
{
    Run chrome.exe "https://objm.svl.com/#/dashboard/order-entry/%jobno%?newWindow=true" " --new-window "   ; ADDED - Open new Chrome window and go to Order Book
    return
}