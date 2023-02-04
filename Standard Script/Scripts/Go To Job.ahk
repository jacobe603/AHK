#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force

match := ""
xl := ""
WinGetTitle, wintitle, A
WinGet, process, ProcessName, A

if process = OUTLOOK.EXE
    {
    oApp := ComObjActive("Outlook.Application")	; Get Outlook
    oExp := oApp.ActiveExplorer					; Get the ActiveExplorer.
    oSel := oExp.Selection						; Get the selection.
    oItem := oSel.Item(1)						; Get a selected item.
    wintitle := oItem.subject                   ; Get Subject Line
    ;MsgBox, % wintitle
    }

if wintitle = Project Tracker JE.xlsm - Excel   ; Use first column of Project Tracker spread sheet. Update to file name of tracker.
    {
    xl := ComObjActive("Excel.Application")
    JNCell := xl.cells(xl.ActiveCell.Row,1).Text
    GoToJobFolder(JNCell)
    return
    }

RegExMatch(wintitle, "(\d{6})", match)          ; Looks for any 6 digit number.  If there is more than 1, it fails.
if RegExMatch(wintitle, "(\d{6})") = 0
   RegExMatch(wintitle, "(\d{7})", match)
else if match1
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
            Clipboard := ""                             ; Last effort to get number.  Copies the text to the left of the cursor.
            send, ^+{Left}
            send, ^c
            ClipWait, 1
            JobNo := Clipboard
		GoToJobFolder(JobNo)
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

