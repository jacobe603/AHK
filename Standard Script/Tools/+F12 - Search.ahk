#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force

global match1
taskCategory := "Quote|Release|Submit|New Project|NOB"


;sEntryId := "000000004FECAB3EDF0A904C9896B8F4C0BF7EEF0700270FEF3F3663FE4CAE44E9989E91A8DE00000000010C0000270FEF3F3663FE4CAE44E9989E91A8DE0003DDCCF9D30000"
;Icon := 0

F12::
WinGetTitle, wintitle, A
WinGet, process, ProcessName, A

if process = OUTLOOK.EXE
{
	oApp := ComObjActive("Outlook.Application")	; Get Outlook
    ;oNameSpace := oApp.GetNamespace("MAPI") ; NEW
    ;oOutlookMsg := oNameSpace.GetItemFromID(sEntryId) ;NEW
    ;oOutlookMsg.MarkAsTask(Icon)
    ;oOutlookMsg.FlagStatus := 2
    ;flag := oItem.FlagStatus
    oExp := oApp.ActiveExplorer					; Get the ActiveExplorer.
    oSel := oExp.Selection						; Get the selection.
    oItem := oSel.Item(1)						; Get a selected item.
    subject := oItem.subject                   ; Get Subject Line
    id := oItem.EntryID
    clipboard := "outlook:"id
    msgbox, Entry ID for %subject% has been added to clipboard. `n%id%
}
return

+F12::
gui, destroy
;oldclipboard :=
;oldclipboard := clipboard
clipboard :=
SendInput, ^c
ClipWait
;msgbox, %Clipboard% & %OldClipboard%
if instr(clipboard,"outlook:")
	run, % Clipboard
if SubStr(Clipboard, 1, 2) = "1Z"
	Run, chrome.exe "https://wwwapps.ups.com/WebTracking/track?track=yes&trackNums="%Clipboard%
if instr(Clipboard, "Received")
    {
        emailCurrent:=ComObjActive("Outlook.Application").ActiveExplorer.Selection.Item(1)
        emailtask := parseEmail(emailCurrent)
        addTask(emailtask)
        exit
    }
if (StrLen(Clipboard)>1)
	GoToJobFolder(clipboard)

;clipboard := oldclipboard
return




GoToJobFolder(match)     ; Function that sorts through folder names on S: drive to find the match and then opens the folder.
{
    match1 := match
    global searchChoice
    searchChoice :=
    count := ""
    Loop, Files, S:\Projects\%match%*, D
        {
            ;count .= 1
            jobpath := A_LoopFileFullPath
            ;Clipboard := A_LoopFileName
            Run, %jobpath%

            ;Break
        }
            if !count {
            ;msgbox,0, File Error, Job NOT Found, 2

                Gui, Add, Text, x10 y10 w90 Center,Search Options for:
                Gui, Add, Edit, w100 vmatch1, %match1%
                Gui, Add, DropDownList, w100 vsearchChoice, Aaon DSO|Daikin SO|Ruskin FON|Email Search|FedEx|Google|Aaon - Kodis|Order Book|Job Manager|JNSearch
                Gui, Add, Button, y+m x+-23 default, OK
                ControlGetPos,x,y,buttonw,h,button
                Gui +ToolWindow
                Gui, Show , , Search Options for %match1%
                ;Gui, Destroy
                return

                GuiClose:
                ButtonOK:
                Gui, Submit
                ;GuiControlGet, searchChoice
                if (searchChoice = "Aaon DSO")
                {
                    run, Chrome.exe "https://www.aaon.com/rep-portal/order-status/details/"%match1%"?repNumber=616&pageNumber=1"
                    return
                }
                else if (searchChoice = "Daikin SO")
                {
                    ;msgbox, %match1%
                    Run, chrome.exe https://sales.daikinapplied.com/Order-Summary
                    Sleep, 700
                    clipboard := "https://sales.daikinapplied.com/Order-Detail?caller=summary&salesordernumber="match1
                    Run, chrome.exe https://sales.daikinapplied.com/Order-Summary?caller=ordernotify&salesordernumber=%match1%
                    SendInput {enter}
                }
                else if (searchChoice = "Ruskin FON")
                {
                    Run, chrome.exe "https://www.ruskin.com/VIPNet/orders?t=6&q="%match1%"&custNum=33881"
                }
                else if (searchChoice = "Email Search")
                {
                    IfWinNotExist, ahk_class rctrl_renwnd32
                    Run, OUTLOOK.EXE
                    GroupAdd, jacobemails, ahk_class rctrl_renwnd32
                    if WinActive("ahk_exe OUTLOOK.EXE")
                        GroupActivate, jacobemails, r
                    else
                        WinActivate ahk_class rctrl_renwnd32 ;you have to use WinActivatebottom if you didn't create a window group.
                    sleep, 50
                    send ^e
                    send ^a
                    Send, {Delete}
                    send, %match1%
                    send, {Enter}
                    return
                }
                else if (searchChoice = "FedEx")
                {
                    Run, chrome.exe "https://www.fedex.com/fedextrack/?trknbr="%match1%
                }
                else if (searchChoice = "Google")
                {
                    url := URIEncode(match1)
                    Run, chrome.exe "https://www.google.com/search?q="%url%
                }
                else if (searchChoice = "Aaon - Kodis")
                {
                    Run, chrome.exe "https://kodisusa.com/tracking/description-search.php?term1=1631744&dso="%match1%
                }
                else if (searchChoice = "Order Book")
                {
                    Run chrome.exe "https://objm.svl.com/#/dashboard/order-inquiry-no-dollars/%match1%?newWindow=true" " --new-window "   ; ADDED - Open new Chrome window and go to Order Book
                }
                else if (searchChoice = "Job Manager")
                {
                    Run chrome.exe "https://objm.svl.com/#/dashboard/job-detail/job-no/%match1%?newWindow=true "   ; ADDED - Open new Chrome window and go to Job Manager
                }
                else if (searchChoice = "JNSearch")
                {
                    var1 := match1
                    msgbox, %A_ScriptDir%\Scripts\JNSearch.ahk %var1%
                    run, autohotkey.exe JNSearch.ahk %var1%  ; Open Script with search as the argument
                }
                Gui Destroy
                Return

        }
        ;Run, %jobpath%
    Return
}

uriDecode(str) {
    Loop
 If RegExMatch(str, "i)(?<=%)[\da-f]{1,2}", hex)
    StringReplace, str, str, `%%hex%, % Chr("0x" . hex), All
    Else Break
 Return, str
}


UriEncode(Uri, full = 0)
{
    oSC := ComObjCreate("ScriptControl")
    oSC.Language := "JScript"
    Script := "var Encoded = encodeURIComponent(""" . Uri . """)"
    oSC.ExecuteStatement(Script)
    encoded := oSC.Eval("Encoded")
    Return encoded
}

parseemail(email)
{
    sender := email.SenderName
    subjectline := email.SubJECT
    task := subjectline " - " sender " @email"
    Return task
}

addTask(task)
{
    global taskCategory
    global taskName
    global newTask
    global DueDate
    dueDate :=
    send, {Ins}
    clipboard := task
    Gui, Font, s11, Calibri  ; Set 10-point Verdana.
    Gui, Add, ComboBox, vnewCategory w100 h35 R30 sort hwndhtaskCategory, %taskCategory%
    PostMessage, 0x0153, -1, 28,, ahk_id %htaskCategory%  ; Set height of selection field.
    PostMessage, 0x0153,  0, 28,, ahk_id %htaskCategory%  ; Set height of list items.
    Gui, Add, DateTime, vDueDate ChooseNone
    Gui, Add, Edit, vtaskName w400, %task%
    Gui, Add, Edit, vnewTask w400, %newTask%
    Gui, Add, Button, gButton1 w144, Create Task
    Gui, Add, Button, gButton2 x200 w144, Add to Todotxt
    gui, show



    ;msgbox % task
    ;insert in GUI to edit the task with + and @, due: date
    ;concanentate strings and append to todo.txt file
    return
}


Button1:
gui Submit, NoHide
newT := taskName
if newCategory
    newT .=" @" newCategory
if DueDate
{
    todoDate := substr(Duedate,1,4) "-" substr(Duedate,5,2) "-" substr(Duedate,7,2)
    newT .=" due:" todoDate
}
GuiControl,, newTask, %newT%
gui, show
return


Button2:
gui Submit
;MsgBox, % newTask
FileAppend, %newTask% `n, C:\Users\jacobe\OneDrive - Schwab Vollhaber Lubratt Inc\Documents\Todo Lists\todo.txt
return