

SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
Menu, Tray, Icon, shell32.dll, 283 ;if commented in, this line will turn the H icon into a little grey keyboard-looking thing.
;SetKeyDelay, 0 ;IDK exactly what this does.

;;WHAT'S THIS ALL ABOUT??

;;THE SHORT VERSION:
;; https://www.youtube.com/watch?v=kTXK8kZaZ8c

;;THE LONG VERSION:
;; https://www.youtube.com/playlist?list=PLH1gH0v9E3ruYrNyRbHhDe6XDfw4sZdZr


;;Location for where to put a shortcut to the script, such that it will start when Windows starts:
;;  Here for just yourself:
;;  C:\Users\YOUR_USERNAME\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup
;;  Or here for all users:
;;  C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp

#NoEnv
SendMode Input
#InstallKeybdHook
#UseHook On
#SingleInstance force
#MaxHotkeysPerInterval 2000

;;The lines below are optional. Delete them if you need to.
#HotkeyModifierTimeout 60 ; https://autohotkey.com/docs/commands/_HotkeyModifierTimeout.htm
#KeyHistory 200 ; https://autohotkey.com/docs/commands/_KeyHistory.htm ; useful for debugging.
#MenuMaskKey vk07 ;https://autohotkey.com/boards/viewtopic.php?f=76&t=57683
#WinActivateForce ;https://autohotkey.com/docs/commands/_WinActivateForce.htm ;prevent taskbar flashing.
;;The lines above are optional. Delete them if you need to.


+!^1::
IfWinNotExist, ahk_class CabinetWClass
	Run, explorer.exe
GroupAdd, groupExplorers, ahk_class CabinetWClass
if WinActive("ahk_exe explorer.exe")
	GroupActivate, groupExplorers, r
else
	WinActivate ahk_class CabinetWClass ;you have to use WinActivatebottom if you didn't create a window group.
Return

+!^2::
IfWinNotExist, ahk_class rctrl_renwnd32
	Run, OUTLOOK.EXE
GroupAdd, groupEmails, ahk_class rctrl_renwnd32
if WinActive("ahk_exe OUTLOOK.EXE")
	GroupActivate, groupEmails, r
else
	WinActivate ahk_class rctrl_renwnd32 ;you have to use WinActivatebottom if you didn't create a window group.
Return

+!^3::
switchToExcel()
Return

;FN Key + Key
F13::
Run, chrome.exe "https://sales.daikinapplied.com/Order-Summary"
return

F14::
Run, chrome.exe "https://www.aaon.com/rep-portal/order-status"
return

F15::
Run, chrome.exe "https://www.titus-hvac.com/mytitus"
return

F16::
Run, chrome.exe "https://repnet.lorencook.com/Orders"
return

F17:: ;****NEEDS UPDATE
Run, C:\Users\%A_username%\OneDrive - Schwab Vollhaber Lubratt Inc\Desktop\AHK - Local\Scripts\Outlook Task Create.ahk
return

F18::
Run, chrome.exe "https://svl1.sharepoint.com/sites/SVL_Sales_Support"
return

F19:: ;****NEEDS UPDATE
Run, C:\Users\%A_username%\OneDrive\SVL\Documents\Commission Splitter.xlsm
return

F20::
RunExist("C:\Program Files (x86)\Hughesoft\todotxt.net\todotxt.exe")
return

;F21::
InputBox, EmailSearch, Outlook Search
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
send, %EmailSearch%
send, {Enter}
return

;F22::
return


;;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++;;
;;+++++++++++++++++ BEGIN SECOND KEYBOARD F24 ASSIGNMENTS +++++++++++++++++++++;;
;;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++;;

;; You should DEFINITELY not be trying to add a 2nd keyboard unless you're already
;; familiar with how AutoHotkey works. I recommend that you at least take this tutorial:
;; https://autohotkey.com/docs/Tutorial.htm

;; The point of these is that THE TOOLTIPS ARE MERELY PLACEHOLDERS. When you add a function of your own, you should delete or comment out the tooltip.

;; You should probably use something better than Notepad for your scripting. (Do NOT use Word.)
;; I use Notepad++. "Real" programmers recoil from it, but it's fine for my purposes.
;; https://notepad-plus-plus.org/
;; You'll probably want the syntax highlighting:  https://stackoverflow.com/questions/45466733/autohotkey-syntax-highlighting-in-notepad

;;COOL BONUS BECAUSE YOU'RE USING QMK:
;;The up and down keystrokes are registered seperately.
;;Therefore, your macro can do half of its action on the down stroke,
;;And the other half on the up stroke. (using "keywait,")
;;This can be very handy in specific situations.
;;The Corsair K55 keyboard fires the up and down keystrokes instantly.

#if (getKeyState("F24", "P")) ;<--Everything after this line will only happen on the secondary keyboard that uses F24.
F24::return ;this line is mandatory for proper functionality
escape::tooltip, "[F24] You might wish to not give a command to escape. Could cause problems. IDK."

;;------------------------TOP ROW--------------------------;;
`::
WinGetTitle, Title, A
PostMessage, 0x112, 0xF060,,, %Title%
return

1::
IfWinNotExist, ahk_class CabinetWClass
	Run, explorer.exe
GroupAdd, groupExplorers, ahk_class CabinetWClass
if WinActive("ahk_exe explorer.exe")
	GroupActivate, groupExplorers, r
else
	WinActivate ahk_class CabinetWClass ;you have to use WinActivatebottom if you didn't create a window group.
Return

2::
IfWinNotExist, ahk_class rctrl_renwnd32
	Run, OUTLOOK.EXE
GroupAdd, groupEmails, ahk_class rctrl_renwnd32
if WinActive("ahk_exe OUTLOOK.EXE")
	GroupActivate, groupEmails, r
else
	WinActivate ahk_class rctrl_renwnd32 ;you have to use WinActivatebottom if you didn't create a window group.
Return

3::
switchToExcel()
Return
4::
switchToWord()
Return
5::
switchToChrome()
Return
6::
switchToBlueBeam()
Return

;;------------------------2ND ROW--------------------------;;
q:: ;Doesn't work for some reason
RunExist("C:\Users\" A_username "\AppData\Local\Daikin Applied\Daikin Tools\Application\Daikin.DaikinTools.exe")
return

w:: ;~ Create New Email
	try
		outlookApp := ComObjActive("Outlook.Application")
	catch
		outlookApp := ComObjCreate("Outlook.Application")
	MailItem := outlookApp.CreateItem(0)
	MailItem.Display
;run, "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" /c ipm.note
return

e::
r::
t::

y::
run, C:\Users\%A_username%\OneDrive - Schwab Vollhaber Lubratt Inc\Documents\SVL\Project Tracker JE.xlsm
return

\::tooltip, [F24]  %A_thisHotKey%
;;capslock::tooltip, [F24] capslock - this should have been remapped to F20. Keep this line commented out.


;;------------------------3RD ROW--------------------------;;

a::
RunExist("C:\Program Files (x86)\AAONECat32\ECat5\AAON.Ecat.Startup.exe")
return

s::
SendEvent !{n}{a}{s}{Down}{Enter}
return

d::
run, ..\Scripts\Commission Split Parser.ahk
return

f::
run, ..\Scripts\Order Book Line Enter.ahk
return

g::
SetTitleMatchMode, 2
GroupAdd, Post, - Message (HTML)
WinClose, ahk_group Post
return

h::
GroupAdd,ExplorerGroup, ahk_class CabinetWClass
GroupAdd,ExplorerGroup, ahk_class ExploreWClass

; close all Explorer windows
  if ( WinExist("ahk_group ExplorerGroup") )
    WinClose,ahk_group ExplorerGroup
return

enter::tooltip, [F24]  %A_thisHotKey%


;;------------------------Bottom ROW--------------------------;;
!z::
Run, ..\Scripts\Go to Order.ahk
return

z::
Run, ..\Scripts\Go To Job.ahk
return

x::
RunExist("C:\Program Files (x86)\EDGE Software\EDGE\Bin\Edge.Borg.Pricing.exe")
return

c::
RunExist("C:\Program Files (x86)\Cookware\cookware.exe")
return

v::
Run, Project Paster_v2.ahk
return


;Not used right now
b::
 IfWinNotExist FILE SEARCH ahk_class AutoHotkeyGUI
           ; IfWinNotExist SearchFiles ahk_class AutoHotkeyGUI
            {
                ;~ SearchAppPath := "%A_ScriptDir%"\Project Search\SearchFiles_v3.7 RC.ahk"
                Run, "%A_ScriptDir%\Project Search\SearchFiles_v3.7 RC.ahk"
            }
          ;  else
           ;     WinActivate SearchFiles ahk_class AutoHotkeyGUI
        Else
            Winactivate FILE SEARCH ahk_class AutoHotkeyGUI
return

/::tooltip, [F24]  %A_thishotKey%


;;------------------------FN 3rd Row--------------------------;;

j::
k::
l::tooltip, [F24]  %A_thisHotKey%
return

;;--------------------------------------------------;;
space::
tooltip, [F24] SPACEBAR. This will now clear remaining tooltips.
sleep 500
tooltip,
return
;;And THAT^^ is how you do multi-line instructions in AutoHotkey.
;;Notice that the very first line, "space::" cannot have anything else on it.
;;Again, these are fundamentals that you should have learned from the tutorial.


;;============== END OF ALL KEYBOARD KEYS ===============;;

#if ;this line will end the F24 secondary keyboard assignments.


;;;---------------IMPORTANT: HOW TO USE #IF THINGIES-------------------

;;You can use more than one #if thingy at a time, but it must be done like so:
#if (getKeyState("F24", "P")) and if WinActive("ahk_exe Excel.exe")
F2::msgbox, You pressed F2 on your secondary keyboard while inside of Excel

;; HOWEVER, You still would still need to block F1 using #if (getKeyState("F24", "P"))
;; If you don't, it'll pass through normally, any time Premiere is NOT active.
;; Does that make sense? I sure hope so. Experiment with it if you don't understand.

;; Alternatively, you can use the following: (Comment it in, and comment out other instances of F1::)
; #if (getKeyState("F24", "P"))
; F1::
; if WinActive("ahk_exe Adobe Premiere Pro.exe")
; {
	; msgbox, You pressed F1 on your secondary keyboard while inside of Premiere Pro
	; msgbox, And you did it by using if WinActive()
; }
; else
	; msgbox, You pressed F1 on your secondary keyboard while NOT inside of Premiere Pro
;;This is easier to understand, but it's not as clean of a solution.

;; Here is a discussion about all this:
;; https://github.com/TaranVH/2nd-keyboard/issues/65

;;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;+++++++++++++++++ END OF SECOND KEYBOARD F24 ASSIGNMENTS ++++++++++++++++++++++
;;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

;;Note that this whole script was written for North American keyboard layouts.
;;IDK what you foreign language peoples are gonna have to do!
;;At the very least, you'll have some duplicate keys.

#if


;;~~~~~~~~~~~~~~~~~DEFINE YOUR FUNCTIONS BELOW THIS LINE~~~~~~~~~~~~~~~~~~~~~~~~~

RunExist(SoftLocation) ; Open if it exists otherwise run the soft location to open program
{
  IfWinExist, ahk_exe %SoftLocation%
  WinActivate
else
  Run, %SoftLocation%
  if (ErrorLevel = 1)
    MsgBox File does not exist.  Double check ini file.
  return
}

switchToChrome()
{
IfWinNotExist, ahk_exe chrome.exe
	Run, chrome.exe

if WinActive("ahk_exe chrome.exe")
	Sendinput ^{tab}
else
	WinActivate ahk_exe chrome.exe
}

switchToBlueBeam()
{
IfWinNotExist, ahk_exe Revu.exe
	Run, C:\Program Files\Bluebeam Software\Bluebeam Revu\20\Revu\Revu.exe

if WinActive("ahk_exe Revu.exe")
	Sendinput ^{tab}
else
	WinActivate ahk_exe Revu.exe
}

switchToCookware()
{
IfWinNotExist, ahk_exe cookware.exe
	Run, C:\Program Files (x86)\Cookware\cookware.exe
GroupAdd, groupCook, ahk_exe cookware.exe
if WinActive("ahk_exe cookware.exe")
	GroupActivate, groupCook, r
else
	WinActivate ahk_exe cookware.exe ;you have to use WinActivatebottom if you didn't create a window group.
Return
}

switchToExcel()
{
IfWinNotExist, ahk_class XLMAIN
	Run, excel.exe
GroupAdd, groupExcels, ahk_class XLMAIN
if WinActive("ahk_exe excel.exe")
	GroupActivate, groupExcels, r
else
	WinActivate ahk_class XLMAIN ;you have to use WinActivatebottom if you didn't create a window group.
Return
}

switchToWord()
{
IfWinNotExist, ahk_exe WINWORD.EXE
	Run, WINWORD.EXE
GroupAdd, groupWord, ahk_class OpusApp
if WinActive("ahk_exe WINWORD.EXE")
	GroupActivate, groupWord, r
else
	WinActivate ahk_exe WINWORD.EXE ;you have to use WinActivatebottom if you didn't create a window group.
Return
}

GetJobPath()
{
global
WinGetTitle, wintitle, A
RegExMatch(wintitle, "(\d{6,7})", match)
if match2
   msgbox, There were multiple matches
if match1
{
    Loop, Files, S:\Projects\%match%*, D
    {
    jobpath := A_LoopFileFullPath
    Break
    }
    Return jobpath
}
else
{
    send, ^+{Left}
    send, ^c
    match := Clipboard
     Loop, Files, S:\Projects\%match%*, D
    {
    jobpath := A_LoopFileFullPath
    Break
    }
return jobpath
}
}

GetJobNumber()
{
global
WinGetTitle, wintitle, A
RegExMatch(wintitle, "(\d{6,7})", match)
if match2
   msgbox, There were multiple matches
if match1
{
    jobnumber := match1
	Return jobnumber
}
}
return





