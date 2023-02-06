#SingleInstance force

toggle := 1
if (!A_IsCompiled && !InStr(A_AhkPath, "_UIA.exe") && toggle) {
	newPath := RegExReplace(A_AhkPath, "\.exe", "U" (32 << A_Is64bitOS) "_UIA.exe")
	Run % StrReplace(DllCall("Kernel32\GetCommandLine", "Str"), A_AhkPath, newPath)
	ExitApp
}

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.

SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;^ = ctrl
;+ = Shift
;! = alt
;# = Windows key

user = %A_UserName%
Run, U:\Projects\AHK-git\AHK\Standard Script\Scripts\JN as Typed.ahk
#Include Scripts\Win_Improvements.ahk ;General improvements to Windows experience
#Include Scripts\Work.ahk ;Specific improvements to using systems at work
#Include Scripts\TextMath.ahk ;Do math on text anywhere in Windows
#Include, Scripts\Text_Expand.ahk ;Very specific to jacobe text expansions
;#Include, Scripts\JN as Typed.ahk ;Grab JN's as they are typed in

;Run "C:\Users\%user%\OneDrive - Schwab Vollhaber Lubratt Inc\Desktop\AHK - Local\Go To Job - OB v1.ahk"



^+!Insert:: ;close all AHK scripts and any other windows with .ahk in them
DetectHiddenWindows, On
SetTitleMatchMode, 2
Loop {
   WinClose, .ahk
   IfWinNotExist, .ahk
      Break     ;No [more] matching windows found
}
return

















