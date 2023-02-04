#SingleInstance force

if (!A_IsCompiled && !InStr(A_AhkPath, "_UIA.exe")) {
	newPath := RegExReplace(A_AhkPath, "\.exe", "U" (32 << A_Is64bitOS) "_UIA.exe")
	Run % StrReplace(DllCall("Kernel32\GetCommandLine", "Str"), A_AhkPath, newPath)
	ExitApp
}

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.

SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;^ = ctrl
;+ = Shift
;! = alt
;# = Windows key

user = %A_UserName%

#Include Scripts\Win_Improvements.ahk
#Include Scripts\SVL.ahk
#Include Scipts\TextMath.ahk

;Run "C:\Users\%user%\OneDrive - Schwab Vollhaber Lubratt Inc\Desktop\AHK - Local\Go To Job - OB v1.ahk"
Run "C:\Users\%user%\OneDrive - Schwab Vollhaber Lubratt Inc\Desktop\AHK - Local\JN as Typed.ahk"


^+!Insert:: ;close all AHK scripts
DetectHiddenWindows, On
SetTitleMatchMode, 2
Loop {
   WinClose, .ahk
   IfWinNotExist, .ahk
      Break     ;No [more] matching windows found
}
return

















