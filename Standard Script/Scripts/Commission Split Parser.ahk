#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Event
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force
SetKeyDelay, 150

; Data needs to look like this Eis,E,5	Foss,E,5	John,E,7	Eis,C,39	Hep,C,39	Erb,C,5


Colors := Clipboard
Loop, parse, Colors, `t
{
    Loop, parse, A_Loopfield, `,
	{	
		Send, %A_LoopField%
		Sleep, 50
		Send, {Tab}
		;Input, var, V, {Esc}
		Sleep, 50
		If A_Index = 3
		{
			Send, {Enter}
		}
	}	
}
return