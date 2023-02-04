#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 50


var := StrSplit(clipboard, A_Tab)
    ;JobNo :=    var[1]
	Mfg :=  	var[2]
    Desc := 	var[6]
    CustPO:=	var[7]
    ;SVLPO :=	var[1]
	EquipType:= var[5]
	Qty :=		var[9]
	Volume :=	var[10]
	Cost :=		var[11]
	;ShipDate :=	var[15]
Send, %Mfg%
SendInput, `t
;Input, x, , {ESC}
SendInput, %Desc%
SendInput, `t
SendInput, ^a
SendInput, %CustPO%
SendInput, `t
SendInput, ^a
;SendInput, %SVLPO%
SendInput, `t
SendInput, %EquipType%
SendInput, `t
SendInput, ^a
SendInput, %Qty%
SendInput, `t
Send %Volume%
SendInput, `t
Send %Cost%
SendInput, {Tab 2}
Verdant	(3) VX-TR-KT-W-HRT BULLDOG TSAT	
;SendInput, {Enter}
;SendInput, %SVLPO%%MFG%12  This does work for PO number
