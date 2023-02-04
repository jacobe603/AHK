#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 50


var := StrSplit(clipboard, A_Tab)
	GuiCode := var[28]
	if (instr(GuiCode,"POGUI"))
	{
		;JobNo :=    var[1]
		SVLPO :=	var[1]
		RegExMatch(SVLPO,"\D{3}",mfgCode)

		Desc := 	var[12]
		;CustPO:=	var[3]  ;3
		;EquipType:= var[7]  ;4
		Qty :=		var[11]  ;5
		;Volume :=	var[9]  ;6
		Cost :=		var[26] ;7
		;ShipDate :=	var[15] ;8
	}

	else if var.MaxIndex() > 20
	{
		JobNo :=    var[1]
		mfgCode :=  var[26] ;1
		Desc := 	var[19] ;2
		CustPO:=	var[3]  ;3
		SVLPO :=	var[1]
		EquipType:= var[7]  ;4
		Qty :=		var[8]  ;5
		Volume :=	var[9]  ;6
		Cost :=		var[10] ;7
		ShipDate :=	var[15] ;8
	}

	else
	{
		;JobNo :=    var[1]
		mfgCode := 	var[10] ;1
		Desc := 	var[11] ;2
		CustPO:=	var[1]  ;3
		;SVLPO :=	var[1]
		EquipType:= var[3]  ;4
		Qty :=		var[4]  ;5
		Volume :=	var[5]  ;6
		Cost :=		var[12] ;7
		ShipDate :=	var[7] ;8
	}



Send, %mfgCode%
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
SendInput, ^a
SendInput, %ShipDate%
SendInput, {Enter}
SendInput, {Tab 9}
;SendInput, {Enter}
;SendInput, %SVLPO%%MFG%12  This does work for PO number

ExitApp
