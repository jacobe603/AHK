#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; Working in script Y | Working when called
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance force
SetWorkingDir %A_ScriptDir%


;Ini File for RC Menu Variable Transfer
ComINIPath=..\Data\Search.ini

IniRead,job_name,%ComINIPath%,CurrentProject,JobName
IniRead,job_no,%ComINIPath%,CurrentProject,JobNumber
IniRead,CustPO1,%ComINIPath%,ProjectData,CustPO1
IniRead,address1,%ComINIPath%,CurrentAddress,AField1
IniRead,address2,%ComINIPath%,CurrentAddress,AField2
IniRead,address3,%ComINIPath%,CurrentAddress,AField3
IniRead,city,%ComINIPath%,CurrentAddress,CityField
IniRead,state,%ComINIPath%,CurrentAddress,StField
IniRead,zip,%ComINIPath%,CurrentAddress,ZipField
IniRead,STContact,%ComINIPath%,CurrentAddress,ContactField
IniRead,STPhone,%ComINIPath%,CurrentAddress,PhoneField

Gui, Add, ListBox, x10 y20 h20, %job_no%
Gui, Add, ListBox, x10 y40 h20, %job_name%
Gui, Add, ListBox, x10 y60 h20, %address1%
Gui, Add, ListBox, x10 y80 h20, %address2%
Gui, Add, ListBox, x10 y100 h20, %address3%
;Gui, Add, Edit, border vNcity x10 y118 w120, %city%
Gui, Add, ListBox, x10 y120 h20, %city%
Gui, Add, ListBox, x10 y140 h20, %state%
Gui, Add, ListBox, x10 y160 h20, %zip%
Gui, Add, ListBox, x10 y180 h20, %STContact%
Gui, Add, ListBox, x10 y200 h20, %STPhone%
;Gui, Add, ListBox, x10 y220 h20, %CustPO1%

Gui, Add, Text, x150 y20 h20 , CTRL+Numpad1
Gui, Add, Text, x150 y40 h20 , CTRL+Numpad2
Gui, Add, Text, x150 y60 h20 , CTRL+Numpad3
Gui, Add, Text, x150 y80 h20 , CTRL+Numpad4
Gui, Add, Text, x150 y100 h20 , CTRL+Numpad5
Gui, Add, Text, x150 y120 h20 , CTRL+Numpad6
Gui, Add, Text, x150 y140 h20 , CTRL+Numpad7
Gui, Add, Text, x150 y160 h20 , CTRL+Numpad8
Gui, Add, Text, x150 y180 h20 , CTRL+Numpad9
Gui, Add, Text, x150 y200 h20 , CTRL+Numpad0

Gui,+AlwaysOnTop
Gui, +resize
; Generated using SmartGUI Creator for SciTE
Gui, Show, w250 h265, %job_no% - %job_name%
Gui, Add, Button, gCopyAdd x150 y240 w80, &Copy Address
return

GuiClose:
ExitApp

^Numpad1::
clipboard := job_no
send ^v
return

^Numpad2::
clipboard := job_name
send ^v
return

^Numpad3::
clipboard := address1
send ^v
return

^Numpad4::
clipboard := address2
send ^v
return

^Numpad5::
clipboard := address3
send ^v
return

^Numpad6::
clipboard := City
send ^v
return

^Numpad7::
clipboard := State
send ^v
return

^Numpad8::
clipboard := zip
send ^v
return

^Numpad9::
clipboard := STContact
send ^v
return

^Numpad0::
clipboard := STPhone
send ^v
return

CopyAdd:
gui, submit
clipboard = %job_name%`r`n%address1%`r`n%city%, %State%  %zip%
Gui, Show, w250 h265, %job_no% - %job_name%
return