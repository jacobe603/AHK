#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance force

now := A_Now
FormatTime, now, %now%, MM/dd/yyyy

Gui, Add, Text,, Press Esc key to Exit App
Gui, Add, Text,, Start Date:
Gui, Add, DateTime, vStart
Gui, Add, Text,, Date after this number of weeks:

Gui, Add, Edit, vAgo
GuiControl, Focus, Ago
Gui, Add, Button, default, OK
Gui, Show,, Date Adder
return

;GuiClose:
ButtonOK:
Gui, Submit  ; Save the input from the user to each control's associated variable.
Old_Date := Start
FormatTime, Old_Date, %Old_Date%, MM/dd/yyyy
Start += Ago*7, days
FormatTime, Start, %Start%, MM/dd/yyyy
MsgBox ,4,Date Adder, %Ago% Weeks Added To %Old_Date% is %Start% `n`r` `n`r` Press Esc Key to Exit App `n`r `n`r` Copy Date to Clipboard??
ifMsgbox Yes
	clipboard = %Start%
else
	Reload
return  ; End of auto-execute section. The script is idle until the user does something.

ExitApp
esc::exitapp
GuiClose:
ExitApp
