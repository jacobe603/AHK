#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%
#SingleInstance force

;Create GUI to make new Outlook Task

Gui, Add, Text, x10 y15 w60 h20, StartDate:
Gui, Add, DateTime, x+10 y10 wp20 h20 vStartDate Section,

Gui, Add, Text, x10 y55 w70 h20 , StartTime:
Gui, Add, DateTime, x+10 y50 wp hp vStartTime 1, HH:mm:ss ;time

Gui, Add, Text, x10 y105 w70 h20 , Task Subject:
Gui, Add, Edit, vTaskSubj x90 y100 w200


Gui, Add, Button, x10 y170 w60 h20 Default, &OK
Gui, Show, w300 h200, AHK-Scheduler

;Gui, Add, DateTime, vMyDateTime, LongDate
;Gui, Add, Button, default, OK
Gui, Show
return



ButtonOK:
Gui, Submit
olApp := ComObjActive("Outlook.Application")
olTask := olApp.CreateItem(3)
olTask.ReminderSet := true
olTask.Status := 1
msgbox, %TaskSubj%
olTask.Subject := TaskSubj

olTask.Display(true)

GUI, Destroy
;ListVars
return