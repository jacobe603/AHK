;Job Commander v0.1 by Jacob Erbele 07/05/2018
#SingleInstance force
;if (!A_IsCompiled && !InStr(A_AhkPath, "_UIA.exe")) {
;	newPath := RegExReplace(A_AhkPath, "\.exe", "U" (32 << A_Is64bitOS) "_UIA.exe")
;	Run % StrReplace(DllCall("Kernel32\GetCommandLine", "Str"), A_AhkPath, newPath)
;	ExitApp
;}
;Menu, Tray, Icon, shell32.dll, 20
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.


; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
Menu, Tray, Icon, shell32.dll, 101

msgBox,0,Job Commander v0.1, Job Commander is now running in the background.  Double right click to activate.,10

{ ;INI variables
ComINIPath=%A_ScriptDir%\JobCommander.ini
IniRead,FormName1,%ComINIPath%,Forms,FormName1
IniRead,FormLocation1,%ComINIPath%,Forms,FormLocation1
IniRead,FormName2,%ComINIPath%,Forms,FormName2
IniRead,FormLocation2,%ComINIPath%,Forms,FormLocation2
IniRead,FormName3,%ComINIPath%,Forms,FormName3
IniRead,FormLocation3,%ComINIPath%,Forms,FormLocation3
IniRead,FormName4,%ComINIPath%,Forms,FormName4
IniRead,FormLocation4,%ComINIPath%,Forms,FormLocation4
IniRead,FormName5,%ComINIPath%,Forms,FormName5
IniRead,FormLocation5,%ComINIPath%,Forms,FormLocation5
IniRead,FormName6,%ComINIPath%,Forms,FormName6
IniRead,FormLocation6,%ComINIPath%,Forms,FormLocation6
IniRead,FormName7,%ComINIPath%,Forms,FormName7
IniRead,FormLocation7,%ComINIPath%,Forms,FormLocation7
IniRead,FormName8,%ComINIPath%,Forms,FormName8
IniRead,FormLocation8,%ComINIPath%,Forms,FormLocation8

IniRead,SoftName1,%ComINIPath%,Selection,SoftName1
IniRead,SoftLocation1,%ComINIPath%,Selection,SoftLocation1
IniRead,SoftName2,%ComINIPath%,Selection,SoftName2
IniRead,SoftLocation2,%ComINIPath%,Selection,SoftLocation2
IniRead,SoftName3,%ComINIPath%,Selection,SoftName3
IniRead,SoftLocation3,%ComINIPath%,Selection,SoftLocation3
IniRead,SoftName4,%ComINIPath%,Selection,SoftName4
IniRead,SoftLocation4,%ComINIPath%,Selection,SoftLocation4
IniRead,SoftName5,%ComINIPath%,Selection,SoftName5
IniRead,SoftLocation5,%ComINIPath%,Selection,SoftLocation5

IniRead,MfgName1,%ComINIPath%,Manufacturer,MfgName1
IniRead,MfgName2,%ComINIPath%,Manufacturer,MfgName2

IniRead,ContactName1,%ComINIPath%,Contractors,ContactName1
IniRead,ContactPhone1,%ComINIPath%,Contractors,ContactPhone1
IniRead,ContactName2,%ComINIPath%,Contractors,ContactName2
IniRead,ContactPhone2,%ComINIPath%,Contractors,ContactPhone2
IniRead,ContactName3,%ComINIPath%,Contractors,ContactName3
IniRead,ContactPhone3,%ComINIPath%,Contractors,ContactPhone3
IniRead,ContactName4,%ComINIPath%,Contractors,ContactName4
IniRead,ContactPhone4,%ComINIPath%,Contractors,ContactPhone4
IniRead,ContactName5,%ComINIPath%,Contractors,ContactName5
IniRead,ContactPhone5,%ComINIPath%,Contractors,ContactPhone5
IniRead,ContactName6,%ComINIPath%,Contractors,ContactName6
IniRead,ContactPhone6,%ComINIPath%,Contractors,ContactPhone6

IniRead,ContactFName1,%ComINIPath%,FargoPhone,ContactFName1
IniRead,ContactFPhone1,%ComINIPath%,FargoPhone,ContactFPhone1
IniRead,ContactFName2,%ComINIPath%,FargoPhone,ContactFName2
IniRead,ContactFPhone2,%ComINIPath%,FargoPhone,ContactFPhone2
IniRead,ContactFName3,%ComINIPath%,FargoPhone,ContactFName3
IniRead,ContactFPhone3,%ComINIPath%,FargoPhone,ContactFPhone3
IniRead,ContactFName4,%ComINIPath%,FargoPhone,ContactFName4
IniRead,ContactFPhone4,%ComINIPath%,FargoPhone,ContactFPhone4
IniRead,ContactFName5,%ComINIPath%,FargoPhone,ContactFName5
IniRead,ContactFPhone5,%ComINIPath%,FargoPhone,ContactFPhone5
IniRead,ContactFName6,%ComINIPath%,FargoPhone,ContactFName6
IniRead,ContactFPhone6,%ComINIPath%,FargoPhone,ContactFPhone6


IniRead,QuickName1,%ComINIPath%,QuickMenu,QuickName1
IniRead,QuickLocation1,%ComINIPath%,QuickMenu,QuickLocation1
IniRead,QuickName2,%ComINIPath%,QuickMenu,QuickName2
IniRead,QuickLocation2,%ComINIPath%,QuickMenu,QuickLocation2
IniRead,QuickName3,%ComINIPath%,QuickMenu,QuickName3
IniRead,QuickLocation3,%ComINIPath%,QuickMenu,QuickLocation3
IniRead,QuickName4,%ComINIPath%,QuickMenu,QuickName4
IniRead,QuickLocation4,%ComINIPath%,QuickMenu,QuickLocation4
IniRead,QuickName5,%ComINIPath%,QuickMenu,QuickName5
IniRead,QuickLocation5,%ComINIPath%,QuickMenu,QuickLocation5
}


; Form Submenu Items
Menu, Submenu1, Add, %FormName1%, MenuHandler2
Menu, Submenu1, Add, %FormName2%, MenuHandler3
Menu, Submenu1, Add, %FormName3%, MenuHandler4
Menu, Submenu1, Add, %FormName4%, MenuHandler5
Menu, Submenu1, Add, %FormName5%, MenuHandler6
Menu, Submenu1, Add, %FormName6%, MenuHandler7
Menu, Submenu1, Add, %FormName7%, MenuHandler8
Menu, Submenu1, Add, %FormName8%, MenuHandler9

; Selection Submenu Items
Menu, Submenu2, Add, %SoftName1%, MenuHandlerA1
 ;Menu, Submenu2, Icon, %SoftName1%, %SoftLocation1%
Menu, Submenu2, Add, %SoftName2%, MenuHandlerA2
 ; Menu, Submenu2, Icon, %SoftName2%, %SoftLocation2%
Menu, Submenu2, Add, %SoftName3%, MenuHandlerA3
Menu, Submenu2, Add, %SoftName4%, MenuHandlerA4
  ;Menu, Submenu2, Icon, %SoftName4%, %SoftLocation4%
Menu, Submenu2, Add, %SoftName5%, MenuHandlerA5
  ;Menu, Submenu2, Icon, %SoftName5%, %SoftLocation5%


; Manufacturer Submenu4 Items
Menu, Submenu4, Add, %MfgName1%, MenuHandlerC1
Menu, Submenu4, Add, %MfgName2%, MenuHandlerC2

; Fargo Submenu5 Items
Menu, Submenu5, Add, %ContactFName1%, MenuHandlerD1
Menu, Submenu5, Add, %ContactFName2%, MenuHandlerD2
Menu, Submenu5, Add, %ContactFName3%, MenuHandlerD3
Menu, Submenu5, Add, %ContactFName4%, MenuHandlerD4
Menu, Submenu5, Add, %ContactFName5%, MenuHandlerD5
Menu, Submenu5, Add, %ContactFName6%, MenuHandlerD6

; Contractors Submenu6 Items
Menu, Submenu6, Add, %ContactName1%, MenuHandlerE1
Menu, Submenu6, Add, %ContactName2%, MenuHandlerE2
Menu, Submenu6, Add, %ContactName3%, MenuHandlerE3
Menu, Submenu6, Add, %ContactName4%, MenuHandlerE4
Menu, Submenu6, Add, %ContactName5%, MenuHandlerE5
Menu, Submenu6, Add, %ContactName6%, MenuHandlerE6

Menu, FormSubmenu1, Add, Forms, formsOPEN
Menu, FormSubmenu1, Add, Projects, projectsOPEN
Menu, FormSubmenu1, Add, Restored S Drive, restoredOPEN

; Contacts Submenu
Menu, Submenu3, Add, Manufacturers, :Submenu4
Menu, Submenu3, Add, Contractors, :Submenu6
Menu, Submenu3, Add, Fargo, :Submenu5

; --------------------------Main Menu - Subs---------------------------------
Menu, MyMenu, Add, Insert Forms, :Submenu1
Menu, MyMenu, Add, Selection Software, :Submenu2
Menu, MyMenu, Add, Contacts, :Submenu3
Menu, MyMenu, Add
Menu, MyMenu, Add, Open Docs/Forms, :FormSubmenu1

; Quick Access Items
Menu, MyMenu, Add  ; Add a separator line below the submenu.
Menu, MyMenu, Add, Job Manager, MenuHandlerB1
  ;Menu, MyMenu, Icon, Job Manager, C:\Program Files (x86)\SVL\Job Manager\JobManager.exe
Menu, MyMenu, Add, Rep Order Book, MenuHandlerB2
  ;Menu, MyMenu, Icon, Rep Order Book, C:\Program Files (x86)\RepOrderBook\RepOrderBook.exe
Menu, MyMenu, Add, MXIE, MenuHandlerB3
  ;Menu, MyMenu, Icon, MXIE, C:\Program Files (x86)\Zultys\MXIE\Bin\mxie.exe
Menu, MyMenu, Add, BlueBeam, MenuHandlerB9
  ;Menu, MyMenu, Icon, BlueBeam, C:\Program Files\Bluebeam Software\Bluebeam Revu\2019\Revu\Revu.exe
Menu, MyMenu, Add, Outlook, MenuHandlerB10
  ;Menu, MyMenu, Icon, Outlook, C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE

Menu, MyMenu, Add
Menu, MyMenu, Add, %QuickName1%, MenuHandlerB4
  ;Menu, MyMenu, Icon, %QuickName1%, %QuickLocation1%
Menu, MyMenu, Add, %QuickName2%, MenuHandlerB5
  ;Menu, MyMenu, Icon, %QuickName2%, %QuickLocation2%
Menu, MyMenu, Add, %QuickName3%, MenuHandlerB6
  ;Menu, MyMenu, Icon, %QuickName3%, %QuickLocation3%
Menu, MyMenu, Add, %QuickName4%, MenuHandlerB7
Menu, MyMenu, Add, %QuickName5%, MenuHandlerB8
 ;Menu, MyMenu, Icon, %QuickName5%, %QuickLocation5%
return  ; End of scripts auto-execute section.




{ ;Submenu1 - Forms
MenuHandler1:
MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
return

MenuHandler2:
^!+4::
Blankfile = %FormLocation1%
Gosub, CopyToCurrent
return

MenuHandler3:
Blankfile = %FormLocation2%
Gosub, CopyToCurrent
return

MenuHandler4:
^!+3::
Blankfile = %FormLocation3%
Gosub, CopyToCurrent
return

MenuHandler5:
Blankfile = %FormLocation4%
Gosub, CopyToCurrent
return

MenuHandler6:
Blankfile = %FormLocation5%
Gosub, CopyToCurrent
return

MenuHandler7:
Blankfile = %FormLocation6%
Gosub, CopyToCurrent
return

MenuHandler8:
Blankfile = %FormLocation7%
Gosub, CopyToCurrent
return

MenuHandler9:
Blankfile = %FormLocation8%
Gosub, CopyToCurrent
return
}

{ ;Submenu2 - Selection Software
MenuHandlerA1:
  RunExist(SoftLocation1)
  ;Run, %SoftLocation1%
Return

MenuHandlerA2:
  RunExist(SoftLocation2)
  ;Run, %SoftLocation2%
Return

MenuHandlerA3:
  RunExist(SoftLocation3)
  ;Run %SoftLocation3%
Return

MenuHandlerA4:
  RunExist(SoftLocation4)
  ;Run, %SoftLocation4%
Return

MenuHandlerA5:
  RunExist(SoftLocation5)
  ;Run, %SoftLocation5%
Return

}

{ ;manufacturer Contacts
MenuHandlerC1: ;Bulldog
  Run, C:\Program Files (x86)\Dial\dial.exe 8882205551
Return

MenuHandlerC2: ;Ruskin
  Run, https://www.ruskin.com/VIPNet/ContactUs
Return
}

{ ;Fargo Contacts
MenuHandlerD1:
  Run, C:\Program Files (x86)\Dial\dial.exe %ContactFPhone1%
Return
MenuHandlerD2:
  Run, C:\Program Files (x86)\Dial\dial.exe %ContactFPhone2%
Return
MenuHandlerD3:
  Run, C:\Program Files (x86)\Dial\dial.exe %ContactFPhone3%
Return
MenuHandlerD4:
  Run, C:\Program Files (x86)\Dial\dial.exe %ContactFPhone4%
Return
MenuHandlerD5:
  Run, C:\Program Files (x86)\Dial\dial.exe %ContactFPhone5%
Return
MenuHandlerD6:
  Run, C:\Program Files (x86)\Dial\dial.exe %ContactFPhone6%
Return
}

{ ;Contractor Contacts
MenuHandlerE1:
  Run, C:\Program Files (x86)\Dial\dial.exe %ContactPhone1%
Return
MenuHandlerE2:
  Run, C:\Program Files (x86)\Dial\dial.exe %ContactPhone2%
Return
MenuHandlerE3:
  Run, C:\Program Files (x86)\Dial\dial.exe %ContactPhone3%
Return
MenuHandlerE4:
  Run, C:\Program Files (x86)\Dial\dial.exe %ContactPhone4%
Return
MenuHandlerE5:
  Run, C:\Program Files (x86)\Dial\dial.exe %ContactPhone5%
Return
MenuHandlerE6:
  Run, C:\Program Files (x86)\Dial\dial.exe %ContactPhone6%
Return
}

{ ;Forms & Docs
formsOPEN:
  Run, X:\Fargo\Inside Sales\Forms
return

projectsOPEN:
Run, C:\Users\jacob\OneDrive - Schwab Vollhaber Lubratt Inc\Project Folders
return

restoredOPEN:
Run, C:\Users\jacob\OneDrive - Schwab Vollhaber Lubratt Inc\Restored S-Drive Files\Jacob Erbele
return

}

{ ;Quick Access Items
MenuHandlerB1:
  Run, https://objm.svl.com/#/dashboard/job-search?newWindow=true
  ;Run, C:\Program Files (x86)\SVL\Job Manager\JobManager.exe
Return

MenuHandlerB2:
 Run, https://objm.svl.com
 sleep, 6000
 send, {Tab 7}
 send, {Down}

Return

MenuHandlerB3:
  RunExist("C:\Program Files (x86)\Zultys\MXIE\Bin\mxie.exe")
Return

MenuHandlerB9:
  RunExist("C:\Program Files\Bluebeam Software\Bluebeam Revu\2018\Revu\Revu.exe")
Return

MenuHandlerB10:
  RunExist("C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE")
Return

MenuHandlerB4:
  Run, %QuickLocation1%
Return

MenuHandlerB5:
  Run, %QuickLocation2%
Return

MenuHandlerB6:
  Run, %QuickLocation3%
Return

MenuHandlerB7:
  Run, %QuickLocation4%
Return

MenuHandlerB8:
  Run, %QuickLocation5%
Return
}


RunExist(SoftLocation)
{
  IfWinExist, ahk_exe %SoftLocation%
  WinActivate
else
  Run, %SoftLocation%
  if (ErrorLevel = 1)
    MsgBox File does not exist.  Double check ini file.
  return
}



CopyToCurrent:
Send !d
clipboard =  ; Start off empty to allow ClipWait to detect when the text has arrived.
Send ^c
ClipWait  ; Wait for the clipboard to contain text.
Send {Enter}
Filecopy %Blankfile%, %Clipboard%
return

; Short Cut to Open Menu
~RButton::
If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 500)
{
Sleep 200 ; wait for right-click menu, fine tune for your PC
Send {Esc} ; close it
Menu, MyMenu, Show
}
Return


; TEST
;^+c::
;clipboard =  ; Start off empty so ClipWait can detect when the copy completes
;Send ^c
;ClipWait
;FileAppend, %ClipboardAll%, C:\clipboard.clp
;return

;^+v::
;FileRead, Clipboard, *c C:\clipboard.clp ; *c must precede the filename
;Send ^v
;return


MenuHandler:
CheckmarkToggle( A_ThisMenuItem, A_ThisMenu )
msgbox %TestFlag%
return


CheckmarkToggle(MenuItem, MenuName){
    global
    %MenuItem%Flag := !%MenuItem%Flag ; Uses ! (Logical-not) to toggle between 0 and 1
    If (%MenuItem%Flag)
        Menu, %MenuName%, Check, %MenuItem%
    else
        Menu, %MenuName%, UnCheck, %MenuItem%
}


; Hotkeys
^Numpad1::Copy(1)
^Numpad4::Paste(1)

^Numpad2::Copy(2)
^Numpad5::Paste(2)

^Numpad3::Copy(3)
^Numpad6::Paste(3)


Copy(clipboardID)
{
	global ; All variables are global by default
	local oldClipboard := ClipboardAll ; Save the (real) clipboard

	Clipboard = ; Erase the clipboard first, or else ClipWait does nothing
	Send ^c
	ClipWait, 2, 1 ; Wait 1s until the clipboard contains any kind of data
	if ErrorLevel
	{
		Clipboard := oldClipboard ; Restore old (real) clipboard
		return
	}

	ClipboardData%clipboardID% := ClipboardAll
	ClipboardText%clipboardID% := Clipboard

	Clipboard := oldClipboard ; Restore old (real) clipboard
 }

Cut(clipboardID)
{
	global ; All variables are global by default
	local oldClipboard := ClipboardAll ; Save the (real) clipboard

	Clipboard = ; Erase the clipboard first, or else ClipWait does nothing
	Send ^x
	ClipWait, 2, 1 ; Wait 1s until the clipboard contains any kind of data
	if ErrorLevel
	{
		Clipboard := oldClipboard ; Restore old (real) clipboard
		return
	}
	ClipboardData%clipboardID% := ClipboardAll

	Clipboard := oldClipboard ; Restore old (real) clipboard
}

Paste(clipboardID)
{
	global
	local oldClipboard := ClipboardAll ; Save the (real) clipboard

	Clipboard := ClipboardData%clipboardID%
	Send ^v

	Clipboard := oldClipboard ; Restore old (real) clipboard
	oldClipboard =
}

Explorer_GetSelection() {
   WinGetClass, winClass, % "ahk_id" . hWnd := WinExist("A")
   if !(winClass ~="Progman|WorkerW|(Cabinet|Explore)WClass")
      Return

   shellWindows := ComObjCreate("Shell.Application").Windows
   if (winClass ~= "Progman|WorkerW")
      shellFolderView := shellWindows.FindWindowSW(0, 0, SWC_DESKTOP := 8, 0, SWFO_NEEDDISPATCH := 1).Document
   else {
      for window in shellWindows
         if (hWnd = window.HWND) && (shellFolderView := window.Document)
            break
   }
   for item in shellFolderView.SelectedItems
	 ; result .= (result = "" ? "" : "`n") . item.Path
	result := shellFolderView.Folder.Self.Path
   if !result
  	  result := shellFolderView.Folder.Self.Path

   Return result
}
