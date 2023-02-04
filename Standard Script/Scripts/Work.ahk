; Specific Improvements for work


; Revisions to Filename
^!a:: ; Press [ Control + Alt + A ].
    Explorer_GetSelection(hwnd="")
    	{
			hwnd := hwnd ? hwnd : WinExist("A")
        	WinGetClass class, ahk_id %hwnd%
        	if (class="CabinetWClass" or class="ExploreWClass" or class="Progman")
        		for window in ComObjCreate("Shell.Application").Windows
        			if (window.hwnd==hwnd)
            CurrentWindow := window.Document.SelectedItems
    	    for item in CurrentWindow
    		{
				CopyNumber := 1
				Selectedfile := ""
    	    	Selectedfile := item.path
    	    	SplitPath, % Selectedfile,name, dir, ext, name_no_ext
                OrigName := SubStr(name_no_ext,3)
				Loop
				{
				    if FileExist( dir "\R" CopyNumber " " name_no_ext "." ext )
				    	CopyNumber += 1
				    else
					{
				        FileCopy, % Selectedfile, %dir%\R%CopyNumber% %OrigName%.%ext%
					    Break
					}
			    }
    	    }
    	}
Return

;Open Notepad on press or Notepad++ with hold
F1::
    KeyWait, F1, T0.2	; 200 ms
	If (!ErrorLevel)
        {
        if WinExist("ahk_class Notepad") or WinExist("ahk_class" . ClassName)
            WinActivate  ; Uses the last found window.
        else
                Run "C:\WINDOWS\system32\notepad.exe"
		}
	else {
		while getkeystate("F1", "p") {
			run, "C:\Program Files\Notepad++\notepad++.exe"
			sleep 50
		}
	}
return

;Outlook - Open Color Categories
^F9::
SendInput !JTGA
return

;Project Data in Word
^#4:: ;Paste project name from WinTitle
    WinGetTitle, Title, A
    FoundPos := RegExMatch(Title, "-\s(.*)\.", SubPat1)
    Clipboard = %SubPat11%
    Send, ^v
Return

^#5:: ;Paste 6 or 7 digit number from WinTitle
    WinGetTitle, Title, A
    FoundPos := RegExMatch(Title, "\d{6,7}", SubPat)
    Clipboard = %SubPat%
    Send, ^v
Return

;Open Projects folder when CTRL + SHIFT + p is pressed
^#p::
  Run, "S:\Projects"
Return

;Open Literature folder when CTRL + SHIFT + l is pressed
^#l::
  Run, "L:\"
Return

;call number using Teams
^F12::
send, ^c
ClipWait
param := "tel://"Clipboard
Run, %param%
return

;## BlueBeam Page Control ##

;CTRL + Mouse Back = Previous Page
^XButton1::
send ^{Left}
Return

;CTRL + Mouse Forward = Next Page
^XButton2::
send ^{Right}
Return

;CTRL + Right Button on Mouse = Maximize Page
^RButton::
send {Ctrl down}49{Ctrl Up}
Return

