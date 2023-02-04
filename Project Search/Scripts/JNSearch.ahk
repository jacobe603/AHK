#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;Ini File for RC Menu Variable Transfer
    ComINIPath=..\Data\Search.ini

jobnumber := A_Args[1]

    Loop, Files, S:\Projects\%jobnumber%*, D
    {
    fullpath := A_LoopFileFullPath
	jobpath := A_LoopFileName
	jobname := SubStr(jobpath, 10)
    msgbox, %fullpath%
    Break
    }

	IfExist, %fullpath%   ; If the file exists and has a path, open the file directory
                    IniWrite,%jobname%,%ComINIPath%,CurrentProject,FolderName
                    IniWrite,%fullpath%,%ComINIPath%,CurrentProject,FolderPath
                    IniWrite,%fullpath%\Orders,%ComINIPath%,CurrentProject,OrdersFolderPath
                    IniWrite,%fullpath%\Submittals,%ComINIPath%,CurrentProject,SubmittalsFolderPath
                    IniWrite,%fullpath%\Quotes,%ComINIPath%,CurrentProject,QuotesFolderPath
                    IniWrite,%fullpath%\Pricing,%ComINIPath%,CurrentProject,PricingFolderPath
                    IniWrite,%jobno%,%ComINIPath%,CurrentProject,JobNumber
                    IniWrite,%jobName%,%ComINIPath%,CurrentProject,JobName

	Return





