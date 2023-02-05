#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance force
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2

menu, tray, icon, shell32.dll, 210
WinGetPos, xIntPos, yIntPos,,,A

; INDEX FILE LOCATION IS IN THE USERS MY DOCUMENTS FOLDER
    INDEX_FILE := A_MyDocuments . "\FileIndex.csv"

;Ini File for RC Menu Variable Transfer
    ComINIPath=..\Data\Search.ini

; ***** FUNCTION, SECONDS TIMER ***
; Calculates the time in seconds from a starting time to ending time.
; Use the StartTimer() to store a starting time.
; Use the EndTimer() function to calculate the seconds between the current time and start time
    StartTimer()
    {
        return (A_DD * 24 * 60 * 60) + (A_Hour * 60 * 60) + (A_Min * 60) + A_Sec + ((A_MSec + 0) / 1000)
    }

    EndTimer(StartTime)
    {
        return (A_DD * 24 * 60 * 60) + (A_Hour * 60 * 60) + (A_Min * 60) + A_Sec + ((A_MSec + 0) / 1000) - StartTime
    }

; ***** FUNCTION DateInDays ***
    ; ------- Requires the custom function 'IsLeapDay' -------
    ; Returns a date & timestamp as an integer representing a
    ;   number of days that can easily be compared to another.
    DateInDays(timestamp)
    {
        If (timestamp = 0)
            timestamp := FormatTime, timestamp
        FormatTime, TM_Year, %timestamp%, yyyy
        FormatTime, TM_Month, %timestamp%, M
        FormatTime, TM_Day, %timestamp%, d

        ; Determine if leap year
            LeapDay = 0
            if(TM_Month > 2)
                LeapDay := IsLeapDay(TM_Year)

        ; Number of days in the year up to the end of the prior month
        MonthDays1 := 0
        MonthDays2 := 31
        MonthDays3 := 59
        MonthDays4 := 90
        MonthDays5 := 120
        MonthDays6 := 151
        MonthDays7 := 181
        MonthDays8 := 212
        MonthDays9 := 243
        MonthDays10 := 273
        MonthDays11 := 304
        MonthDays12 := 334
        MonthDays := MonthDays%TM_Month%

        Return (TM_Year * 365) + (MonthDays%TM_Month%) + TM_Day + LeapDay   ; return total days
    }

; Get the last search term if it was searched recently
    LAST_SEARCH_FILE := A_MyDocuments . "\LastSearch.tmp"

    IfExist, %LAST_SEARCH_FILE%
    {
        FileGetTime, LastSearchTime, %LAST_SEARCH_FILE% ; get the timestamp of the last search
        CurrentDateTime := A_YYYY . A_MM . A_DD . A_Hour . A_Min . A_Sec    ; current timestamp
        LastSearchAge := CurrentDateTime - LastSearchTime   ; get the age of the last search in seconds

        If (LastSearchAge < 900 AND LastSearchAge > 0)  ; if the last search was made less than 15 minutes (900 seconds) ago...
            FileRead, LastSearchTerm, %LAST_SEARCH_FILE%    ; get the last search term to use as the default new search
    }

; USE UN-EXACT WINDOW NAME MATCH BY DEFAULT
    SetTitleMatchMode, 2
    SendMode, Input

; SET WORKING DIRECTORY

WaitAll(x, y = 0)
; x = Window Name
; y = Window Text (optional)
{
    if y = 0
        y :=
    WinWait, %x%, %y%
    IfWinNotActive, %x%, %y%
        WinActivate, %x%, %y%
    WinWaitActive, %x%, %y%
    ; return x + y   ; "Return" expects an expression.
}
; END FUNCTION ***

; CLEAR VARIABLE WHEN SEARCH IS RUN
;   LastSearchTerm :=

; FUNCTION TO FORMAT NUMBERS with commas (ex. 43485394 = 43,485,394)
    ThousandsSep(x, s=",")
    {
       return RegExReplace(x, "\G\d+?(?=(\d{3})+(?:\D|$))", "$0" s)
    }

; ***** FUNCTION IsLeapDay ****************************
IsLeapDay(year)
{
    ; Returns a 1 if the year is a leap year.
    ; Otherwise returns a 0.
    ; Year is a 4 digit number or a variable resolving to a 4 digit number.

    if((year / 4) != ROUND(year / 4, 0))
        Return 0
    Else
        if((year / 100) != ROUND(year / 100, 0))
            Return 1
        Else
            if((year / 400) != ROUND(year / 400, 0))
                Return 0
            Else
                Return 1
}

StartThisMacro:
; Limit the number of search results
    ResultsLimit := 50

; Identify the current GUI window
    CurrentGUI := "Parameters"

; CLEAR VARIABLES EACH TIME SEARCH PARAMETERS ARE UPDATED
    SearchTerm :=
    ThisMessage :=
    Loop, 99
        SearchTerms%A_INDEX% :=
    SearchTermCount :=
    Loop, 99
        QuotePos%A_INDEX% :=
    nn :=
    Loop, 99
        Quote%A_INDEX% :=
    Loop, 99
        SearchTermQuoteArray%A_INDEX% :=
    RefPos :=
    NumQuotes :=
    QuoteA :=
    QuoteB :=
    Modifier :=
    QuoteLen :=
    QuotePos :=
    LastQuotePos :=
    SearchTermTemp :=
    ThisQuote :=
    LineNumber :=
    NotFoundA :=
    ThisFileLocationNew :=
    ThisSearch :=
    ThisSearchCheck :=
    NumberOfResults :=
    FoundMatch :=
    MatchedFileName :=
    MatchedFilePath :=
    MatchedFileDir :=
    MatchedFileExtension :=
    MatchedFileNameNoExt :=
    MatchedFileDrive :=
    NumberOfResultsFormatted :=
    BreakOutExclude :=
    SearchTerm2 :=
    ThisSearchTerm2 :=
    NewSearchTerm :=
    FilterText :=
    SearchTermLen :=
    FoundPos :=
    FoundPos2 :=
    FoundPos3 :=
    SubPat :=
    SubPatLen :=
    DateModifiedFilterDays :=
    SearchTermL :=
    SearchTermR :=

; CLEAN UP LAST SEARCH PARAMETERS
    StringRight, LastSearchTermLastChar, LastSearchTerm, 1
    If (LastSearchTermLastChar = " ")
        StringTrimRight, LastSearchTerm, LastSearchTerm, 1

;  NEW SEARCH SECTION

; Identify the current GUI window
    ;Gui, Test1:New
    CurrentGUI := "Parameters"

; SEARCH BOX
    Gui, font,s11, Segoe UI ; Set font
    Gui, Add, Edit, vSearchTerm w537, %LastSearchTerm%

; SEARCH FILES BUTTON
    Gui, font, s11, Segoe UI ; Set font
    Gui, Add, Button, gSearchFiles default w115, &Search  ; The label StartSearch (if it exists) will be run when the button is pressed.
    Gui, Add, Button, gJobsBidding w115, Jobs &Bidding
    Gui, Add, Button, gJobManager w115, &Job Manager
    Gui, Add, Button, gStkInventory w115, Stock &Inv
    Gui, Add, Button, gDateAdd w115, Date Adder
    Gui, Add, Button, gProjectTracker w115, &Proj Tracker
    Gui, Add, Button, gAaonDSO w115, &Aaon DSO Info
    Gui, Add, Button, gDaikinSO w115, &Daikin SO Info
    ; BLANK SPACE
    Gui, font, s14, Segoe UI ; Set font
        Loop, 1
            Gui, Add, Text,,

; RE-INDEX BUTTON
    Gui, font, s11, Segoe UI ; Set font
        Gui, Add, Button, x10 y400 gReIndex, &Re-Index  ; The label Re-Index (if it exists) will be run when the button is pressed.

; SHOW THE DATE/TIME OF THE LAST RE-INDEX
    FileGetTime, LastIndexTime, %INDEX_FILE%, M     ; get last modified timestamp

    ; Format modified date and time
    LastIndexDays := DateInDays(LastIndexTime)
    TodayDays := DateInDays(0)
    LastIndexDaysAgo := TodayDays - LastIndexDays
    FormatTime, TM_TIME, %LastIndexTime%, h:mm tt
    if (LastIndexDaysAgo = 0)
        LastIndexDaysAgo := "today at " . TM_TIME
    Else if (LastIndexDaysAgo = 1)
        LastIndexDaysAgo := "yesterday at " . TM_TIME
    else
        LastIndexDaysAgo .= " days ago"

    Gui, add, text,, Last indexed %LastIndexDaysAgo%    ; display modified date

    FormatTime, MatchedFileModified_DayOfYear, %MatchedFileModified%, YDay
    FormatTime, MatchedFileModified_Year, %MatchedFileModified%, yyyy
    ModifiedDays := (MatchedFileModified_Year * 365) + MatchedFileModified_DayOfYear
    FormatTime, CurrentDayOfYear,, YDay
    FormatTime, CurrentYear,, yyyy

; Old Project List
    Gui, Add, Listview, vTest1 x150 y45 w400 r12 gOldList altsubmit, Job Num|Name|LinkName|LinkPath
    Loop,11
    {
    FileReadLine, line, %A_ScriptDir%\JobList.csv, %A_Index%
    if ErrorLevel
        break
    StringSplit, item, line,`,
	LV_Add("",item1,item2,item3,item4)
    LV_ModifyCol(2,329)
    LV_ModifyCol(3,0)
    LV_ModifyCol(4,0)

    }

;MAIN GUI
    Gui, font, s11, Segoe UI ; Set font
    Gui, Add, Button, x490 y400 gBCancel, &Cancel  ; The label CANCEL (if it exists) will be run when the button is pressed.
    MENU,ToolsMenu,Add,&Update Auth Key, MenuAuthKey
    MENU,ToolsMenu,Add,&Search Jobs, SearchJobs
    Menu, MyMenuBar, Add, &Tools, :ToolsMenu
    Gui, Menu, MyMenuBar ; Attach MyMenuBar to the GUI
    Gui, Show, w570 h470, Project Search v4.0

; End of auto-execute section. The script is idle until the user does something.
return

MenuAuthKey:
run, "REST - GET AUTH.ahk"
return

SearchJobs:
run, "REST - SEARCH JOBS.ahk"
return


;functions for Double Click and Right Click on the Old Job List **SHOULD UPDATE SO IT JUST OPENS PROJECT
OldList:
    if (A_GuiEvent = "DoubleClick")
    {



    LV_GetText(RowText, A_EventInfo)  ; Get the text from the row's first field.
        Loop, Files, S:\Projects\%RowText%*, D
        {
        jobpath := A_LoopFileFullPath
        Run, %jobpath%
        ;Clipboard := A_LoopFileName
        Break
        }
    ;LV_GetText(SearchTerm, A_EventInfo)
        ;Gui, Submit  ; Save the input from the user to each control's associated variable.
    ;GUI, Destroy
    ;Goto, StartSearch
    }
;10/28/2022-------------------------------
   If A_GuiEvent = RightClick  ; IF USER RIGHT CLICKS A SEARCH RESULT, OPEN THE FILE LOCATION
    {
        GuiControlGet, FocusedControl, FocusV
        FocusedRow := LV_GetNext("", "Focused")
    ;   If no file is selected, do nothing.
        if FocusedRow = 0
            return

        LV_GetText(jobNo, FocusedRow, 1)
        LV_GetText(jobName, FocusedRow, 2)
        LV_GetText(LinkFileName, FocusedRow, 3)
        LV_GetText(LinkFilePath, FocusedRow, 4)

        Gosub, RCsub

    }
    ;if A_GuiEvent = Normal
     ;   {
     ;   msgbox, Lefty
      ;  }
;10/28/2022---------------------------

Return

SearchContacts:
    MsgBox, Search Disabled
    GUI, Destroy
    ExitApp
Return

GuiClose:
BCancel:   ; If CANCEL is clicked
GuiEscape:      ; If ESCAPE is pressed
If (CurrentGUI = "Search")  ; if in the search window, go back to the parameters window
{
    SearchInProgress = 0
    CurrentGUI = "Parameters"   ; identify the current gui window
    Goto, RefineSearch
}
ExitApp ; if not in the search window, exit
return

ReIndex:     ; REINDEX SEARCH LOCATIONS
            Run, SearchFilesReIndex.ahk
            Goto, EndMacroSearchFiles
Return

JobsBidding:
run, REST - Jobs Bidding.ahk
return

JobManager:
run, Chrome.exe https://objm.svl.com/#/dashboard/job-overview/2?newWindow=true
return

StkInventory:
run, Chrome.exe https://objm.svl.com/#/dashboard/inv-stock-inquiry?newWindow=true
return

DateAdd:
run, Date Adder.ahk
return

ProjectTracker:
run, C:\Users\%A_username%\OneDrive - Schwab Vollhaber Lubratt Inc\Documents\SVL\Project Tracker JE.xlsm
return

AaonDSO:
InputBox, DSO, AaonDSO, Please enter DSO.
run, Chrome.exe "https://www.aaon.com/rep-portal/order-status/details/%DSO%?repNumber=616&pageNumber=1"
return

DaikinSO:
InputBox, SO, DaikinSO, Please enter SO.
;MsgBox, https://sales.daikinapplied.com/Order-Detail?caller=summary&salesordernumber=%SO%
Run, chrome.exe https://sales.daikinapplied.com/Order-Summary
Sleep 2000
Send !d
SendInput https://sales.daikinapplied.com/Order-Detail?caller=summary&salesordernumber=%SO%
SendInput {enter}
return

run, Chrome.exe "https://sales.daikinapplied.com/Order-Detail?caller=summary&salesordernumber=%SO%"
return

POSearch:
 GuiControlGet, FocusedControl, FocusV
    FocusedRow := LV_GetNext("", "Focused")

    ; If no file is selected, do nothing.
        if FocusedRow = 0
            return

    ; Get the name and path of the currently selected file
        LV_GetText(LinkFileName, FocusedRow, 1)

            CurrentGUI :=                           ; Clear the CurrentGUI variable to allow the search program to completely close
            StringLeft, jobno, LinkFileName, 7      ; ADDED - Grab the job number from path
            jobno := Trim(jobno, " -")
               POSearchTerm := jobno
            ;#Include U:\AHK\PO Search\POSearch.ahk
            ;ExitApp
            ;Goto, GuiClose                          ; Close the search program

Return

Gui, Submit  ; Save the input from the user to each control's associated variable.
        Gui, Destroy
        ExitApp
    Return

RefineSearch:   ; DESTROY THE SEARCH GUI AND RETURN TO THE PARAMETERS GUI
    Cancel:
        Gui, Destroy
        Goto, StartThisMacro
Return

SearchFiles:
    Gui, Submit  ; Save the input from the user to each control's associated variable.
    GUI, Destroy
    Goto, StartSearch
Return

ButtonHiddenButtonOpenFile:     ; When user hits enter, open the selected file.
    GuiControlGet, FocusedControl, FocusV
    FocusedRow := LV_GetNext("", "Focused")

    ; If no file is selected, do nothing.
        if FocusedRow = 0
            return
    ; Get the name and path of the currently selected file
        LV_GetText(LinkFileName, FocusedRow, 1)
        LV_GetText(LinkFilePath, FocusedRow, 5)
    ; If the selected file exists, open it and then close the search program
        IfExist, %LinkFilePath%
        {
            Run, %LinkFilePath%     ; Open the selected file
            CurrentGUI :=                           ; Clear the CurrentGUI variable to allow the search program to completely close
            StringLeft, jobno, LinkFileName, 7 ; this is how to handle longer job numbers when in path name.  it just grabs 7 digits and trims off the space and - (which is probably not needed)
            jobno := Trim(jobno, " -")
            Clipboard := jobno                      ; ADDED - Place job number in clipboard.  This overwrites whatever was in clipboard before.
            Run chrome.exe "https://objm.svl.com/#/dashboard/order-inquiry-no-dollars/%jobno%?newWindow=true" " --new-window "   ; ADDED - Open new Chrome window and go to Order Book
            Goto, GuiClose                          ; Close the search program
        }
    ; If the file does not exist, display an informational message
        IfNotExist, %LinkFilePath%
            Msgbox, The file no longer exists or has been moved.
Return

^q::     ; When user hits enter, open the selected file.
    GuiControlGet, FocusedControl, FocusV
    FocusedRow := LV_GetNext("", "Focused")

    ; If no file is selected, do nothing.
        if FocusedRow = 0
            return
    ; Get the name and path of the currently selected file
        LV_GetText(LinkFileName, FocusedRow, 1)
        LV_GetText(LinkFilePath, FocusedRow, 5)
    ; If the selected file exists, open it and then close the search program
        IfExist, %LinkFilePath%
        {
            Run, %LinkFilePath%\Quotes\    ; Open the selected file
            MsgBox, %LinkFilePath%\Quotes\
            CurrentGUI :=                           ; Clear the CurrentGUI variable to allow the search program to completely close
            Goto, GuiClose                          ; Close the search program
        }
    ; If the file does not exist, display an informational message
        IfNotExist, %LinkFilePath%
            Msgbox, The file no longer exists or has been moved.
Return

^p::
^r::     ; When user hits enter, open the selected file.
    GuiControlGet, FocusedControl, FocusV
    FocusedRow := LV_GetNext("", "Focused")
    ; If no file is selected, do nothing.
        if FocusedRow = 0
            return
    ; Get the name and path of the currently selected file
        LV_GetText(LinkFileName, FocusedRow, 1)
        LV_GetText(LinkFilePath, FocusedRow, 5)
    ; If the selected file exists, open it and then close the search program
        IfExist, %LinkFilePath%
        {
            Run, %LinkFilePath%\Pricing\    ; Open the selected file
            CurrentGUI :=                           ; Clear the CurrentGUI variable to allow the search program to completely close
            Goto, GuiClose                          ; Close the search program
        }
    ; If the file does not exist, display an informational message
        IfNotExist, %LinkFilePath%
            Msgbox, The file no longer exists or has been moved.
Return

^s::     ; When user hits enter, open the selected file.
    GuiControlGet, FocusedControl, FocusV
    FocusedRow := LV_GetNext("", "Focused")
    ; If no file is selected, do nothing.
        if FocusedRow = 0
            return
    ; Get the name and path of the currently selected file
        LV_GetText(LinkFileName, FocusedRow, 1)
        LV_GetText(LinkFilePath, FocusedRow, 5)
    ; If the selected file exists, open it and then close the search program
        IfExist, %LinkFilePath%
        {
            Run, %LinkFilePath%\Submittals\    ; Open the selected file
            CurrentGUI :=                           ; Clear the CurrentGUI variable to allow the search program to completely close
            Goto, GuiClose                          ; Close the search program
        }
    ; If the file does not exist, display an informational message
        IfNotExist, %LinkFilePath%
            Msgbox, The file no longer exists or has been moved.
Return

;---------------------------------
;Use ctrl + c or Space to capture JN and Job Name into Clipboard
SetTitleMatchMode, 2 ; This let's any window that partially matches the given name get activated
#IfWinActive, ProjectSearch Results
^c::
Space::
GuiControlGet, FocusedControl, FocusV
        FocusedRow := LV_GetNext("", "Focused")
        ; If no file is selected, do nothing.
            if FocusedRow = 0
                return
        ; Get the name and path of the currently selected file
            LV_GetText(LinkFileName, FocusedRow, 1)
            ;MsgBox, %LinkFileName%
            SoundPlay *48
         Clipboard := LinkFileName
    if A_ThisHotKey = Space
        ExitApp
    Return
#IfWinActive
;-----------------------

MyListView:
    If A_GuiEvent = DoubleClick ; IF USER DOUBLE CLICKS A SEARCH RESULT, OPEN THE FILE
    {
        LV_GetText(LinkFileName, A_EventInfo, 1)
        LV_GetText(LinkFilePath, A_EventInfo, 5)
        StringLeft, jobno, LinkFileName, 7 ; this is how to handle longer job numbers when in path name.  it just grabs 7 digits and trims off the space and - (which is probably not needed)
        jobno := Trim(jobno, " -")
        Jobname := SubStr(LinkFileName,10,60)
        IfExist, %LinkFilePath%
            Run, %LinkFilePath%     ; opens the file
        IfNotExist, %LinkFilePath%
            Msgbox, The file no longer exists or has been moved.

        Gui, Test1:New
        Gui, Add, Listview, vTest1 x250 y45 w300 gOldList, Job Num|Name|FileName|FilePath
            Loop,25
            {
            FileReadLine, line, JobList.csv, %A_Index%
            if ErrorLevel
                break
            StringSplit, item, line, `,
            LV_Add("",item1,item2,item3,item4)
            }
        LV_Insert(1,,jobno,jobname,LinkFileName,LinkFilePath)
            Filename :=  "joblist.csv"
            FileDelete % FileName
            Gui Test1:ListView, test1       ;It's positioned in the desired Listview
            Loop % LV_GetCount() {
                LV_GetText(oCol1, A_index, 1)
                LV_GetText(oCol2, A_index, 2)
            ;---------10/28/2022
                LV_GetText(oCol3, A_index, 3)
                LV_GetText(oCol4, A_index, 4)

              oLine := oCol1 "," oCol2 "," oCol3 "," oCol4 "`n"
            ;--------10/28/2022
             FileAppend, % oLine, % FileName
               }

            Return
    }


; MAIN SEARCH RESULTS - IF USER RIGHT CLICKS A SEARCH RESULT, OPEN THE FILE LOCATION
    If A_GuiEvent = RightClick
        {
        GuiControlGet, FocusedControl, FocusV
        FocusedRow := LV_GetNext("", "Focused")

    ;   If no file is selected, do nothing.
        if FocusedRow = 0
            return

        LV_GetText(LinkFileName, FocusedRow, 1)
        LV_GetText(LinkFilePath, FocusedRow, 5)
        StringLeft, jobno, LinkFileName, 7 ; this is how to handle longer job numbers when in path name.  it just grabs 7 digits and trims off the space and - (which is probably not needed)
            jobno := Trim(jobno, " -")
            Jobname := SubStr(LinkFileName,10,60)
            ; If the file exists and has a path, open the file directory
            Gosub, RCSub
    }
    return

EndMacroSearchFiles:
    ExitApp

Push:
    GuiControl, , lvl, 1
Return

NewOrderBook:
Run chrome.exe "https://objm.svl.com/#/dashboard/order-book/?newWindow=true "   ; Go to New Order
return


StartSearch:
    ; Begin timing the search
        SearchStartTime := StartTimer()

    ; Remember the search parameters to use as a default if "refine search" is clicked
        LastSearchTerm := SearchTerm

    ; Save the search parameters to file to use as a default the next time the search program is run
        IfExist, %LAST_SEARCH_FILE%
            FileDelete, %LAST_SEARCH_FILE%
        FileAppend, %LastSearchTerm%, %LAST_SEARCH_FILE%

    ; Identify the current gui window
        CurrentGUI := "Search"

    ; REMOVE '*' FROM FILE EXTENSION SEARCHES (EX. '*.pdf' IS CONVERTED TO '.pdf')
        IfInString, SearchTerm, *.
            StringReplace, SearchTerm, SearchTerm, *., ., All

    ; SEPARATE SEARCH TERMS INTO INDIVIDUAL WORDS, KEEP QUOTED TEXT TOGETHER
        SearchTermCount :=
        Loop, Parse, SearchTerm, """", %A_Space%
        {
            If (A_Index / 2) <> ROUND( A_Index / 2, 0)
                Loop, Parse, A_Loopfield, %A_SPACE%, %A_Space%
                {
                    SearchTermCount++
                    SearchTerms%SearchTermCount% := A_Loopfield
                    If (SearchTerms%SearchTermCount% = "")
                        SearchTermCount--
                }
            Else
            {
                SearchTermCount++
                SearchTerms%SearchTermCount% := A_Loopfield
                If (SearchTerms%SearchTermCount% = "")
                    SearchTermCount--
            }
        }

    ; BUILD & DISPLAY SEARCH RESULTS GUI
        ;Gui, +AlwaysOnTop
        Gui, +Resize
        Gui, Add, Text,,SEARCHING FOR...%SearchTerm%%FilterText%    ; Name the GUI window (creates a header name).
        Gui, Font, S12  ; Set font for the GUI.
        Gui, Add, ListView, AltSubmit left r20 w600 h300 gMyListView, Name|Ext.|MB|Modified|Location ; Create columns and set formatting.
        LV_ModifyCol(2,0)
        LV_ModifyCol(3,0)
        LV_ModifyCol(5,0)
        LV_ModifyCol(4)
        Gui, Add, StatusBar,, Bar's starting text (omit to start off empty).    ; Create a status bar.

        Gui, Add, Text,,Press ESCAPE to cancel. ; Add text in the GUI.
        SB_SetText("Search in Progress")        ; Update GUI status bar
        gui, font, s16               ; Set the font for the buttons.
        Gui, Add, Button, gNewOrderBook, &New Order Book ; Create a "Refine Search" button.
        ;Gui, Add, Button, gPOSearch, &POSearch ; Create a "PO SEARCH" button.
        Gui, Add, Progress, vlvl  -Smooth 0x8 w250 h18
        SetTimer, Push, 45
        ;Gui, Add, Button, gGuiClose xp, &CLOSE        ; Create a "Close" button.
        Gui, Add, Button, Hidden Default h0, HiddenButtonOpenFile   ; Default action if enter is hit, to open selected file
        Gui, Show,,ProjectSearch Results   ; Show the newly created GUI objects.

    ; SEARCH INDEX FILE FOR MATCHES AND DISPLAY IN GUI
        ResultsLimited :=
        SearchInProgress = 1    ; search in progress until set to zero
        Loop, read, %INDEX_FILE%
        {
            if (SearchInProgress = 0)   ; If the search was cancelled, stop searching
                return

            LineNumber := A_Index
            Loop, parse, A_LoopReadLine, CSV
            {
                NotFoundA :=
                If A_Index = 2
                {
                    Loop, %SearchTermCount% ; Search for Unquoted Text
                    {
                        ThisSearchCheck :=
                        BreakOutExclude :=
                        ThisFileLocationNew := A_LoopField
                        ThisSearch := SearchTerms%A_INDEX%

                        ; Check if search string starts with "-" and if so use it to exclude files
                            StringLeft, ThisSearchCheck, ThisSearch, 1
                            If (ThisSearchCheck = "-")
                            {
                                StringTrimLeft, ThisSearch, ThisSearch, 1
                                BreakOutExclude :=
                                IfInString, A_LoopField, %ThisSearch%
                                {
                                    NotFoundA = 1
                                    Break
                                }
                            }
                            Else
                                IfNotInString, A_LoopField, %ThisSearch%
                                    NotFoundA = 1
                        }

                    If (NotFoundA <> 1) ; If a match and no excluded search terms are found..
                    {
                        NumberOfResults++
                        FoundMatch = 1
                        MatchedFileName := ThisFileLocationNew
                        LV_ModifyCol(1)
                        LV_ModifyCol(4)                        ; Auto-size each column to fit its contents.
                        LastMatchTime := StartTimer()   ; Timestamp of when match was found, used to limit how fast search result counts update on the gui
                    }
                    else
                        Break
                }
                If A_Index = 3
                    ; GET THE FULL FILE PATH INCLUDING NAME
                    If (FoundMatch = 1) ; If a match and no excluded search terms are found...
                    {
                        FoundMatch :=
                        StringSplit, FileDetails, A_LoopReadLine, `,    ; NEW 12/11/14 edit

                        Matched_Full_File_Path := A_LoopField   ; NEW 2014/12/11 edit
                        SplitPath, Matched_Full_File_Path, MatchedFilePath, MatchedFileDir, MatchedFileExtension, MatchedFileNameNoExt, MatchedFileDrive
                        StringUpper, MatchedFileExtension, MatchedFileExtension
                    }

                ; GET FILE MODIFIED DATE
                If A_Index = 1
                    MatchedFileModified := A_LoopField

                If (A_Index = 4)    ; GET FILE SIZE
                {
                    MatchedFileSize := A_LoopField
                    MatchedFileSize := ROUND((MatchedFileSize / 1024),2)
                    If (MatchedFileSize = 0)
                        MatchedFileSize := 0.01

                    ; Update GUI with search results
                    gui, font, s12  ; format next
                    FormatTime, MatchedFileModified, %MatchedFileModified%, yyyy-MM-dd  HH:mm   ; Format modified date for GUI
                    Gui, Font, S12
                    LV_Add("", MatchedFileName, MatchedFileExtension,MatchedFileSize,MatchedFileModified,Matched_Full_File_Path)    ; Update GUI with search result
                    LV_ModifyCol(1)  ; Auto-size each column to fit its contents.
                    LV_ModifyCol(4)

                    If (NumberOfResults = 1)
                    {
                        LV_Modify(0, "-Select")  ; to deselect all selected rows
                        LV_Modify(1,"Focus Select")
                                            }
                    MatchedFileName :=

                }
                If (NumberOfResults > ResultsLimit)
                    Break
            }
            If (NumberOfResults > ResultsLimit)
                Break
        }
        SearchInProgress = 0    ; identify the search as not currently in progress

        NumberOfResultsFormatted := ThousandsSep(NumberOfResults)
        SB_SetText(ProgressDots . " Matching Files: " . NumberOfResultsFormatted)

    ; Update GUI with number of results returned
        SB_SetText("Matching Files: " . NumberOfResultsFormatted)

    ; UPDATE GUI WITH NUMBER OF RESULTS RETURNED AND CONFIRM SEARCH COMPLETION
        IF (NumberOfResults > ResultsLimit)
            ResultsLimited := "RESULTS LIMITED TO " . ResultsLimit . " FILES"
        ELSE
            SearchResult := "SEARCH COMPLETE"

        ; Calculate the search time
            SearchTimeInSeconds := ROUND( EndTimer(SearchStartTime), 4)

        ; Hide progress bar when search is complete
            GuiControl, Hide, msctls_progress321        ; The ClassNN of the progress bar equals 'msctls_progress321'

        ; Gui, Font, norm s12 cDA4F49   ; format next
        Gui, Font, cDA4F49 s36  ; format next
        ResultsLimitedFormatted := ThousandsSep(ResultsLimited)
        SB_SetText("Matching Files: " . NumberOfResultsFormatted . A_SPACE . A_SPACE . A_SPACE . A_SPACE . A_SPACE . A_SPACE . A_SPACE . A_SPACE . A_SPACE . A_SPACE . ResultsLimitedFormatted . A_SPACE .A_SPACE .A_SPACE .A_SPACE .A_SPACE . A_SPACE .A_SPACE .A_SPACE .A_SPACE .A_SPACE . "Search completed in " . SearchTimeInSeconds . " seconds")
        ; SB_SetText("Search completed in " . SearchTimeInSeconds . " seconds")

    ; FORMAT GUI COLUMNS
        LV_ModifyCol(1)  ; Auto-size each column to fit its contents.
        LV_ModifyCol(3, "Float SortDesc")   ; For sorting purposes, indicate that column 3 (file size in MB) is an integer.
        LV_ModifyCol(4, "SortDesc")
        LV_ModifyCol(4)
   ; if (NumberOfResults = 1) ; IF ONLY ONE RESULT THEN GO AHEAD AND OPEN IT
     ;   Goto, ButtonHiddenButtonOpenFile


return
; END START SEARCH

RCSub:
If FileExist(LinkFileName)  ; If the file exists and has a path, open the file directory
{
    IniWrite,%LinkFileName%,%ComINIPath%,CurrentProject,FolderName
    IniWrite,%LinkFilePath%,%ComINIPath%,CurrentProject,FolderPath
    IniWrite,%LinkFilePath%\Orders,%ComINIPath%,CurrentProject,OrdersFolderPath
    IniWrite,%LinkFilePath%\Submittals,%ComINIPath%,CurrentProject,SubmittalsFolderPath
    IniWrite,%LinkFilePath%\Quotes,%ComINIPath%,CurrentProject,QuotesFolderPath
    IniWrite,%LinkFilePath%\Pricing,%ComINIPath%,CurrentProject,PricingFolderPath
    IniWrite,%jobno%,%ComINIPath%,CurrentProject,JobNumber
    IniWrite,%jobName%,%ComINIPath%,CurrentProject,JobName
    Run, RC.ahk
}
;   If the file does not exist or there is not path, display an informational message.
If !FileExist(LinkFileName) 
    Msgbox, The file no longer exists or has been moved.
Return
