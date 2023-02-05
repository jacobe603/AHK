#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


; SET SEARCH LOCATIONS (sub-folders included by default)
    Location1   := "S:\Projects"          ; Main Projects Folder
    Location2   := "R:\Projects\2017 thru 2020"   ; Archive Location 
    ;Location3   := "R:\Projects\2016"   ; Archive Location 
; add additional locations (e.g. Location3, Location4, etc. Looks like you can have as many of these as you want or maybe up to 9?)


;  SPECIFIC PROGRAMS/FILES TO INCLUDE AS FIRST RESULTS IF MATCHED
    CurrentDateTime := A_YYYY . A_MM . A_DD . A_Hour . A_Min . A_Sec    ; current timestamp, to show these files as the most current, and therefore at the top of search results    
    FileList .= CurrentDateTime . ",calc.exe,C:\Windows\system32\calc.exe,15`n"
    FileList .= CurrentDateTime . ",notepad.exe,C:\Windows\system32\notepad.exe,15`n"
    FileList .= CurrentDateTime . ",cmd.exe,C:\Windows\system32\cmd.exe,15`n"
    FileList .= CurrentDateTime . ",Microsoft Outlook 2010,C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Outlook 2010.lnk,1`n"
    FileList .= CurrentDateTime . ",mspaint.exe,C:\Windows\system32\mspaint.exe,15`n"


; INDEX FILE LOCATION IS IN THE CURRENT USER'S MY DOCUMENTS FOLDER
    INDEX_FILE := A_MyDocuments . "\FileIndex.csv"
    INDEX_FILE2 := A_MyDocuments . "\FileIndexFULL.csv"

; INCREASE MAXIMUM MEMORY ALLOWED FOR EACH VARIABLE - NECESSARY FOR THE INDEXING VARIABLE
    #MaxMem 4095

; COUNT THE NUMBER OF SEARCH LOCATIONS
    n :=
    Loop
    {
        n++
        Location := Location%n%
        If (Location%n% = "")
        {
            n--
            Break
        }
    }

; CREATE GUI
    Gui, Add, Text, w1250, Indexing Files...
    Gui, Add, StatusBar, +wrap, Bar's starting text
    SB_SetParts(65)
    SB_SetText(A_LoopFileFullPath)
    Gui, Show

; INDEX FILES AND UPDATE GUI WITH STATUS
    Loop, %n%
    {
        ThisFolder := Location%A_Index%
        Loop, %ThisFolder%\*.*, 2, 0
        {
            ; EXCLUDE FILES FROM THE SEARCH DATABASE (modify as necessary)
            if A_LoopFileName not contains $,~,.tmp,.ddf,.bak,.gg,.ini,.DS_Store,.db,.url,.ahk.bak,.upd,.ctx,.dll,.indd,.SESSION,.PROPERTIES,.TEXTCLIPPING,.eps,.idml,.zbin,.swf,.strings,.rsrc,.psb,.lst,.lrtemplate,framework,.aapp,.aaui,.ai,.asfx,.bin,.eve,.exv,.nav,.nib,.pak,.plist,.sequ,.abr,.1,.acb,.acv,.afm,.ase,.atn,.cer,.cnv,.CR2,.crv,.csh,.css,.dat,.der,.dic,.emf,.eml,.fla,.flt,.FON,.h,.hqx,.htm,.icc,.inx,.ipp,.js,.kfp,.kfr,.ln,.lnk,.log,.lwo,.mac,.MMM,.NEF,.PFB,.PFM,.ps,.rc,.so,.spp,.svg,.ttc,.vc,.vcf,.wmf,.xdc,.ico,.ilex,.api,.bat,.BML,.BMS,.BUP,.cab,.CAC,.CD,.cfg,.conf,.cptx,.DFN,.dif,.dir,.glox,.icns,.ID,.idms,.IFO,.indb,.indl,.indt,.inf,.io,.lua,.ocr,.ocx,.otf,.p12,.pat,.pct,.pds,.pem,.pfx,.php,.rrt,.rwz,.scpt,.sdef,.sitx,.thmx,.ttf,.VOB,.vsd,.vss,.xmp,.xnk,.xsd,.xsl,.img,.xml,.psb,.psd,.eps


            {
                ; Skip any file that is either H (Hidden), R (Read-only), or S (System). Note: No spaces in "H,R,S".
                    if A_LoopFileAttrib contains H,R,S
                        continue

                ; Skip any file without a file extension
                    ;If (A_LoopFileExt = "")
                       ; continue


                ; If the file name contains a comma, put quotation marks around the field to prepare it for insertion into a CSV file
                    ModA_LoopFileName := A_LoopFileName
                    IfInString, A_LoopFileName, `,
                        ModA_LoopFileName := """" . A_LoopFileName . """"

                ; If the file path contains a comma, put quotation marks around the field to prepare it for insertion into a CSV file
                    ModA_LoopFileFullPath := A_LoopFileFullPath
                    IfInString, A_LoopFileFullPath, `,
                        ModA_LoopFileFullPath := """" . A_LoopFileFullPath . """"

                ; Save details of file
                    FileList .= A_LoopFileTimeModified . "," . ModA_LoopFileName . "," . ModA_LoopFileFullPath . "," . A_LoopFileSizeKB . "`n"
                    FileList2 .= A_LoopFileTimeModified . "," . ModA_LoopFileName . "," . ModA_LoopFileFullPath . "," . A_LoopFileSizeKB . "`n"

                ; GUI Status Update
                    FileCount++
                    If ((FileCount / 13) = ROUND(FileCount / 13, 0))
                    {
                        SB_SetText(FileCount . " files", 1)
                            SB_SetText(A_LoopFileFullPath, 2)
                        Gui, Show, NA
                    }
            }
            ; EXCLUDED FILES:  Save to a second file as a backup source
            else
            {
                ; If the file NAME contains a comma, put quotation marks around the field to prepare it for insertion into a CSV file
                    ModA_LoopFileName := A_LoopFileName
                    IfInString, A_LoopFileName, `,
                        ModA_LoopFileName := """" . A_LoopFileName . """"

                ; If the file PATH contains a comma, put quotation marks around the field to prepare it for insertion into a CSV file
                    ModA_LoopFileFullPath := A_LoopFileFullPath
                    IfInString, A_LoopFileFullPath, `,
                        ModA_LoopFileFullPath := """" . A_LoopFileFullPath . """"

                ; Save details of file
                    FileList2 .= A_LoopFileTimeModified . "," . ModA_LoopFileName . "," . ModA_LoopFileFullPath . "," . A_LoopFileSizeKB . "`n"
            }               
        }
    }

; GUI Status Update
    SB_SetText("SORTING RESULTS...", 1)
        SB_SetText("", 2)
    Gui, Show, NA


; SORT the file list from newest to oldest based on the modified date, which is now the first CSV column
    ; Sort, FileList , N R
    Sort, FileList , R

; GUI Status Update
    SB_SetText("RE-SORTING RESULTS...", 1)
        SB_SetText("", 2)
    Gui, Show, NA


; GUI Status Update
    SB_SetText("DELETING INDEX FILE", 1)
        SB_SetText("", 2)
    Gui, Show, NA


; DELETE INDEX FILE IF IT ALREADY EXISTS
    IfExist, %INDEX_FILE%
        FileDelete, %INDEX_FILE%
    IfExist, %INDEX_FILE2%
        FileDelete, %INDEX_FILE2%
    Sleep, 2500


; GUI Status Update
    SB_SetText("CREATING PRIMARY INDEX FILE", 1)
        SB_SetText("", 2)
    Gui, Show, NA


; CREATE PRIMARY INDEX FILE
    FileAppend, %FileList%, %INDEX_FILE%


; GUI Status Update
    SB_SetText("CREATING SECONDARY INDEX FILE (NO EXCLUDED FILES)", 1)
        SB_SetText("", 2)
    Gui, Show, NA


; CREATE SECONDARY INDEX FILE
    FileAppend, %FileList2%, %INDEX_FILE2%


; GUI Status Update
    SB_SetText("COMPLETE", 1)
        SB_SetText("", 2)
    Gui, Show, NA


; END RE-INDEX
    GUI Destroy
    ExitApp