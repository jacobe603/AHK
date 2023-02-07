#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance force
SetTitleMatchMode 2
SetWorkingDir %A_ScriptDir%
#Include Library\JSON.AHK
;#Include Library\CreateFormData.AHK

    ComINIPath=..\Data\Search.ini
    IniRead,AuthKey,%ComINIPath%,Auth,AuthKey
    ;IniRead,JobNumber,%ComINIPath%,CurrentProject,JobNumber
    ;IniRead,FolderName,%ComINIPath%,CurrentProject,FolderName
    statusArray := {"Bid": 3,"Chase - T": 12, "Awaiting Award": 4,"Chase": 5,"Sold": 6,"Completed": 7,"No Bid": 10,"Order Book": 16,"Sold - Released": 17}

    ;---------------Initial GUI-------
    Gui, Font, s11, Calibri  ; Set 10-point Verdana.

    Gui, Add, ComboBox, vStatus x90 y45 w124 h35 R10 sort hwndhMfgNames gStatusCode, Bid|Chase - T|Awaiting Award|Chase|Sold|Complete|No Bid|Order Book|Sold - Released

    ;Gui, Add, Edit, vstatusCode x225 y45 w50
    Gui, Add, Edit, vJobName x90 y93 w326
    Gui, Add, Edit, vCity x90 y141 w163
    Gui, Add, Edit, vState x90 y189 w163
    Gui, Add, Edit, vEngineer x90 y237 w326
    Gui, Add, Edit, vSalesCode x90 y285 w150

    ;------------------------------------- Text Labels
    Gui, Add, Text, x20 y45 w57, Status
    Gui, Add, Text, x20 y93 w57, Job Name
    Gui, Add, Text, x20 y141 w57, City
    Gui, Add, Text, x20 y189 w57, State
    Gui, Add, Text, x20 y237 w57, Engineer
    Gui, Add, Text, x20 y285 w57, Sales Code
    ;------------------------------------- Buttons

    Gui, Add, Button, gButton1 x90 y335 w144, Search
    Gui, Add, Button, gCancel x350 y335 w144, Cancel
    Gui, Show, w550 h400, Search Jobs

    return

    Button1:
    gui, Submit
    StatusCode := statusArray[(status)]
    gui, Destroy

    ;oHTTP:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
    oHTTP:=ComObjCreate("Msxml2.XMLHTTP.6.0")
    ;oHTTP.Option(6) := false
    If !IsObject(oHTTP)
    {
      MsgBox,4112,Fatal Error,Unable to create HTTP object
      ExitApp
    }
    ResponseFileDownload:="c:\temp\DocparserDownload.txt"
    FileDelete,%ResponseFileDownload%
    URL:="https://objm.svl.com/svl-services/rest/job/search"
    data = {"statusId":"%statusCode%","name":"%jobName%","city":"%city%","state":"%state%","engineerName":"%engineer%"}
    ;Payload := JSON.Load(PayloadJSON)
    oHTTP.Open("POST",URL, false)
    oHTTP.SetRequestHeader("Authorization", AuthKey)
    oHTTP.SetRequestHeader("Content-Type", "application/json")
    oHTTP.Send(data)

    status := oHTTP.status
    Response := oHTTP.ResponseText
    if !(status = 200 || status = 201)
        MsgBox, Bad HTTP status: %status%`nResponse text:`n%Response%
    FileAppend,%Response%,%ResponseFileDownload%
    FileRead json_String, c:\temp\DocparserDownload.txt
    parsed := JSON.Load(json_String)


    ;data := parsed.result[1].equipDesc
    index := 1
        ; Create the ListView:
        Gui, Add, ListView, r20 w1000 gMyListView altsubmit grid, Job Num|Status|Name|City|State|Eng|Location|Sales Team|Quote Amount
        for index, value in parsed.result
            if InStr(value.steam,salesCode)
        {
            LV_Add("",value.jobNo,value.statusName,value.name,value.city,value.state,value.engineer,value.location,value.sTeam,value.quoteAmnt)
        }

        LV_ModifyCol()  ; Auto-size each column to fit its contents.

        ; Display the window and return. The script will be notified whenever the user double clicks a row.
        Gui, Add, Button, Hidden Default, OK
        ;LV_ModifyCol(6, "Integer")
        ;LV_ModifyCol(1, "Integer")

        Gui, show,,Job Search - %data%
        Gui, +Resize

        return

MyListView:
if (A_GuiEvent = "DoubleClick")
{
    LV_GetText(jobNum, A_EventInfo, 1)  ; Get the text from the row's first field.
    GoToJobFolder(jobNum)
}
return


GoToJobFolder(match)     ; Function that sorts through folder names on S: drive to find the match and then opens the folder.
{
    count := ""
    Loop, Files, S:\Projects\%match%*, D
        {
            count .= 1
            jobpath := A_LoopFileFullPath
            Clipboard := A_LoopFileName
            Break
        }
            if !count {
            msgbox,0, File Error, Job NOT Found, 2
        }
        Run, %jobpath%
    Return
}


LogError(exception) { ;Most errors are due to oAuth. This just autoruns the GET AUTH and reloads script.
    run, REST - GET AUTH.ahk
    Reload
    ;msgbox, % "Error on line " exception.Line ": " exception.Message "`n"
    ExitApp
}
return


GuiClose:  ; Indicate that the script should exit automatically when the window is closed.
GuiEscape:
Gui, Destroy
ExitApp
;~ ExitApp



StatusCode:
return

