#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; Working In Scrpit Y | Out of Script Y
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance force
SetTitleMatchMode 2
SetWorkingDir %A_ScriptDir%

#Include Library\JSON.AHK
ComINIPath=..\Data\Search.ini

	IniRead,FolderName,%ComINIPath%,CurrentProject,FolderName
    IniRead,FolderPath,%ComINIPath%,CurrentProject,FolderPath
    IniRead,SubmittalsFolderPath,%ComINIPath%,CurrentProject,SubmittalsFolderPath
    IniRead,QuotesFolderPath,%ComINIPath%,CurrentProject,QuotesFolderPath
    IniRead,PricingFolderPath,%ComINIPath%,CurrentProject,PricingFolderPath
    IniRead,JobNumber,%ComINIPath%,CurrentProject,JobNumber
    IniRead,AuthKey,%ComINIPath%,Auth,AuthKey

    jobNumArray := StrSplit(FolderName, A_Space,,3)


    ;oHTTP:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
    oHTTP:=ComObjCreate("Msxml2.XMLHTTP.6.0")
    ;oHTTP.Option(6) := false
    If !IsObject(oHTTP)
    {
      MsgBox,4112,Fatal Error,Unable to create HTTP object
      ExitApp
    }
    ResponseFileDownload:="c:\temp\DocparserDownload1.txt"
    FileDelete,%ResponseFileDownload%
    URL:="https://objm.svl.com/svl-services/rest/job/search"
    data = {"jobNo":"%jobNumber%","name":"%jobName%","city":"%city%","state":"%state%","engineerName":"%engineer%"}
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
    FileRead json_String, c:\temp\DocparserDownload1.txt
    parsed := JSON.Load(json_String)

    IniWrite, % value.jobID, %ComINIPath%, CurrentProject,JobID






ExitApp

