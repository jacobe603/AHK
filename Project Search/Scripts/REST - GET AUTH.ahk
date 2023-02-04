#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; Working in scrpit
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance force
SetTitleMatchMode 2
SetWorkingDir %A_ScriptDir%

ComINIPath = ..\Data\Search.ini

menu, Tray, Icon, icons\svl.png
IniRead,AuthKey,%ComINIPath%,Auth,AuthKey
IniRead,SalesCode,%ComINIPath%,Sales,SalesCode

    oHTTP:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
    ;oHTTP.Option(6) := false
    If !IsObject(oHTTP)
    {
        MsgBox,4112,Fatal Error,Unable to create HTTP object
        ExitApp
    }
    ResponseFileDownload:="c:\temp\DocparserDownload.txt"
    FileDelete,%ResponseFileDownload%
    URL:="https://objm.svl.com/svl-services/rest/login"
    JSONString = ["jacobe","647 Trunk Theft Phrase"]
    oHTTP.Open("PUT",URL)
    ;oHTTP.SetRequestHeader("Authorization", AuthKey)
    oHTTP.SetRequestHeader("Content-Type", "application/json")
    ;oHTTP.SetRequestHeader("Cookie", "_ga=GA1.2.91386688.1639424321; _hjSessionUser_1366802=eyJpZCI6ImY5ZTViNDJhLTRlYWYtNWI2NS1hMjZlLThmNGU4ZTNkZjQ5ZCIsImNyZWF0ZWQiOjE2Mzk0MjQzMjExOTIsImV4aXN0aW5nIjp0cnVlfQ==")
    ;SetCookies(oHTTP, cookies) ;The first parameter is the variable to which we intialized the WinHttpRequest object, the second parameter is the variable from which we will read our cookies.
    ;OnError("LogError")
    oHTTP.Send(JSONString)
    Response:=oHTTP.ResponseText
    FileAppend,%Response%,%ResponseFileDownload%
    FileRead json_String, c:\temp\DocparserDownload.txt
    ;parsed := JSON.Load(json_String)
    auth = Bearer %json_String%
    msgbox, % auth
    IniWrite,%auth%,%ComINIPath%,Auth,AuthKey

ExitApp

