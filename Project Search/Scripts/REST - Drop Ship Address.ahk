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

    oHTTP:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
    ;oHTTP.Option(6) := false
    If !IsObject(oHTTP)
    {
      MsgBox,4112,Fatal Error,Unable to create HTTP object
      ExitApp
    }
    ResponseFileDownload:="c:\temp\DocparserDownload.txt"
    FileDelete,%ResponseFileDownload%
    URL:="https://objm.svl.com/svl-services/rest/drop-ship?orderNo=" . JobNumber
    oHTTP.Open("GET",URL)
    oHTTP.SetRequestHeader("Authorization", AuthKey)
    OnError("LogError")
    oHTTP.Send()
    Response:=oHTTP.ResponseText
    FileAppend,%Response%,%ResponseFileDownload%

    FileRead json_String, c:\temp\DocparserDownload.txt
    parsed := JSON.Load(json_String)

    index := 1
        ; Create the ListView:
        Gui, Add, ListView, r20 w1000 gMyListView vMyListView altsubmit grid, Line|Address 1|Address 2|Address 3|City|State|Zip|Attention|Drop Ship Phone
        for index, value in parsed.result
            LV_Add("",value.lineNo,value.address1,value.address2,value.address3,value.city, value.state, value.zipCode, value.attention, value.dropPhone)
            LV_ModifyCol()  ; Auto-size each column to fit its contents.
        Gui, show,,Addresses - %FolderName%  - %customer%
        Gui, +Resize
        return

        MyListView:
        if (A_GuiEvent = "DoubleClick")
        {
            LV_GetText(lvAdd1, A_EventInfo, 2)  ; Get the text from the row's first field.
            LV_GetText(lvAdd2, A_EventInfo, 3)
            LV_GetText(lvAdd3, A_EventInfo, 4)
            LV_GetText(lvCity, A_EventInfo, 5)
            LV_GetText(lvState, A_EventInfo, 6)
            LV_GetText(lvZip, A_EventInfo, 7)
            LV_GetText(lvAttn, A_EventInfo, 8)
            LV_GetText(lvDropPhone, A_EventInfo, 9)
            IniWrite, % jobNumArray[3], %ComINIPath%, CurrentProject,JobName
            IniWrite, %lvAdd1%, %ComINIPath%, CurrentAddress,AField1
            IniWrite, %lvAdd2%, %ComINIPath%, CurrentAddress,AField2
            IniWrite, %lvAdd3%, %ComINIPath%, CurrentAddress,AField3
            IniWrite, %lvCity%, %ComINIPath%, CurrentAddress,CityField
            IniWrite, %lvState%, %ComINIPath%, CurrentAddress,StField
            IniWrite, %lvZip%, %ComINIPath%, CurrentAddress,ZipField
            IniWrite, %lvAttn%, %ComINIPath%, CurrentAddress,ContactField
            IniWrite, %lvDropPhone%, %ComINIPath%, CurrentAddress,PhoneField
            Run, Project Paster.ahk
            ExitApp
        }
        return

GUISize:
	width:=A_GuiWidth-15
	height:=A_GuiHeight-50
    buttonx := 20
    buttony := A_GuiHeight-30
	guicontrol, move, MyListView, w%width% h%height%
return

GUIClose:
GUIEscape:
ExitApp

LogError(exception) { ;Most errors are due to oAuth. This just autoruns the GET AUTH and reloads script.
    run, REST - GET AUTH.ahk
    Reload
    ;msgbox, % "Error on line " exception.Line ": " exception.Message "`n"
    ExitApp
}
return