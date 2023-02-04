#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance force
SetTitleMatchMode 2
SetWorkingDir %A_ScriptDir%

#Include Library\JSON.AHK

;SetTimer, ForceExitApp,  10000 ; 10 seconds
;Ini File for RC Menu Variable Transfer
    ComINIPath = ..\Data\Search.ini
    IniRead,FolderName,%ComINIPath%,CurrentProject,FolderName
    IniRead,FolderPath,%ComINIPath%,CurrentProject,FolderPath
    IniRead,SubmittalsFolderPath,%ComINIPath%,CurrentProject,SubmittalsFolderPath
    IniRead,QuotesFolderPath,%ComINIPath%,CurrentProject,QuotesFolderPath
    IniRead,PricingFolderPath,%ComINIPath%,CurrentProject,PricingFolderPath
    IniRead,JobNumber,%ComINIPath%,CurrentProject,JobNumber
    IniRead,AuthKey,%ComINIPath%,Auth,AuthKey

html1 =
(
<html></head><body lang=EN-US link="#0563C1" vlink="#954F72"><div class=WordSection1><p class=MsoNormal>This project&#8217;s quote and recap are ready for Outside Sales Review.&nbsp; When review is complete, <u>please send to Projects</u> and have them send this quote out.<o:p></p><p class=MsoNormal>Let me know if you have any questions.<o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><a href="file://SVLFS2.ad.svl.com/Sales-2$/Projects/%FolderName%">%FolderPath%</a><o:p></o:p></p><p class=MsoNormal><a href="https://objm.svl.com/#/dashboard/job-detail/job-no/%JobNumber%?newWindow=true" target="_blank"><span style='font-size:12.0pt;background:white'>Job Manager - %FolderName%</span></a><span style='font-size:12.0pt;color:black'><o:p></o:p></span></p>
)

html2 =
(
<html></head><body lang=EN-US link="#0563C1" vlink="#954F72"><div class=WordSection1><p class=MsoNormal>This project&#8217;s quote and recap have been reviewed and the project is ready to be sent.&nbsp;<u>Please send to contacts on Bidder&#8217;s List</u>.<o:p></p><p class=MsoNormal>Let me know if you have any questions.<o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><a href="file://SVLFS2.ad.svl.com/Sales-2$/Projects/%FolderName%">%FolderPath%</a><o:p></o:p></p><p class=MsoNormal><a href="https://objm.svl.com/#/dashboard/job-detail/job-no/%JobNumber%?newWindow=true" target="_blank"><span style='font-size:12.0pt;background:white'>Job Manager - %FolderName%</span></a><span style='font-size:12.0pt;color:black'><o:p></o:p></span></p>
)

; OBJM Submenu Items
Menu, MyMenu, Add, Order Book, MenuHandlerA1
Menu, MyMenu, Icon, Order Book, Icons\OrderBook.ico
Menu, MyMenu, Add, Job Manager, MenuHandlerA2
Menu, MyMenu, Icon, Job Manager, Icons\Svllogo.ico
Menu, MyMenu, Add, New Order, MenuHandlerA3
Menu, MyMenu, Add

;Lookup Items
Menu, MyMenu, Add, Order Details, MenuHandlerA4
Menu, MyMenu, Icon, Order Details, Icons\OrderDetails.ico
Menu, MyMenu, Add, Address, MenuHandlerE1
Menu, MyMenu, Icon, Address, Icons\Address.ico
Menu, MyMenu, Add, Project Status, MenuHandlerE3
Menu, MyMenu, Icon, Project Status, Icons\clock.ico
Menu, MyMenu, Add

; Email Submenu Items
Menu, MyMenu, Add, Review Project, MenuHandlerB1
Menu, MyMenu, Icon, Review Project, Moricons.dll , 65
Menu, MyMenu, Add, Send Project, MenuHandlerB2
Menu, MyMenu, Add

;Folders
Menu, MyMenu, Add, Quote, MenuHandlerD1
Menu, MyMenu, Icon, Quote, Icons\Quote.ico
Menu, MyMenu, Add, Recap/Pricing, MenuHandlerD2
Menu, MyMenu, Icon, Recap/Pricing, Icons\Recap.ico
Menu, MyMenu, Add, Submittals, MenuHandlerD3
Menu, MyMenu, Icon, Submittals,Icons\Submittals.ico
Menu, MyMenu, Add, Orders, MenuHandlerD4
Menu, MyMenu, Icon, Orders,Icons\Submittals.ico

;PO Tracking
Menu, MyMenu, Add
Menu, MyMenu, Add, Ruskin Order, MenuHandlerC2
; --------------------------Main Menu - Subs---------------------------------

; Quick Access Items
Menu, MyMenu, Add  ; Add a separator line below the submenu.
Menu, MyMenu, Add, Copy File Path, MenuHandlerE2
Menu, MyMenu, Icon, Copy File Path, Icons\FilePath.ico


Menu, MyMenu, Add  ; Add a separator line below the submenu.
Menu, MyMenu, Add, Update Auth Key , MenuHandlerF1

Menu, MyMenu, Show

return

MenuHandler:
;MsgBox Is this your project number %JobNumber%
return


MenuHandlerA1:
    Run chrome.exe "https://objm.svl.com/#/dashboard/order-inquiry-no-dollars/%JobNumber%?newWindow=true" " --new-window "   ; ADDED - Open new Chrome window and go to Order Book
    return
MenuHandlerA2:
    Run chrome.exe "https://objm.svl.com/#/dashboard/job-detail/job-no/%JobNumber%?newWindow=true "   ; ADDED - Open new Chrome window and go to Order Book
return

MenuHandlerA3:
    Run chrome.exe "https://objm.svl.com/#/dashboard/order-book/?newWindow=true "   ; Go to New Order
    send, {tab 9}
    clipboard := JobNumber
return

MenuHandlerA4:
    run, "REST - Order Query.ahk"
return

GuiClose:  ; Indicate that the script should exit automatically when the window is closed.
    ExitApp

MenuHandlerB1:
    if WinExist(JobNumber)
    {
        MsgBox, You have open files for this job
        WinMinimizeAll
        WinGet, Array, List, %JobNumber%
        Loop, %Array%
        {
            WinActivate, % "ahk_id " Array%A_Index%
        }
        Return
    }
    else

    Outlook := ComObjActive("Outlook.Application")
    email := Outlook.Application.CreateItem(0)
    email.To := "bene@svl.com; leef@svl.com"
    email.HTMLBody := html1
    email.Subject := "REVIEW QUOTE - " FolderName
    ;email.Attachments.Add("some full path", 1, 1, "some full path")
    email.Display(true)
    ; email.Send() ;could lose the display and uncomment this
    return

MenuHandlerB2:
    if WinExist(JobNumber)
    {
        MsgBox, You have open files for this job
        WinMinimizeAll
        WinGet, Array, List, %JobNumber%
        Loop, %Array%
        {
            WinActivate, % "ahk_id " Array%A_Index%
        }
        Return
    }
    else
        Outlook := ComObjActive("Outlook.Application")
        email := Outlook.Application.CreateItem(0)
        email.To := "Projects@svl.com"
        email.HTMLBody := html2
        email.Subject := "SEND QUOTE - " FolderName
        ;email.Attachments.Add("some full path", 1, 1, "some full path")
        email.Display(true)
        ; email.Send() ;could loose the display and uncomment this
    return


MenuHandlerC1:
    POSearchTerm := JobNumber
    ;#Include U:\AHK\PO Search\Cook PO Search.ahk
    ExitApp

MenuHandlerC2:
    Run chrome.exe "https://www.ruskin.com/VIPNet/orders?t=5&q=%JobNumber%&custNum=33881 "   ; Ruskin VIP Net PO Search
    ;POSearchTerm := JobNumber
    ;#Include U:\AHK\PO Search\Ruskin PO Search.ahk
    ExitApp

MenuHandlerC3:
    POSearchTerm := JobNumber
    ;#Include U:\AHK\PO Search\Titus PO Search.ahk
    ExitApp

MenuHandlerC4:
    POSearchTerm := JobNumber
    ;#Include U:\AHK\PO Search\Aaon PO Search.ahk
    ExitApp

MenuHandlerC5:
    POSearchTerm := JobNumber
    ;#Include U:\AHK\PO Search\Daikin PO Search.ahk
    ExitApp

MenuHandlerD1:
    run, %QuotesFolderPath%
    ExitApp

MenuHandlerD2:
    run, %PricingFolderPath%
    ExitApp

MenuHandlerD3:
    run, %SubmittalsFolderPath%
    ExitApp

MenuHandlerD4:
    run, %SubmittalsFolderPath%
    ExitApp

MenuHandlerE1:
        run, REST - Drop Ship Address.ahk
        ;run, Project Paster_v3.ahk
    ExitApp

MenuHandlerE2:
    clipboard := FolderPath
    ExitApp


MenuHandlerE3:
        run, ProjTrackLookUp.ahk
    ExitApp

MenuHandlerF1:
    Global ComINIPath
    InputBox, authkeynew, Authorization Key
    winset, AlwaysOnTop, On, Authorization Key
    IniWrite,"%authkeynew%",%ComINIPath%,Auth,AuthKey
    ExitApp


;SetTitleMatchMode, 2 ; This lets any window that partially matches the given name get activated
#IfWinActive ahk_class AutoHotkeyGUI
^c::
    GuiControlGet, FocusedControl, FocusV
    FocusedRow := LV_GetNext("", "Focused")
    ; If no file is selected, do nothing.
        if FocusedRow = 0
            return

    ; Get the name and path of the currently selected file
        LV_GetText(OrderLine1, FocusedRow, 1)
        LV_GetText(OrderLine2, FocusedRow, 2)
        LV_GetText(OrderLine3, FocusedRow, 3)
        LV_GetText(OrderLine4, FocusedRow, 4)
        LV_GetText(OrderLine5, FocusedRow, 5)
        LV_GetText(OrderLine6, FocusedRow, 6)
        LV_GetText(OrderLine7, FocusedRow, 7)
        SoundPlay *48
     Clipboard := OrderLine1 " " OrderLine2 " " OrderLine3 " " OrderLine4 " " OrderLine5 " " OrderLine6 " " OrderLine7
Return


ForceExitApp:
if WinActive("ahk_exe AutoHotkey.exe")
{
    SetTimer, ForceExitApp,  10000
    return
}
SetTimer,  ForceExitApp, Off
ExitApp
