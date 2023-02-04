    #Include Library\JSON.AHK

    TreeRoot := "S:\Projects\942610 - Spaulding Stone Dehumidification"
    PlansFold := TreeRoot "\Plans"
    QuoteFold := TreeRoot "\Quotes"
    PricingFold := TreeRoot "\Pricing"
    SpecFold := TreeRoot "\Spec"  ;Updated to just root folder
    SubmitFold := TreeRoot "\Submittals"
    OrderFold := TreeRoot "\Orders"
    FileRead json_String, c:\temp\DocparserDownload.txt
    parsed := JSON.Load(json_String)
    customer := parsed1.custName
    totalVol := 0
    index := 1
    ; Allow the user to maximize or drag-resize the window:
    ;Gui +Resize

    gui, new
    gui, font,s10, Seqoe UI

    Gui,Add,Tab3, BackgroundTrans,Quote|Sold|Released


    gui, font, s12
    Gui, Tab, Quote
    Gui, Add, Text, section w400 Center, Root
    Gui Add, ActiveX, border y+5 w400 h200 vWB1, Shell.Explorer
    WB1.Navigate(TreeRoot)
    Gui, Add, Text, y+10 w400 Center, Quote
    Gui Add, ActiveX, border y+5 w400 h200 vWB2, Shell.Explorer
    WB2.Navigate(QuoteFold)
    Gui, Add, Text, x+10 ys w400 Center, Plans
    Gui Add, ActiveX, border y+5 w400 h200 vWB3, Shell.Explorer
    WB3.Navigate(PlansFold)
    Gui, Add, Text, y+10 w400 Center, Pricing
    Gui Add, ActiveX, border y+5 w400 h200 vWB4, Shell.Explorer
    WB4.Navigate(PricingFold)

    ;---------------------
    ; Create the ListView:
    gui, font, s8
    Gui, Add, ListView, y+10 xs r15 w810 grid, Line No|Mfg Code|Mfg|Item|Qty|Volume|Cost|Customer PO|Rep PO|Date Added|Ship Date
    for index, value in parsed.result
    {
        totalVol += value.lnVolume
        LV_Add("",value.lineNo,value.mfgCode,value.mfgName,value.equipdesc,value.lnQty, "$" . value.lnVolume,"$" . value.lnCost, value.custPO, value.       repTrackingNo, SubStr(value.dateAdded,1,10), SubStr(value.curShipDate,1,10))
    }
    LV_ModifyCol()  ; Auto-size each column to fit its contents.

;------------------------Back to the other tabs
    gui, font, s12
    Gui, Tab, Sold
    Gui, Add, Text, section w400 Center, Root
    Gui Add, ActiveX, border y+5 w400 h200 vWB5, Shell.Explorer
    WB5.Navigate(TreeRoot)
    Gui, Add, Text, y+10 w400 Center, Quote
    Gui Add, ActiveX, border y+5 w400 h200 vWB6, Shell.Explorer
    WB6.Navigate(QuoteFold)
    Gui, Add, Text, x+10 ys w400 Center, Pricing
    Gui Add, ActiveX, border y+5 w400 h200 vWB7, Shell.Explorer
    WB7.Navigate(PricingFold)
    Gui, Add, Text, y+10 w400 Center, Orders
    Gui Add, ActiveX, border y+5 w400 h200 vWB8, Shell.Explorer
    WB8.Navigate(OrderFold)

    Gui, Tab, Released
    Gui, Add, Text, section w400 Center, Root
    Gui Add, ActiveX, border y+5 w400 h200 vWB9, Shell.Explorer
    WB9.Navigate(TreeRoot)
    Gui, Add, Text, y+10 w400 Center, Quote
    Gui Add, ActiveX, border y+5 w400 h200 vWB10, Shell.Explorer
    WB10.Navigate(QuoteFold)
    Gui, Add, Text, x+10 ys w400 Center, Pricing
    Gui Add, ActiveX, border y+5 w400 h200 vWB11, Shell.Explorer
    WB11.Navigate(PricingFold)
    Gui, Add, Text, y+10 w400 Center, Submittals
    Gui Add, ActiveX, border y+5 w400 h200 vWB12, Shell.Explorer
    WB12.Navigate(SubmitFold)
    Gui, Tab

    Gui, Add, Button, gFilePath, Copy Path

    ;Gui, Add, Button, x- w400 H40 gSHOW, SHOW
    Gui, Show,autosize, Project Explorer - %TreeRoot%
    Return


    SetTitleMatchMode, 1
    #If WinActive("Project Explorer - "TreeRoot)
    XButton1::
    send, {LButton}

    ControlGetFocus, controlname, Project Explorer - %TreeRoot%
    if (controlname = "SysListView321")
        WB1.GoBack()
    if (controlname = "SysListView322")
        WB2.GoBack()
    if (controlname = "SysListView323")
        WB3.GoBack()
    if (controlname = "SysListView324")
        WB4.GoBack()
    if (controlname = "SysListView325")
        WB5.GoBack()
    if (controlname = "SysListView326")
        WB6.GoBack()
    if (controlname = "SysListView327")
        WB7.GoBack()
    if (controlname = "SysListView328")
        WB8.GoBack()
    if (controlname = "SysListView329")
        WB9.GoBack()
    if (controlname = "SysListView3210")
        WB10.GoBack()
    if (controlname = "SysListView3211")
        WB11.GoBack()
    if (controlname = "SysListView3212")
        WB12.GoBack()
        return


    FilePath:
    clipboard = %TreeRoot%
    return

    ;GuiClose:
    ;GuiEscape:      ; If ESCAPE is pressed
    ;GUI, Destroy
    ;return