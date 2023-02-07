#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force
SetTitleMatchMode 2

Menu, Tray, Icon, Icons\SVL.png

ComINIPath = ..\Data\Search.ini

IniRead,SalesCode,%ComINIPath%,Sales,SalesCode
IniRead,jobNum,%ComINIPath%,CurrentProject,JobNumber

myConn := ComObjCreate("ADODB.Connection")
path := "X:\Jacob\Project Tracker v0.accdb"
myConn.Open("Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" path ";")
query = SELECT * FROM mainProjTrack WHERE (mainProjTrack.jobNo = '%jobNum%');

OnError("LogError")

rs := myConn.Execute(query)
customer := rs.Fields("custName").value
FolderName := rs.Fields("jobNo").value . " - " rs.Fields("jobName").value

Gui, Add, ListView, r20 w1000 gMyListView vMyListView grid, Cust PO|Mfg|Equip Desc|Qty|Vol|SO #|Ship Date|Status|Notes|Mfg Code|Desc|Cost
;LV_ModifyCol(5,"Integer")

while ! rs.EOF
{
    LV_Add("",rs.Fields("custPONo").value, rs.Fields("mfgName").value, rs.Fields("equipDesc").value, format("{:d}", rs.Fields("equipQty").value), "$" . format("{:.2f}",rs.Fields("lnVolume").value),rs.Fields("mfgTrackingNo").value, rs.Fields("shipDate").value, rs.Fields("statusName").value, rs.Fields("notes").value, rs.("mfgCode").value, rs.("fullDesc").value, rs.("lnCost").value)

    rs.MoveNext()

        ; Display the window and return. The script will be notified whenever the user double clicks a row.
}

LV_ModifyCol()  ; Auto-size each column to fit its contents.
Gui, Add, Button, gButtonPT w115, Open Project Tracker
Gui, Add, Button, Hidden Default, OK
Gui, +Resize
Gui, +AlwaysOnTop
Gui, show,,Project Detail - %FolderName%  - %customer%

rs.Close(), rs := ""
conn.Close(), conn := ""

Return

ButtonPT:
msgbox OK
return

GuiClose:
GuiEscape:
{
rs.Close(), rs := ""
conn.Close(), conn := ""

ExitApp
}

GUISize:
	width:=A_GuiWidth-15
	height:=A_GuiHeight-50
    buttonx := 20
    buttony := A_GuiHeight-30
	guicontrol, move, MyListView, w%width% h%height%
    guicontrol, move, Open Project Tracker, x%buttonx% y%buttony%
return


        MyListView:
        if (A_GuiEvent = "DoubleClick")
        {
            LV_GetText(lvSO, A_EventInfo, 6)  ; Get the text from the row's first field.
            clipboard = %lvSO%
            ;ExitApp
        }
        return

        
LogError(exception) {
if WinExist("Project Tracker JE")
    WinActivate
MsgBox, Database or Excel file may be open.  Close and try again.
    ExitApp
}
return

ButtonOK:
            Controlget, SelectedItems, List,Selected, SysListView321, Project Detail
            X =
            D =
            Loop, Parse, SelectedItems, `n  ; Rows are delimited by linefeeds (`n).
                {
                    y =
                    RowNumber := A_Index
                    Loop, Parse, A_LoopField, %A_Tab%  ; Fields (columns) in each row are delimited by tabs (A_Tab).
                    {
                        y .= A_LoopField "`t"
                        If A_Index = 5
                        {
                            text := SubStr(A_LoopField,2)
                            number := text*1
                            D += number
                        }
                    }
                    x .= y "`n"

                }
            calcTotal := "$" . format("{:.2f}",D)
            msgbox,4096,Volume of Selected Lines, %calcTotal%
            clipboard := x
return