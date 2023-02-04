#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance force
SetTitleMatchMode 2

#Include Library\csv.ahk

SetWorkingDir %A_ScriptDir%
ComINIPath = ..\Data\Search.ini
MfgCodeData = ..\Data\MfgCodes.csv
AddressData = ..\Data\Addresses.csv

CSV_Load(MfgCodeData,"mfgData")
mfgNames := CSV_ReadCol("mfgData",1)
CSV_LOad(AddressData,"addData")
addresses := CSV_ReadCol("addData",1)

;msgbox % addresses
salesCode := "12"

IniRead,jobNo,%ComINIPath%,CurrentProject,JobNumber
IniRead,OrdersFolderPath,%ComINIPath%,CurrentProject,OrdersFolderPath
IniRead,jobName,%ComINIPath%,CurrentProject,JobName

poDate := A_Now
FormatTime, poDate, %poDate%, MM/dd/yyyy

filepath :=
lnQty1 := 0
lnQty2 := 0
lnQty3 := 0
lnQty4 := 0
lnQty5 := 0
lnUnitPrice1 := 0
lnUnitPrice2 := 0
lnUnitPrice3 := 0
lnUnitPrice4 := 0
lnUnitPrice5 := 0
lnTotalPrice1 := 0
lnTotalPrice2 := 0
lnTotalPrice3 := 0
lnTotalPrice4 := 0
lnTotalPrice5 := 0
multi := 1.00
freight := 0
grandTotal := 0

Gui, Font, s11, Calibri  ; Set 10-point Verdana.
Gui, Add, Tab3,, Address|Line Items

Gui, Add, ComboBox, vmfgPOname x90 y45 w124 h35 R10 sort hwndhMfgNames gupdatePO, %mfgNames%
PostMessage, 0x0153, -1, 28,, ahk_id %hMfgNames%  ; Set height of selection field.
PostMessage, 0x0153,  0, 28,, ahk_id %hMfgNames%  ; Set height of list items.

Gui, Add, Edit, vmfgCode x225 y45 w50
Gui, Add, Edit, vpoNumber x285 y45 w134 , %poNumber%

Gui, Add, Combobox, vshipToName x90 y93 w326 R10 hwndhshipTo gupdateAddress, %addresses%
PostMessage, 0x0153, -1, 28,, ahk_id %hshipTo%  ; Set height of selection field.
PostMessage, 0x0153,  0, 28,, ahk_id %hshipTo%  ; Set height of list items.

Gui, Add, Edit, vAddress1 x90 y141 w326, %Address1%
Gui, Add, Edit, vAddress2 x90 y189 w326, %Address2%
Gui, Add, Edit, vAddress3 x90 y237 w326, %Address3%
Gui, Add, Edit, vCity x90 y285 w163, %City%
Gui, Add, Edit, vState x90 y333 w163, %State%
Gui, Add, Edit, vZip x90 y381 w163, %Zip%
Gui, Add, Edit, vcontactName x90 y429 w163, %contactName%
Gui, Add, Edit, vcontactNum x285 y429 w163, %contactNum%

;------------------------------------- Text Labels
Gui, Add, Text, x20 y45 w57, Supplier
;Gui, Add, Text, x285 y45 w76, PO Number
Gui, Add, Text, x20 y93 w57, Ship To
Gui, Add, Text, x20 y141 w57, Address 1
Gui, Add, Text, x20 y189 w57, Address 2
Gui, Add, Text, x20 y237 w57, Address 3
Gui, Add, Text, x20 y285 w57, City
Gui, Add, Text, x20 y333 w57, State
Gui, Add, Text, x20 y381 w57, Zip
Gui, Add, Text, x20 y429 w57, Contact
Gui, Add, Text, x20 y477 w57, Date

;------------------------------------- Text Labels


Gui, Add, DateTime, x90 y477 w163 vDate,
Gui, Add, Button, gButton1 x90 y535 w144, Submit
Gui, Add, Button, gUpdateAdd x285 y535 w144, Grab Address
Gui, Add, Button, gCancel x475 y535 w144, Cancel
Gui, Show, w650 h580, PO Generator - %jobNo% - %jobName%


Gui, Tab, 2

Gui, Add, Edit, vlnQty1 x75 y93 w30, %lnQty1%
Gui, Add, Edit, vlnDesc1  x120 y93 w275 r2 limit72, %lnDesc1%
Gui, Add, Edit, vlnUnitPrice1 x400 y93 w90, %lnUnitPrice1%
Gui, Add, Edit, vlnTotalPrice1 x500 y93 w90, %lnTotalPrice1%

Gui, Add, Edit, vlnQty2 x75 y141 w30, %lnQty2%
Gui, Add, Edit, vlnDesc2  x120 y141 w275 r2 limit72, %lnDesc2%
Gui, Add, Edit, vlnUnitPrice2 x400 y141 w90, %lnUnitPrice2%
Gui, Add, Edit, vlnTotalPrice2 x500 y141 w90, %lnTotalPrice2%

Gui, Add, Edit, vlnQty3 x75 y189 w30, %lnQty3%
Gui, Add, Edit, vlnDesc3  x120 y189 w275 r2 limit72, %lnDesc3%
Gui, Add, Edit, vlnUnitPrice3 x400 y189 w90, %lnUnitPrice3%
Gui, Add, Edit, vlnTotalPrice3 x500 y189 w90, %lnTotalPrice3%

Gui, Add, Edit, vlnQty4 x75 y237 w30, %lnQty4%
Gui, Add, Edit, vlnDesc4  x120 y237 w275 r2 limit72, %lnDesc4%
Gui, Add, Edit, vlnUnitPrice4 x400 y237 w90, %lnUnitPrice4%
Gui, Add, Edit, vlnTotalPrice4 x500 y237 w90, %lnTotalPrice4%

Gui, Add, Edit, vlnQty5 x75 y285 w30, %lnQty5%
Gui, Add, Edit, vlnDesc5  x120 y285 w275 r2 limit72, %lnDesc5%
Gui, Add, Edit, vlnUnitPrice5 x400 y285 w90, %lnUnitPrice5%
Gui, Add, Edit, vlnTotalPrice5 x500 y285 w90, %lnTotalPrice5%

Gui, Add, Edit, vcustPoNum x120 y390 w200, %custPoNum%
Gui, Add, Edit, vcustPoVolume x120 y430 w200, %custPoVolume%

Gui, Add, Edit, vsubTotal   x500 y343 w90, %subTotal%
Gui, Add, Edit, vmulti      x500 y391 w90, %multi%
Gui, Add, Edit, vfreight    x500 y439 w90, %freight%
Gui, Add, Edit, vgrandTotal x500 y487 w90, %grandTotal%

Gui, Add, Text, x78 y70 w30, Qty
Gui, Add, Text, x210 y70 w30, Description
Gui, Add, Text, x415 y70 w70, Unit Price
Gui, Add, Text, x520 y70 w30, Total
Gui, Add, Text, x20 y93 w50, Line 1
Gui, Add, Text, x20 y141 w50, Line 2
Gui, Add, Text, x20 y189 w50, Line 3
Gui, Add, Text, x20 y237 w50, Line 4
Gui, Add, Text, x20 y285 w50, Line 5
Gui, Add, Text, x20 y390 w97, Cust PO Number
Gui, Add, Text, x20 y430 w97, Cust PO Volume


Gui, Add, Text, x415 y343 w70, Sub Total
Gui, Add, Text, x415 y391 w70, Multi
Gui, Add, Text, x415 y439 w70, Freight
Gui, Add, Text, x415 y487 w70, Total

Gui, Add, Button, gButton2 x90 y535 w144, Submit
Gui, Add, Button, gButton3 x300 y535 w144, Goto Folder
Gui, Add, Button, gPDF x475 y535 w144, PDF
return

Cancel:
GuiClose:
ExitApp

updatePO:
Gui Submit, NoHide
Result := CSV_Search("mfgData", mfgPOname)
Result := StrSplit(Result,",")
Result := CSV_ReadCell("mfgData",Result[1],Result[2]+1)

	if Result !=
	{
		GuiControl,, mfgCode, %result%
		GuiControl,,poNumber, % jobNo . result . salesCode

	}
	else
	{
		GuiControl,, mfgCode, SVS
		GuiControl,,poNumber, % jobNo . "SVS" . salesCode

	}

return

updateAddress:
Gui Submit, NoHide
Result := CSV_Search("addData", shipToName)
Result := StrSplit(Result,",")
shipToName := CSV_ReadCell("addData",Result[1],Result[2]+1)
shipToPhone := CSV_ReadCell("addData",Result[1],Result[2]+2)
address1 := CSV_ReadCell("addData",Result[1],Result[2]+3)
city := CSV_ReadCell("addData",Result[1],Result[2]+4)
state := CSV_ReadCell("addData",Result[1],Result[2]+5)
zip := CSV_ReadCell("addData",Result[1],Result[2]+6)

GuiControl,, address1, %address1%
GuiControl,, address2, %address2%
GuiControl,, address3, %address3%
GuiControl,, city, %city%
GuiControl,, state, %state%
GuiControl,, zip, %zip%
GuiControl,, contactName, %shipToName%
GuiControl,, contactNum, %shipToPhone%
return

Button1:
gui Submit, NoHide
MsgBox, % poNumber "`n" shipToName "`n" mfgcode
return

Button2:
gui Submit, NoHide
lnTotalPrice1 := lnQty1*lnUnitPrice1
lnTotalPrice2 := lnQty2*lnUnitPrice2
lnTotalPrice3 := lnQty3*lnUnitPrice3
lnTotalPrice4 := lnQty4*lnUnitPrice4
lnTotalPrice5 := lnQty5*lnUnitPrice5
GuiControl,, lnTotalPrice1, %lnTotalPrice1%
GuiControl,, lnTotalPrice2, %lnTotalPrice2%
GuiControl,, lnTotalPrice3, %lnTotalPrice3%
GuiControl,, lnTotalPrice4, %lnTotalPrice4%
subtotal := lnTotalPrice1+lnTotalPrice2+lnTotalPrice3+lnTotalPrice4
GuiControl,, subtotal, %subtotal%
grandTotal := (subTotal*multi)+freight
GuiControl,, grandTotal, %grandTotal%
return

Button3:
gui Submit, NoHide

OrdersFolder := OrdersFolderPath "\" mfgPOname
msgbox % OrdersFolder
If FileExist(OrdersFolder)
	{
	run, %OrdersFolder%
	return
	}
msgbox, No Folder for %mfgPOname%
return

UpdateAdd:
ComINIPath = ..\Data\Search.ini
IniRead,address1,%ComINIPath%,CurrentAddress,AField1
IniRead,address2,%ComINIPath%,CurrentAddress,AField2
IniRead,address3,%ComINIPath%,CurrentAddress,AField3
IniRead,city,%ComINIPath%,CurrentAddress,CityField
IniRead,state,%ComINIPath%,CurrentAddress,StField
IniRead,zip,%ComINIPath%,CurrentAddress,ZipField
IniRead,STContact,%ComINIPath%,CurrentAddress,ContactField
IniRead,STPhone,%ComINIPath%,CurrentAddress,PhoneField
GuiControl,, address1, %address1%
GuiControl,, address2, %address2%
GuiControl,, address3, %address3%
GuiControl,, city, %city%
GuiControl,, state, %state%
GuiControl,, zip, %zip%
GuiControl,, contactName, %STContact%
GuiControl,, contactNum, %STPhone%
gui Submit, NoHide
return


PDF:
gui Submit, NoHide
MsgBox, % poNumber
OrdersFolder :=
OrdersFolder := OrdersFolderPath "\" mfgPOname
IfNotExist, %OrdersFolder%
   FileCreateDir, %OrdersFolder%
PDFscriptPath := "C:\Program Files\Bluebeam Software\Bluebeam Revu\20\Revu\ScriptEngine.exe "
BlankForm := "C:\Users\jacobe\OneDrive - Schwab Vollhaber Lubratt Inc\Desktop\AHK - Local\Project Search v4\Scripts\TemplateForms\SVL PO - empty.pdf"
FormData := OrdersFolder "\formdata.fdf"

If FileExist(OrdersFolderPath "\orders.csv")
	{
		lineCount := 1
		Loop, Read, %OrdersFolderPath%\orders.csv
		{
			searchString := A_LoopReadLine
			if InStr(searchString, mfgPOname)
				lineCount += 1
		}
	}
Else lines := 1

MsgBox, Total number of %mfgPOname% lines in the file "orders.csv" : %lineCount%

FilledForm := OrdersFolder "\SVL PO " poNumber " " lineCount ".pdf"
;msgBox, % FilledForm

;pdfScript := """Open('" BlankForm "') FormImport('" FormData "') Flatten(false,65536) Save('" FilledForm "') Close()"""
pdfScript := """Open('" BlankForm "') FormImport('" FormData "') Save('" FilledForm "') Close()"""
finalpath := PDFscriptPath . pdfScript
;msgbox % finalpath


FileAppend, %poNumber%`,%mfgPOname%`,%shipToName%`,%address1%`,%address2%`,%city%`,%state%`,%zip%`,%ContactName%`,%ContactNum%`,%lnQty1%`,%lnDesc1%`,%lnUnitPrice1%`,%lnTotalPrice1%`,%lnQty2%`,%lnDesc2%`,%lnUnitPrice2%`,%lnTotalPrice2%`,%lnQty3%`,%lnDesc3%`,%lnUnitPrice3%`,%lnTotalPrice3%`,%subTotal%`,%multi%`,%freight%`,%grandTotal%`,%poDate%`,POGUI`,%custPoNum%`,%custPoVolume%`,`n,%OrdersFolderPath%\orders.csv

;format the text for BlueBeam by removing 0 and changing dollar values to dollar formatted values
lnQty1 := checkZero(lnQty1)
lnQty2 := checkZero(lnQty2)
lnQty3 := checkZero(lnQty3)
lnQty4 := checkZero(lnQty4)
lnUnitPrice1 := checkZeroD(lnUnitPrice1)
lnUnitPrice2 := checkZeroD(lnUnitPrice2)
lnUnitPrice3 := checkZeroD(lnUnitPrice3)
lnUnitPrice4 := checkZeroD(lnUnitPrice4)
lnTotalPrice1 := checkZeroD(lnTotalPrice1)
lnTotalPrice2 := checkZeroD(lnTotalPrice2)
lnTotalPrice3 := checkZeroD(lnTotalPrice3)
lnTotalPrice4 := checkZeroD(lnTotalPrice4)
subtotal := checkZeroD(subtotal)
freight := checkZeroD(freight)
grandTotal := checkZeroD(grandTotal)


text =
(
`%FDF-1.2
1 0 obj
<<
/FDF
<<
/Fields [
<<
/V (%poNumber%)
/T (poNumber)
>>
<<
/V (%mfgPOname%)
/T (supplierName)
>>
<<
/V (%shipToName%)
/T (address1)
>>
<<
/V ()
/T (supplierAdd1)
>>
<<
/V (%address1%)
/T (address2)
>>
<<
/V ()
/T (supplierAdd2)
>>
<<
/V (%address2%)
/T (address3)
>>
<<
/V ()
/T (supplierCityStZip)
>>
<<
/V (%city%, %state%  %zip%)
/T (citystzip)
>>
<<
/V ()
/T (supplierPhone)
>>
<<
/V (%ContactName% @ %ContactNum%)
/T (shippingContact)
>>
<<
/V ()
/T (Text12)
>>
<<
/V ()
/T (tag)
>>
<<
/V (%lnQty1%)
/T (qtyLine1)
>>
<<
/V (%lnDesc1%)
/T (descLine1)
>>
<<
/V (%lnUnitPrice1%)
/T (unitPriceLine1)
>>
<<
/V (%lnTotalPrice1%)
/T (totalPriceLine1)
>>
<<
/V (%lnQty2%)
/T (qtyLine2)
>>
<<
/V (%lnDesc2%)
/T (descLine2)
>>
<<
/V (%lnUnitPrice2%)
/T (unitPriceLine2)
`)
>>
<<
/V (%lnTotalPrice2%)
/T (totalPriceLine2)
>>
<<
/V (%lnQty3%)
/T (qtyLine3)
>>
<<
/V (%lnDesc3%)
/T (descLine3)
>>
<<
/V (%lnUnitPrice3%)
/T (unitPriceLine3)
>>
<<
/V (%lnTotalPrice3%)
/T (totalPriceLine3)
>>
<<
/V (%lnQty4%)
/T (qtyLine4)
>>
<<
/V (%lnDesc4%)
/T (descLine4)
>>
<<
/V (%lnUnitPrice4%)
/T (unitPriceLine4)
>>
<<
/V (%lnTotalPrice4%)
/T (totalPriceLine4)
>>
<<
/V (%subTotal%)
/T (subTotal)
>>
<<
/V (%multi%)
/T (multiplier)
>>
<<
/V (%freight%)
/T (freight)
>>
<<
/V (%grandTotal%)
/T (totalCost)
>>
<<
/V (%poDate%)
/T (poDate)
>>]
>>
>>
endobj
trailer
<<
/Root 1 0 R
>>
`%`%EOF
)

FileDelete, % FilledForm
FileDelete, % Formdata
FileAppend, %text%, %FormData%



clipboard := finalpath
run, % finalpath
sleep, 5000
run, % filledForm
return


checkZero(var){
if var = 0
	{
		var := ""
		return var
	}
else return var
}

checkZeroD(var){
if var = 0
	{
		var := ""
		return var
	}
else
	{
	var := Format("{1:.2f}",var)
	var := "$" RegExReplace(var, "\G\d+?(?=(\d{3})+(?:\D|$))", "$0,")
	return var
	}
}
