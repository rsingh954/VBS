[PCOMM SCRIPT HEADER]
LANGUAGE=VBSCRIPT
DESCRIPTION=
[PCOMM SCRIPT SOURCE]
OPTION EXPLICIT
autECLSession.SetConnectionByName(ThisSessionName)

REM This line calls the macro subroutine
subSub1_

sub subSub1_()
Dim excel
Dim objWorkbook
Dim row
Dim SoldAsFlag
Dim invalidPart

Set excel = CreateObject("Excel.Application")

'CHANGE PATH TO REFLECT YOUR WORKBOOK LOCATION
Set objWorkbook = excel.Workbooks.Open("C:\Users\rsingh\Desktop\Check Sold As.xlsx")
excel.visible = true

'THIS LINE SAYS THAT AS400 WILL START AT ROW 2 IN EXCEL
row = 2

autECLSession.autECLPS.Wait 1000

'WHILE COLUMN A IS NOT BLANK
while excel.Cells(row,1).Value <> ""

   autECLSession.autECLOIA.WaitForAppAvailable
   
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys excel.Cells(row, 1).Value
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys "[fldext]"
   autECLSession.autECLPS.Wait 1000
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys excel.Cells(row, 2).Value
   autECLSession.autECLPS.Wait 1000
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys "[enter]"
   autECLSession.autECLPS.Wait 1000
   SoldAsFlag=autECLSession.autECLPS.GetText(02,012,3)
   invalidPart=autECLSession.autECLPS.GetText(01,032,3)

   if SoldAsFlag = "Sol" Then 
     autECLSession.autECLPS.SendKeys "[fldext]"
     autECLSession.autECLPS.SendKeys "[fldext]"
     excel.Cells(row,3).Value = "Sold As"
   end if

   if  invalidPart = "Par" Then 
     autECLSession.autECLPS.Wait 1000
     autECLSession.autECLPS.SendKeys "[enter]"
     excel.Cells(row,3).Value = "Invalid Part"
   end if

   autECLSession.autECLOIA.WaitForAppAvailable

   excel.Cells(row,4).Value = "Entered"
   row = row + 1
wend
Set excel = Nothing
end sub
