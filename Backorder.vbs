[PCOMM SCRIPT HEADER]
LANGUAGE=VBSCRIPT
DESCRIPTION=
[PCOMM SCRIPT SOURCE]
OPTION EXPLICIT
autECLSession.SetConnectionByName(ThisSessionName)
Dim rows 
rows =2
REM This line calls the macro subroutine
subSub1_(rows)

sub subSub1_(row)
Dim Rows 
Dim excel
Dim objWorkbook
Dim count
Dim check
Dim rname
dim MFG
dim err
dim Ln
Dim iLastRow
dim invalidPO
Dim oCountWS 
Set excel = CreateObject("Excel.Application")
Set objWorkbook = excel.Workbooks.Open("G:\Raja\Back Order Automate 3.xlsx")
excel.visible = true
Set oCountWS = objWorkbook.Worksheets("Sheet1")
iLastRow = oCountWS.UsedRange.Rows.Count

 while excel.Cells(row,7).Value <> ""

   autECLSession.autECLOIA.WaitForAppAvailable
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys excel.Cells(row, 1).Value
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys "[fldext]"
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys "2"
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys "[enter]"
   autECLSession.autECLPS.Wait 1000
   
   'CHECKING FOR INVALID PO
   invalidPO= autECLSession.autECLPS.GetText(24,002,3)
   if invalidPO = "Inv" Then 
      excel.Cells(row,13).Value = "Invalid PO"
      row = row + 1
      excel.DisplayAlerts = False
      objWorkbook.Close True
      Exit Sub
   End if


   autECLSession.autECLPS.WaitForCursor 10,2,10000
   autECLSession.autECLOIA.WaitForAppAvailable
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys "[backtab]"
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys excel.Cells(row, 7).Value
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys "[enter]"
   autECLSession.autECLPS.Wait 1000
   autECLSession.autECLOIA.WaitForAppAvailable
   autECLSession.autECLOIA.WaitForInputReady
   Rows = autECLSession.autECLPS.GetText(12,002,3)
   autECLSession.autECLPS.SendKeys Rows
   autECLSession.autECLPS.Wait 0050
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys "[enter]"
   autECLSession.autECLPS.Wait 1000
   
   'CHECKING IF LINE ITEM
   Ln =  CStr(autECLSession.autECLPS.GetText(24,2,1))
   If Ln = "L" Then
      excel.Cells(row,13).Value = "Line Item"
      autECLSession.autECLOIA.WaitForInputReady
      autECLSession.autECLPS.SendKeys "[pf3]" 
      autECLSession.autECLOIA.WaitForInputReady
      autECLSession.autECLPS.SendKeys "2"
      autECLSession.autECLOIA.WaitForInputReady
      autECLSession.autECLPS.SendKeys "[enter]" 
      excel.DisplayAlerts = False
      objWorkbook.Close True
      row = row + 1
      subSub1_(row)
   End If 

   autECLSession.autECLOIA.WaitForAppAvailable

    autECLSession.autECLOIA.WaitForInputReady
    autECLSession.autECLPS.SendKeys "[backtab]"
    autECLSession.autECLOIA.WaitForInputReady
    autECLSession.autECLPS.SendKeys "[backtab]"
    autECLSession.autECLOIA.WaitForInputReady
    autECLSession.autECLPS.SendKeys "[backtab]"
    autECLSession.autECLOIA.WaitForInputReady
    autECLSession.autECLPS.SendKeys excel.Cells(row, 12).Value
    autECLSession.autECLOIA.WaitForInputReady
    autECLSession.autECLPS.SendKeys "[fldext]"
    autECLSession.autECLOIA.WaitForInputReady
    autECLSession.autECLPS.SendKeys "[fldext]"
    autECLSession.autECLOIA.WaitForInputReady
    autECLSession.autECLPS.SendKeys "[enter]"
    autECLSession.autECLPS.Wait 2000

    'CHECKING IF ORDER IS OPEN
    err =  CStr(autECLSession.autECLPS.GetText(24,2,1))
      If err = "O" Then 
        autECLSession.autECLOIA.WaitForInputReady
        autECLSession.autECLPS.SendKeys "[pf3]" 
        autECLSession.autECLOIA.WaitForInputReady
        autECLSession.autECLPS.SendKeys "[pf3]" 
        autECLSession.autECLOIA.WaitForInputReady
        autECLSession.autECLPS.SendKeys "2"
        autECLSession.autECLOIA.WaitForInputReady
        autECLSession.autECLPS.SendKeys "[enter]"
        excel.Cells(row,13).Value = "Open Order"
        excel.DisplayAlerts = False
        objWorkbook.Close True
        row = row + 1
        subSub1_(row)
      End If
      'CHECKING IF VENDOR QTY = 0
      If err = "V" Then
         excel.Cells(row,13).Value = "Vendor Qty 0"
         autECLSession.autECLOIA.WaitForInputReady
         autECLSession.autECLPS.SendKeys "[pf3]" 
         autECLSession.autECLOIA.WaitForInputReady
         autECLSession.autECLPS.SendKeys "[pf3]" 
         autECLSession.autECLOIA.WaitForInputReady
         autECLSession.autECLPS.SendKeys "2"
         autECLSession.autECLOIA.WaitForInputReady
         autECLSession.autECLPS.SendKeys "[enter]"
         excel.Cells(row,13).Value = "Vendor Qty 0"
         excel.DisplayAlerts = False
         objWorkbook.Close True
         row = row + 1
         subSub1_(row)
      End If

        autECLSession.autECLOIA.WaitForInputReady
        autECLSession.autECLPS.SendKeys "[pf8]" 
        excel.Cells(row,13).Value = "Entered"
        row = row + 1
 wend

Set excel = Nothing
Set rname = Nothing
Set MFG = Nothing

End Sub

