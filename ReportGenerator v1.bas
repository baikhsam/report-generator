Attribute VB_Name = "ReportGenerator"
Sub CommandButton1()

End Sub
Sub GenerateReport()
Attribute GenerateReport.VB_ProcData.VB_Invoke_Func = " \n14"
'HongSeok, Sam, Baik
'7/13/2018
'Macro created to generate B&L report from given data excel file


Dim ws As Worksheet
Set ws = Sheets.Add

'Creating Skeleton of Report

With ws.Columns("B")
.ColumnWidth = .ColumnWidth * 3
End With
With ws.Columns("M")
.ColumnWidth = .ColumnWidth * 2
End With
With ws.Columns("F")
.ColumnWidth = .ColumnWidth * 1.5
End With
With ws.Columns("H")
.ColumnWidth = .ColumnWidth * 0.2
End With
ws.Range("A1:P1").EntireColumn.AutoFit
ws.Cells.Font.Name = "Arial"
ws.Cells.Font.Size = 7
ws.Cells.VerticalAlignment = xlBottom

ws.Range("A1") = "GLOBAL ENERGY TRADING PTE LTD BORROW/LOAN FOR CONSORT BUNKERS"
ws.Range("A1").Font.Size = 10
ws.Range("A1").Font.Bold = True
ws.Range("A1").Font.Name = "Garamond"
ws.Range("A3") = " FOR THE MONTH OF :"
'ws.Range("A3").WrapText = True
ws.Range("A3").Font.Bold = True
ws.Range("C3") = "MONTH"
ws.Range("D3") = "YEAR"
ws.Range("C3").Font.Bold = True
ws.Range("D3").Font.Bold = True
ws.Range("G3") = "PAGE    1/2"
ws.Range("N3") = "PAGE    2/2"
ws.Range("G3").Font.Italic = True
ws.Range("O3").Font.Italic = True

ws.Range("A5") = "LOADING"
ws.Range("A7") = "BEFORE"
ws.Range("A31") = "TOTAL"
ws.Range("A6") = "DATE"
ws.Range("B6") = "VESSEL/BARGE"
ws.Range("C6") = "LOCN"
ws.Range("D6") = "380CST"
ws.Range("E6") = "500CST"
ws.Range("F6") = "BARGE/TMNL"
ws.Range("G6") = "REMARKS"

ws.Range("I5") = "DELIVERY"
ws.Range("I7") = "BEFORE"
ws.Range("I31") = "TOTAL"
ws.Range("I6") = "DATE"
ws.Range("J6") = "VESSEL"
ws.Range("K6") = "380CST"
ws.Range("L6") = "500CST"
ws.Range("M6") = "BARGE/TMNL"
ws.Range("N6") = "REMARKS"

ws.Range("A34") = "CONSORT BUNKERS B&L MONTHLY REPORT FOR MTH OF :"
ws.Range("D34") = "=C3"

ws.Range("A5").HorizontalAlignment = xlLeft
ws.Range("A5").Font.Bold = True
ws.Range("A7").HorizontalAlignment = xlCenter
ws.Range("A7").Font.Bold = True
ws.Range("A31").HorizontalAlignment = xlCenter
ws.Range("A31").Font.Bold = True
ws.Range("A6").HorizontalAlignment = xlCenter
ws.Range("A6").Font.Bold = True
ws.Range("B6").HorizontalAlignment = xlCenter
ws.Range("B6").Font.Bold = True
ws.Range("C6").HorizontalAlignment = xlCenter
ws.Range("C6").Font.Bold = True
ws.Range("D6").HorizontalAlignment = xlRight
ws.Range("D6").Font.Bold = True
ws.Range("E6").HorizontalAlignment = xlRight
ws.Range("E6").Font.Bold = True
ws.Range("F6").HorizontalAlignment = xlCenter
ws.Range("F6").Font.Bold = True
ws.Range("G6").HorizontalAlignment = xlCenter
ws.Range("G6").Font.Bold = True

ws.Range("I5").HorizontalAlignment = xlLeft
ws.Range("I5").Font.Bold = True
ws.Range("I7").HorizontalAlignment = xlCenter
ws.Range("I7").Font.Bold = True
ws.Range("I31").HorizontalAlignment = xlCenter
ws.Range("I31").Font.Bold = True
ws.Range("I6").HorizontalAlignment = xlCenter
ws.Range("I6").Font.Bold = True
ws.Range("J6").HorizontalAlignment = xlCenter
ws.Range("J6").Font.Bold = True
ws.Range("K6").HorizontalAlignment = xlCenter
ws.Range("K6").Font.Bold = True
ws.Range("L6").HorizontalAlignment = xlRight
ws.Range("L6").Font.Bold = True
ws.Range("M6").HorizontalAlignment = xlCenter
ws.Range("M6").Font.Bold = True
ws.Range("N6").HorizontalAlignment = xlCenter
ws.Range("N6").Font.Bold = True
ws.Range("O6").HorizontalAlignment = xlCenter
ws.Range("O6").Font.Bold = True

'Creating table outline for Loading
ws.Range("A5:G31").Borders.LineStyle = xlContinuous
'ws.Range("A5:G32").WrapText = True
ws.Range("A5:G5").Merge
ws.Range("A7:C7").Merge
ws.Range("A31:C31").Merge
ws.Range("A34:C34").Merge

'Creating table outline for Delivery
ws.Range("I5:N31").Borders.LineStyle = xlContinuous
'ws.Range("I5:O32").WrapText = True
ws.Range("I5:N5").Merge
ws.Range("I7:J7").Merge
ws.Range("I31:J31").Merge

'Copying Worksheet into a new Excel Workbook file
'ws.Copy
End Sub
