Attribute VB_Name = "ReportGenerator"
Sub GenerateReport()
'HongSeok (Sam) Baik for Global Energy Trading Ltd
'7/16/2018
'Macro created to generate B&L report from a given data excel file

Application.ScreenUpdating = False

'Checking if everything was properly set up
output = MsgBox("Did you name the Excel Worksheet you are trying to generate a report with, DATA?", vbYesNoCancel, "Report Generator")
If output = vbNo Then
    MsgBox ("In order to successfully generate a report, please ensure the Excel Worksheet that you are summarizing is named DATA. Thank you.")
    Exit Sub
ElseIf output = vbCancel Then
    Exit Sub
End If


Dim ask_month As String
Dim ask_name As String
Dim outsider1 As String
Dim outsider2 As String

ask_month = InputBox("Enter the month this report is for: ", "Enter month", "Feb")
ask_name = InputBox("Enter the name you wish to create the report under", "Enter name", "ENTER")
ask_name = WorksheetFunction.Proper(ask_name)

ask_month = WorksheetFunction.Proper(ask_month)
outsider1 = InputBox("Please enter the abbreviated code name of the B & L outsider you wish the report be made for: ", "Enter code name", "TS")
outsider2 = InputBox("SECOND OUTSIDER OPTIONAL: Please enter another abbreviated code name of the B & L outsider you wish the report be made for: ", "OPTIONAL: Enter code name")


Worksheets("DATA").Activate
Range("A2:AN4000").Sort _
Key1:=Range("I1"), Order1:=xlAscending

Dim ws As Worksheet
Set ws = Sheets.Add
ws.Name = "Report " & ws.Name

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
With ws.Columns("J")
    .ColumnWidth = .ColumnWidth * 3
End With
With ws.Columns("H")
.ColumnWidth = .ColumnWidth * 0.2
End With
ws.Range("A1:P1").EntireColumn.AutoFit
ws.Cells.Font.Name = "Arial"
ws.Cells.Font.Size = 7
ws.Cells.VerticalAlignment = xlBottom

ws.Range("A1") = "GLOBAL ENERGY TRADING PTE LTD BORROW/LOAN FOR " & outsider1 & " BUNKERS"
ws.Range("A1").Font.Size = 10
ws.Range("A1").Font.Bold = True
ws.Range("A1").Font.Name = "Garamond"
ws.Range("A3") = " FOR THE MONTH OF :"
'ws.Range("A3").WrapText = True
ws.Range("A3").Font.Bold = True
ws.Range("C3") = ask_month
ws.Range("C3").Interior.ColorIndex = 6
ws.Range("D102").Interior.ColorIndex = 6
ws.Range("B108").Interior.ColorIndex = 6
ws.Range("C108").Interior.ColorIndex = 6
ws.Range("D108").Interior.ColorIndex = 36
ws.Range("E108").Interior.ColorIndex = 36
ws.Range("D3") = "2018"
ws.Range("C3").Font.Bold = True
ws.Range("D3").Font.Bold = True
ws.Range("G3") = "PAGE    1/2"
ws.Range("N3") = "PAGE    2/2"
ws.Range("G3").Font.Italic = True
ws.Range("N3").Font.Italic = True

ws.Range("A5") = "LOADING"
ws.Range("A7") = "BEFORE"
ws.Range("A100") = "TOTAL"
ws.Range("A6") = "DATE"
ws.Range("B6") = "VESSEL/BARGE"
ws.Range("C6") = "LOCN"
ws.Range("D6") = "380CST"
ws.Range("E6") = "500CST"
ws.Range("F6") = "BARGE/TMNL"
ws.Range("G6") = "REMARKS"
ws.Range("D100") = "=SUM(D7:D99)"
ws.Range("E100") = "=SUM(E7:E99)"
ws.Range("D106") = "=D100"
ws.Range("E106") = "=E100"

ws.Range("D108") = "=D106-D107"
ws.Range("I102") = "TO:"
ws.Range("J102") = "ENTER"
ws.Range("I103") = "ATTN:"
ws.Range("J103") = "ENTER"
ws.Range("J105") = "PLEASE RETURN SIGN ACKNOWLEDGEMENT THIS COPY OF STOCK"
ws.Range("J106") = "TRANSACTION CONFIRMATION WITH STAMP OF THE COMPANY"
ws.Range("J107") = "AND EMAIL IN RETURN SOONEST."
ws.Range("J109") = "ACKNOWLEDGEMENT BY:"
ws.Range("J111") = "NAME:"
ws.Range("J112") = "DATE:"
ws.Range("J113") = "COMPANY STAMP:"

ws.Range("I5") = "DELIVERY"
ws.Range("I7") = "BEFORE"
ws.Range("I100") = "TOTAL"
ws.Range("I6") = "DATE"
ws.Range("J6") = "VESSEL"
ws.Range("K6") = "380CST"
ws.Range("L6") = "500CST"
ws.Range("M6") = "BARGE/TMNL"
ws.Range("N6") = "REMARKS"
ws.Range("D107") = "=K100"
ws.Range("E107") = "=L100"
ws.Range("A102") = "CONSORT BUNKERS B&L MONTHLY REPORT FOR MTH OF :"
ws.Range("A106") = "TOTAL LOADED QUANTITY BY CONSORT"
ws.Range("A107") = "TOTAL DELIVERED QUANTITY BY CONSORT"
ws.Range("A108") = "B & L"
ws.Range("A110") = "THANK YOU FOR YOUR VALUED SUPPORT AND COORPERATION"
ws.Range("A112") = "YOURS FAITHFULLY"
ws.Range("A113") = "GLOBAL ENERGY TRADING PTE LTD"
ws.Range("A114") = UCase(ask_name)
ws.Range("D102") = "=C3"
ws.Range("D105") = "380CST"
ws.Range("E105") = "500CST"
ws.Range("K100") = "=SUM(K7:K99)"
ws.Range("L100") = "=SUM(L7:L99)"
ws.Range("D107") = "=K100"
ws.Range("E107") = "=L100"
ws.Range("E108") = "=E106-E107"
ws.Range("B108") = "ENTER"
ws.Range("C108") = "ENTER"

ws.Range("A100").Font.Bold = True
ws.Range("D100").Font.Bold = True
ws.Range("E100").Font.Bold = True
ws.Range("K100").Font.Bold = True
ws.Range("L100").Font.Bold = True
ws.Range("A106").Font.Bold = True
ws.Range("A107").Font.Bold = True
ws.Range("A108").Font.Bold = True
ws.Range("I102").Font.Bold = True
ws.Range("J102").Font.Bold = True
ws.Range("I103").Font.Bold = True
ws.Range("J103").Font.Bold = True
ws.Range("D105").Font.Bold = True
ws.Range("D105").HorizontalAlignment = xlRight
ws.Range("E105").Font.Bold = True
ws.Range("E105").HorizontalAlignment = xlRight
ws.Range("D108").Font.Bold = True
ws.Range("E108").Font.Bold = True

ws.Range("F105").Font.Bold = True
ws.Range("F105").HorizontalAlignment = xlRight
ws.Range("G105").Font.Bold = True
ws.Range("G105").HorizontalAlignment = xlRight
ws.Range("H105").Font.Bold = True
ws.Range("H105").HorizontalAlignment = xlRight
ws.Range("F108").Font.Bold = True
ws.Range("G108").Font.Bold = True
ws.Range("H108").Font.Bold = True

With Range("D108").Borders(xlEdgeTop)
.LineStyle = xlContinuous
.Weight = xlThin
End With
With Range("D108").Borders(xlEdgeBottom)
.LineStyle = xlDouble
.Weight = xlThick
End With
With Range("E108").Borders(xlEdgeTop)
.LineStyle = xlContinuous
.Weight = xlThin
End With
With Range("E108").Borders(xlEdgeBottom)
.LineStyle = xlDouble
.Weight = xlThick
End With

    
ws.Range("A5").HorizontalAlignment = xlLeft
ws.Range("A5").Font.Bold = True
ws.Range("A7").HorizontalAlignment = xlCenter
ws.Range("A7").Font.Bold = True
ws.Range("A100").HorizontalAlignment = xlCenter
ws.Range("A100").Font.Bold = True
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
ws.Range("I100").HorizontalAlignment = xlCenter
ws.Range("I100").Font.Bold = True
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
ws.Range("A5:G100").Borders.LineStyle = xlContinuous
ws.Range("D105:E105").Borders.LineStyle = xlContinuous
'ws.Range("A5:G32").WrapText = True
ws.Range("A5:G5").Merge
ws.Range("A7:C7").Merge
ws.Range("A100:C100").Merge
ws.Range("A103:C103").Merge

'Creating table outline for Delivery
ws.Range("I5:N100").Borders.LineStyle = xlContinuous
'ws.Range("I5:O32").WrapText = True
ws.Range("I5:N5").Merge
ws.Range("I7:J7").Merge
ws.Range("I100:J100").Merge
ws.Range("J111:K111").Merge
ws.Range("J112:K112").Merge
ws.Range("J113:K113").Merge

'Sorting data by BDR/CQ Date
Dim copysheet As Worksheet
Set copysheet = ThisWorkbook.Worksheets("DATA")
copysheet.Copy ThisWorkbook.Sheets(Sheets.Count)
'copysheet.Name = "Copysheet " & copysheet.Name

Dim ws2 As Worksheet
If outsider2 <> "" Then
    'Setting up second spreadsheet for another outsider if requested
    Set ws2 = Sheets.Add
    ws2.Name = "Second Report " & ws2.Name
    
    'Creating Skeleton of Report 2
    With ws2.Columns("B")
    .ColumnWidth = .ColumnWidth * 3
    End With
    With ws2.Columns("M")
    .ColumnWidth = .ColumnWidth * 2
    End With
    With ws2.Columns("F")
    .ColumnWidth = .ColumnWidth * 1.5
    End With
    With ws2.Columns("J")
    .ColumnWidth = .ColumnWidth * 3
    End With
    With ws2.Columns("H")
    .ColumnWidth = .ColumnWidth * 0.2
    End With
    ws2.Range("A1:P1").EntireColumn.AutoFit
    ws2.Cells.Font.Name = "Arial"
    ws2.Cells.Font.Size = 7
    ws2.Cells.VerticalAlignment = xlBottom
    
    ws2.Range("A1") = "GLOBAL ENERGY TRADING PTE LTD BORROW/LOAN FOR " & outsider2 & " BUNKERS"
    ws2.Range("A1").Font.Size = 10
    ws2.Range("A1").Font.Bold = True
    ws2.Range("A1").Font.Name = "Garamond"
    ws2.Range("A3") = " FOR THE MONTH OF :"
    ws2.Range("A3").Font.Bold = True
    ws2.Range("C3") = ask_month
    ws2.Range("C3").Interior.ColorIndex = 6
    ws2.Range("D102").Interior.ColorIndex = 6
    ws2.Range("B108").Interior.ColorIndex = 6
    ws2.Range("C108").Interior.ColorIndex = 6
    ws2.Range("D108").Interior.ColorIndex = 36
    ws2.Range("E108").Interior.ColorIndex = 36
    ws2.Range("D3") = "2018"
    ws2.Range("C3").Font.Bold = True
    ws2.Range("D3").Font.Bold = True
    ws2.Range("G3") = "PAGE    1/2"
    ws2.Range("N3") = "PAGE    2/2"
    ws2.Range("G3").Font.Italic = True
    ws2.Range("N3").Font.Italic = True
    
    ws2.Range("A5") = "LOADING"
    ws2.Range("A7") = "BEFORE"
    ws2.Range("A100") = "TOTAL"
    ws2.Range("A6") = "DATE"
    ws2.Range("B6") = "VESSEL/BARGE"
    ws2.Range("C6") = "LOCN"
    ws2.Range("D6") = "380CST"
    ws2.Range("E6") = "500CST"
    ws2.Range("F6") = "BARGE/TMNL"
    ws2.Range("G6") = "REMARKS"
    ws2.Range("D100") = "=SUM(D7:D99)"
    ws2.Range("E100") = "=SUM(E7:E99)"
    ws2.Range("D106") = "=D100"
    ws2.Range("E106") = "=E100"
    ws2.Range("D108") = "=D106-D107"
    ws2.Range("I102") = "TO:"
    ws2.Range("J102") = "ENTER"
    ws2.Range("I103") = "ATTN:"
    ws2.Range("J103") = "ENTER"
    ws2.Range("J105") = "PLEASE RETURN SIGN ACKNOWLEDGEMENT THIS COPY OF STOCK"
    ws2.Range("J106") = "TRANSACTION CONFIRMATION WITH STAMP OF THE COMPANY"
    ws2.Range("J107") = "AND EMAIL IN RETURN SOONEST."
    ws2.Range("J109") = "ACKNOWLEDGEMENT BY:"
    ws2.Range("J111") = "NAME:"
    ws2.Range("J112") = "DATE:"
    ws2.Range("J113") = "COMPANY STAMP:"
    ws2.Range("I5") = "DELIVERY"
    ws2.Range("I7") = "BEFORE"
    ws2.Range("I100") = "TOTAL"
    ws2.Range("I6") = "DATE"
    ws2.Range("J6") = "VESSEL"
    ws2.Range("K6") = "380CST"
    ws2.Range("L6") = "500CST"
    ws2.Range("M6") = "BARGE/TMNL"
    ws2.Range("N6") = "REMARKS"
    ws2.Range("D107") = "=K100"
    ws2.Range("E107") = "=L100"
    ws2.Range("A102") = "CONSORT BUNKERS B&L MONTHLY REPORT FOR MTH OF :"
    ws2.Range("A106") = "TOTAL LOADED QUANTITY BY CONSORT"
    ws2.Range("A107") = "TOTAL DELIVERED QUANTITY BY CONSORT"
    ws2.Range("A108") = "B & L"
    ws2.Range("A110") = "THANK YOU FOR YOUR VALUED SUPPORT AND COORPERATION"
    
    ws2.Range("A112") = "YOURS FAITHFULLY"
    ws2.Range("A113") = "GLOBAL ENERGY TRADING PTE LTD"
    ws2.Range("A114") = UCase(ask_name)
    ws2.Range("D102") = "=C3"
    ws2.Range("D105") = "380CST"
    ws2.Range("E105") = "500CST"
    ws2.Range("K100") = "=SUM(K7:K99)"
    ws2.Range("L100") = "=SUM(L7:L99)"
    ws2.Range("D107") = "=K100"
    ws2.Range("E107") = "=L100"
    ws2.Range("E108") = "=E106-E107"
    ws2.Range("B108") = "ENTER"
    ws2.Range("C108") = "ENTER"
    
    ws2.Range("A100").Font.Bold = True
    ws2.Range("D100").Font.Bold = True
    ws2.Range("E100").Font.Bold = True
    ws2.Range("K100").Font.Bold = True
    ws2.Range("L100").Font.Bold = True
    ws2.Range("A106").Font.Bold = True
    ws2.Range("A107").Font.Bold = True
    ws2.Range("A108").Font.Bold = True
    ws2.Range("I102").Font.Bold = True
    ws2.Range("J102").Font.Bold = True
    ws2.Range("I103").Font.Bold = True
    ws2.Range("J103").Font.Bold = True
    ws2.Range("D105").Font.Bold = True
    ws2.Range("D105").HorizontalAlignment = xlRight
    ws2.Range("E105").Font.Bold = True
    ws2.Range("E105").HorizontalAlignment = xlRight
    ws2.Range("D108").Font.Bold = True
    ws2.Range("E108").Font.Bold = True
    With Range("D108").Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With Range("D108").Borders(xlEdgeBottom)
    .LineStyle = xlDouble
    .Weight = xlThick
    End With
    With Range("E108").Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With Range("E108").Borders(xlEdgeBottom)
    .LineStyle = xlDouble
    .Weight = xlThick
    End With
    With Range("F108").Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With Range("G108").Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With Range("H108").Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    
    ws2.Range("A5").HorizontalAlignment = xlLeft
    ws2.Range("A5").Font.Bold = True
    ws2.Range("A7").HorizontalAlignment = xlCenter
    ws2.Range("A7").Font.Bold = True
    ws2.Range("A100").HorizontalAlignment = xlCenter
    ws2.Range("A100").Font.Bold = True
    ws2.Range("A6").HorizontalAlignment = xlCenter
    ws2.Range("A6").Font.Bold = True
    ws2.Range("B6").HorizontalAlignment = xlCenter
    ws2.Range("B6").Font.Bold = True
    ws2.Range("C6").HorizontalAlignment = xlCenter
    ws2.Range("C6").Font.Bold = True
    ws2.Range("D6").HorizontalAlignment = xlRight
    ws2.Range("D6").Font.Bold = True
    ws2.Range("E6").HorizontalAlignment = xlRight
    ws2.Range("E6").Font.Bold = True
    ws2.Range("F6").HorizontalAlignment = xlCenter
    ws2.Range("F6").Font.Bold = True
    ws2.Range("G6").HorizontalAlignment = xlCenter
    ws2.Range("G6").Font.Bold = True
    
    ws2.Range("I5").HorizontalAlignment = xlLeft
    ws2.Range("I5").Font.Bold = True
    ws2.Range("I7").HorizontalAlignment = xlCenter
    ws2.Range("I7").Font.Bold = True
    ws2.Range("I100").HorizontalAlignment = xlCenter
    ws2.Range("I100").Font.Bold = True
    ws2.Range("I6").HorizontalAlignment = xlCenter
    ws2.Range("I6").Font.Bold = True
    ws2.Range("J6").HorizontalAlignment = xlCenter
    ws2.Range("J6").Font.Bold = True
    ws2.Range("K6").HorizontalAlignment = xlCenter
    ws2.Range("K6").Font.Bold = True
    ws2.Range("L6").HorizontalAlignment = xlRight
    ws2.Range("L6").Font.Bold = True
    ws2.Range("M6").HorizontalAlignment = xlCenter
    ws2.Range("M6").Font.Bold = True
    ws2.Range("N6").HorizontalAlignment = xlCenter
    ws2.Range("N6").Font.Bold = True
    ws2.Range("O6").HorizontalAlignment = xlCenter
    ws2.Range("O6").Font.Bold = True
    
    'Creating table outline for Loading
    ws2.Range("A5:G100").Borders.LineStyle = xlContinuous
    ws2.Range("D105:E105").Borders.LineStyle = xlContinuous
    'ws2.Range("A5:G32").WrapText = True
    ws2.Range("A5:G5").Merge
    ws2.Range("A7:C7").Merge
    ws2.Range("A100:C100").Merge
    ws2.Range("A103:C103").Merge
    
    'Creating table outline for Delivery
    ws2.Range("I5:N100").Borders.LineStyle = xlContinuous
    'ws2.Range("I5:O32").WrapText = True
    ws2.Range("I5:N5").Merge
    ws2.Range("I7:J7").Merge
    ws2.Range("I100:J100").Merge
    ws2.Range("J111:K111").Merge
    ws2.Range("J112:K112").Merge
    ws2.Range("J113:K113").Merge
End If

Dim row As Integer
'row keeps track of Copysheet row
row = 2
Dim holder As Double
holder = 0
Dim wrow As Integer
'wrow keeps track of report sheet row of loading
wrow = 8
Dim drow As Integer
'drow keeps track of report sheet row of delivery
drow = 8
Dim i As Integer
i = 1
'month_track keeps track of the month integer value of the user inputted month
Dim month_track As Integer
month_track = month(ask_month & " 1")
'month_track2 keeps track of the month integer value of the current copysheet row
Dim month_track2 As Integer
month_track2 = 0

'Adding columns for more fuel grade options in first report sheet
ws.Range("M:O").EntireColumn.Insert
ws.Range("M6") = "F"
ws.Range("N6") = "G"
ws.Range("O6") = "LG"
'Adding columns for more fuel grade options in second report sheet
If outsider2 <> "" Then
    ws2.Range("M:O").EntireColumn.Insert
    ws2.Range("M6") = "F"
    ws2.Range("N6") = "G"
    ws2.Range("O6") = "LG"
End If
'Array to keep track of F, G, and LG fuel row indices
Dim row_list(2000)
Dim row_list_track As Integer
row_list_track = 0

If outsider1 <> "" Then
    Do While copysheet.Cells(i, 1).Value <> ""
        month_track2 = month(copysheet.Range("I" & row))
        If copysheet.Range("M" & row) <> outsider1 Or month_track <> month_track2 Then
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "OUT" And copysheet.Range("M" & row) = outsider1 And copysheet.Range("H" & row) = "OPENING STOCK" And copysheet.Range("J" & row) = "FF" Then
            holder = copysheet.Range("O" & row)
            ws.Range("D7") = holder
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "IN" And copysheet.Range("M" & row) = outsider1 And copysheet.Range("H" & row) = "OPENING STOCK" And copysheet.Range("J" & row) = "FFF" Then
            holder = copysheet.Range("O" & row)
            ws.Range("E7") = holder * -1
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "OUT" And copysheet.Range("M" & row) = outsider1 Then
            If copysheet.Range("J" & row) = "F" Then
                row_list(row_list_track) = wrow
                wrow = wrow + 1
                row = row + 1
                row_list_track = row_list_track + 1
            ElseIf copysheet.Range("J" & row) = "G" Then
                row_list(row_list_track) = wrow
                wrow = wrow + 1
                row = row + 1
                row_list_track = row_list_track + 1
            ElseIf copysheet.Range("J" & row) = "LG" Then
                row_list(row_list_track) = wrow
                wrow = wrow + 1
                row = row + 1
                row_list_track = row_list_track + 1
            End If
            holder = copysheet.Range("O" & row)
            If copysheet.Range("J" & row) = "FF" Then
                ws.Range("D" & wrow) = holder
                ws.Range("A" & wrow) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws.Range("B" & wrow) = copysheet.Range("H" & row)
                ws.Range("F" & wrow) = copysheet.Range("AE" & row)
                
                wrow = wrow + 1
            End If
            If copysheet.Range("J" & row) = "FFF" Then
                ws.Range("E" & wrow) = holder
                ws.Range("A" & wrow) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws.Range("B" & wrow) = copysheet.Range("H" & row)
                ws.Range("F" & wrow) = copysheet.Range("AE" & row)
                
                wrow = wrow + 1
            End If
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "IN" And copysheet.Range("M" & row) = outsider1 Then
            If copysheet.Range("J" & row) = "F" Then
                row_list(row_list_track) = drow
                drow = drow + 1
                row = row + 1
                row_list_track = row_list_track + 1
            ElseIf copysheet.Range("J" & row) = "G" Then
                row_list(row_list_track) = drow
                drow = drow + 1
                row = row + 1
                row_list_track = row_list_track + 1
            ElseIf copysheet.Range("J" & row) = "LG" Then
                row_list(row_list_track) = drow
                drow = drow + 1
                row = row + 1
                row_list_track = row_list_track + 1
            End If
            holder = copysheet.Range("O" & row)
            If copysheet.Range("J" & row) = "FF" Then
                ws.Range("K" & drow) = holder
                ws.Range("I" & drow) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws.Range("J" & drow) = copysheet.Range("H" & row)
                ws.Range("P" & drow) = copysheet.Range("AE" & row)
                
                drow = drow + 1
            End If
            If copysheet.Range("J" & row) = "FFF" Then
                ws.Range("L" & drow) = holder
                ws.Range("I" & drow) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws.Range("J" & drow) = copysheet.Range("H" & row)
                ws.Range("P" & drow) = copysheet.Range("AE" & row)
                
                drow = drow + 1
            End If
            row = row + 1
        End If
        i = i + 1
    Loop
End If



'RESETING VARIABLES FOR NEXT LOOP
'row keeps track of copysheet row
row = 2
holder = 0
'wrow keeps track of report sheet row of loading
wrow = 8
'drow keeps track of report sheet row of delivery
drow = 8
i = 1
row_list_track = 0
'index counter to use in loop
Dim row_holder As Integer
row_holder = 0

'New Loop to add rows that are of fuel grade F, G, and LG for loading/delivery table
If outsider1 <> "" Then
    ws.Range("F:H").EntireColumn.Insert
    ws.Range("F6") = "F"
    ws.Range("G6") = "G"
    ws.Range("H6") = "LG"
    Do While copysheet.Cells(i, 1).Value <> ""
        month_track2 = month(copysheet.Range("I" & row))
        row_holder = row_list(row_list_track)
        If copysheet.Range("M" & row) <> outsider1 Or month_track <> month_track2 Then
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "OUT" And copysheet.Range("M" & row) = outsider1 And copysheet.Range("H" & row) = "OPENING STOCK" And copysheet.Range("J" & row) = "F" Then
            holder = copysheet.Range("O" & row)
            ws.Range("F7") = holder
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "IN" And copysheet.Range("M" & row) = outsider1 And copysheet.Range("H" & row) = "OPENING STOCK" And copysheet.Range("J" & row) = "G" Then
            holder = copysheet.Range("O" & row)
            ws.Range("G7") = holder * -1
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "OUT" And copysheet.Range("M" & row) = outsider1 And copysheet.Range("H" & row) = "OPENING STOCK" And copysheet.Range("J" & row) = "LG" Then
            holder = copysheet.Range("O" & row)
            ws.Range("H7") = holder
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "OUT" And copysheet.Range("M" & row) = outsider1 Then
            holder = copysheet.Range("O" & row)
            If copysheet.Range("J" & row) = "F" Then
                ws.Range("F" & row_holder) = holder
                ws.Range("A" & row_holder) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws.Range("B" & row_holder) = copysheet.Range("H" & row)
                ws.Range("I" & row_holder) = copysheet.Range("AE" & row)
                
                row_list_track = row_list_track + 1
            End If
            If copysheet.Range("J" & row) = "G" Then
                ws.Range("G" & row_holder) = holder
                ws.Range("A" & row_holder) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws.Range("B" & row_holder) = copysheet.Range("H" & row)
                ws.Range("I" & row_holder) = copysheet.Range("AE" & row)
                
                row_list_track = row_list_track + 1
            End If
            If copysheet.Range("J" & row) = "LG" Then
                ws.Range("H" & row_holder) = holder
                ws.Range("A" & row_holder) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws.Range("B" & row_holder) = copysheet.Range("H" & row)
                ws.Range("I" & row_holder) = copysheet.Range("AE" & row)
                
                row_list_track = row_list_track + 1
            End If
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "IN" And copysheet.Range("M" & row) = outsider1 Then
            holder = copysheet.Range("O" & row)
            If copysheet.Range("J" & row) = "F" Then
                ws.Range("P" & row_holder) = holder
                ws.Range("L" & row_holder) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws.Range("M" & row_holder) = copysheet.Range("H" & row)
                ws.Range("S" & row_holder) = copysheet.Range("AE" & row)
                
                row_list_track = row_list_track + 1
            End If
            If copysheet.Range("J" & row) = "G" Then
                ws.Range("Q" & row_holder) = holder
                ws.Range("L" & row_holder) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws.Range("M" & row_holder) = copysheet.Range("H" & row)
                ws.Range("S" & row_holder) = copysheet.Range("AE" & row)
                
                row_list_track = row_list_track + 1
            End If
            If copysheet.Range("J" & row) = "LG" Then
                ws.Range("R" & row_holder) = holder
                ws.Range("L" & row_holder) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws.Range("M" & row_holder) = copysheet.Range("H" & row)
                ws.Range("S" & row_holder) = copysheet.Range("AE" & row)
                
                row_list_track = row_list_track + 1
            End If
            row = row + 1
        End If
        i = i + 1
    Loop
    'Implementing correct format
    With Range("F108").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Range("G108").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Range("H108").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ws.Range("F105:H105").Borders.LineStyle = xlContinuous
    ws.Range("P100") = "=SUM(P7:P99)"
    ws.Range("Q100") = "=SUM(Q7:Q99)"
    ws.Range("R100") = "=SUM(R7:R99)"
    ws.Range("F100") = "=SUM(F7:F99)"
    ws.Range("G100") = "=SUM(G7:G99)"
    ws.Range("H100") = "=SUM(H7:H99)"
    ws.Range("F105") = "F"
    ws.Range("G105") = "G"
    ws.Range("H105") = "LG"
    ws.Range("F106") = "=F100"
    ws.Range("G106") = "=G100"
    ws.Range("H106") = "=H100"
    ws.Range("F107") = "=P100"
    ws.Range("G107") = "=Q100"
    ws.Range("H107") = "=R100"
    ws.Range("F108") = "=F106-F107"
    ws.Range("G108") = "=G106-G107"
    ws.Range("H108") = "=H106-H107"
End If

'RESETING VARIABLES FOR NEXT LOOP
'row keeps track of copysheet row
row = 2
holder = 0
'wrow keeps track of report sheet row of loading
wrow = 8
'drow keeps track of report sheet row of delivery
drow = 8
i = 1
row_list_track = 0

If outsider2 <> "" Then
    Do While copysheet.Cells(i, 1).Value <> ""
        month_track2 = month(copysheet.Range("I" & row))
        If copysheet.Range("M" & row) <> outsider2 Or (month_track <> month_track2) Then
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "OUT" And copysheet.Range("M" & row) = outsider2 And copysheet.Range("H" & row) = "OPENING STOCK" And copysheet.Range("J" & row) = "FF" Then
            holder = copysheet.Range("O" & row)
            ws2.Range("D7") = holder
            'Cells(i, 4).Value = holder
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "IN" And copysheet.Range("M" & row) = outsider2 And copysheet.Range("H" & row) = "OPENING STOCK" And copysheet.Range("J" & row) = "FFF" Then
            
            holder = copysheet.Range("O" & row)
            
            ws2.Range("E7") = holder * -1
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "OUT" And copysheet.Range("M" & row) = outsider2 Then
            holder = copysheet.Range("O" & row)
            If copysheet.Range("J" & row) = "FF" Then
                ws2.Range("D" & wrow) = holder
                ws2.Range("A" & wrow) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws2.Range("B" & wrow) = copysheet.Range("H" & row)
                ws2.Range("F" & wrow) = copysheet.Range("AE" & row)
                
                wrow = wrow + 1
            End If
            If copysheet.Range("J" & row) = "FFF" Then
                ws2.Range("E" & wrow) = holder
                ws2.Range("A" & wrow) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws2.Range("B" & wrow) = copysheet.Range("H" & row)
                ws2.Range("F" & wrow) = copysheet.Range("AE" & row)
                
                wrow = wrow + 1
            End If
            row = row + 1
        ElseIf (month_track = month_track2) And copysheet.Range("N" & row) = "IN" And copysheet.Range("M" & row) = outsider2 Then
            holder = copysheet.Range("O" & row)
            If copysheet.Range("J" & row) = "FF" Then
                ws2.Range("K" & drow) = holder
                ws2.Range("I" & drow) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws2.Range("J" & drow) = copysheet.Range("H" & row)
                ws2.Range("P" & drow) = copysheet.Range("AE" & row)
                
                drow = drow + 1
            End If
            If copysheet.Range("J" & row) = "FFF" Then
                ws2.Range("L" & drow) = holder
                ws2.Range("I" & drow) = Format(CDate(copysheet.Range("I" & row)), "Medium Date")
                ws2.Range("J" & drow) = copysheet.Range("H" & row)
                ws2.Range("P" & drow) = copysheet.Range("AE" & row)
                
                drow = drow + 1
            End If
            row = row + 1
        End If
        i = i + 1
    Loop
End If

'Hiding any blank Columns or Rows
If outsider2 <> "" Then
    ws2.Range("A31:A99").AutoFilter 1, "<>", , , False
    If IsEmpty(ws2.Range("P8").Value) Then ws2.Range("M1").EntireColumn.Hidden = True
    If IsEmpty(ws2.Range("Q8").Value) Then ws2.Range("N1").EntireColumn.Hidden = True
    If IsEmpty(ws2.Range("R8").Value) Then ws2.Range("O1").EntireColumn.Hidden = True
End If

'Hiding any blank Columns or Rows
If outsider1 <> "" Then
    ws.Range("A31:A99").AutoFilter 1, "<>", , , False
    'If IsEmpty(ws.Range("F8").Value) Then ws.Range("F1").EntireColumn.Hidden = True
    'If IsEmpty(ws.Range("G8").Value) Then ws.Range("G1").EntireColumn.Hidden = True
    'If IsEmpty(ws.Range("H8").Value) Then ws.Range("H1").EntireColumn.Hidden = True
    'If IsEmpty(ws.Range("P8").Value) Then ws.Range("P1").EntireColumn.Hidden = True
    'If IsEmpty(ws.Range("Q8").Value) Then ws.Range("Q1").EntireColumn.Hidden = True
    'If IsEmpty(ws.Range("R8").Value) Then ws.Range("R1").EntireColumn.Hidden = True
End If
ws.Activate

'Copying Worksheet into a new Excel Workbook file
'ws.Copy
'ws2.Copy
Erase row_list
Application.ScreenUpdating = True
End Sub
