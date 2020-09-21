Attribute VB_Name = "Module1"


Sub stock_n_roll()

'define needed vaiables
    Dim LastRow As Long
    Dim LastCond As Integer
    Dim Ticker As String
    Dim OpenValue As Double
    Dim EndValue As Double
    Dim TotalVol As Variant
    Dim YrChng As Double
    Dim PerChng As Double
    Dim cout1 As Integer
    Dim coutvol As Long
    Dim ws As Worksheet
    
For Each ws In Worksheets

'populate the sheet with cell headers
    'first table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    'second table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

'Set the LastRow
    LastRow = (ws.Cells(Rows.Count, 1).End(xlUp).Row)

'Start OpenValue and counter off on the right cell
    OpenValue = ws.Cells(2, 3).Value
    coutvol = 0

'filling in the values of the first table
    For i = 2 To (LastRow - 1)
        'create a counter for the sum column
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            coutvol = coutvol + 1
    
        Else
            'adjust the EndValue for the iteration
            EndValue = ws.Cells(i, 6).Value
            
            'Get the values
            Ticker = ws.Cells(i, 1).Value
            YrChng = (EndValue - OpenValue)
            'make sure there isnt division by zero
                If OpenValue <> 0 Then
                    PerChng = ((EndValue - OpenValue) / OpenValue)
                Else
                    PerChng = 0
                End If
            TotalVol = Excel.WorksheetFunction.Sum(ws.Range(ws.Cells((i - coutvol), 7), ws.Cells(i, 7)))
            
            'place the values
            ws.Cells(i, 9).Value = Ticker
            ws.Cells(i, 10).Value = YrChng
            ws.Cells(i, 11).Value = PerChng
            ws.Cells(i, 12).Value = TotalVol
            
            'addjust the OpenValue and coutnter for the next iteraton
            OpenValue = ws.Cells(i + 1, 3).Value
            coutvol = 0
            
       End If
    Next i
       
'format table 1
    'remove the empty cells for the Yearly Change Column
        'set the counter
            cout1 = 1
        'get the values without blanks
            For i = 2 To LastRow
                If ws.Cells(i, 9).Value <> "" Then
                    ws.Cells(cout1 + 1, 13).Value = ws.Cells(i, 9).Value
                    cout1 = cout1 + 1
                End If
            Next i
        'clear out the cells and move them to the correct location
            ws.Range(ws.Cells(2, 9), ws.Cells(LastRow, 9)).Value = ""
            ws.Range(ws.Cells(2, 9), ws.Cells(LastRow, 9)).Value = ws.Range(ws.Cells(2, 13), ws.Cells(LastRow, 13)).Value
            ws.Range(ws.Cells(2, 13), ws.Cells(LastRow, 13)).Value = ""
            
    'remove the empty cells for the Yearly Change Column
        'set the counter
            cout1 = 1
        'get the values without blanks
            For i = 2 To LastRow
                If ws.Cells(i, 10).Value <> "" Then
                    ws.Cells(cout1 + 1, 13).Value = ws.Cells(i, 10).Value
                    cout1 = cout1 + 1
                End If
            Next i
        'clear out the cells and move them to the correct location
            ws.Range(ws.Cells(2, 10), ws.Cells(LastRow, 10)).Value = ""
            ws.Range(ws.Cells(2, 10), ws.Cells(LastRow, 10)).Value = ws.Range(ws.Cells(2, 13), ws.Cells(LastRow, 13)).Value
            ws.Range(ws.Cells(2, 13), ws.Cells(LastRow, 13)).Value = ""
    
    'remove the empty cells for the Percent Change Column
        'set the counter
            cout1 = 1
        'get the values without blanks
            For i = 2 To LastRow
                If ws.Cells(i, 11).Value <> "" Then
                    ws.Cells(cout1 + 1, 13).Value = ws.Cells(i, 11).Value
                    cout1 = cout1 + 1
                End If
            Next i
        'clear out the cells and move them to the correct location
            ws.Range(ws.Cells(2, 11), ws.Cells(LastRow, 11)).Value = ""
            ws.Range(ws.Cells(2, 11), ws.Cells(LastRow, 11)).Value = ws.Range(ws.Cells(2, 13), ws.Cells(LastRow, 13)).Value
            ws.Range(ws.Cells(2, 13), ws.Cells(LastRow, 13)).Value = ""
    
    'remove the empty  cells for the Total Stock Volume Column
        'reset the counter
            cout1 = 1
        'get the values without blanks
            For i = 2 To LastRow
                If ws.Cells(i, 12).Value <> "" Then
                    ws.Cells(cout1 + 1, 13).Value = ws.Cells(i, 12).Value
                    cout1 = cout1 + 1
                End If
            Next i
        'clear out the cells and move them to the correct location
            ws.Range(ws.Cells(2, 12), ws.Cells(LastRow, 12)).Value = ""
            ws.Range(ws.Cells(2, 12), ws.Cells(LastRow, 12)).Value = ws.Range(ws.Cells(2, 13), ws.Cells(LastRow, 13)).Value
            ws.Range(ws.Cells(2, 13), ws.Cells(LastRow, 13)).Value = ""
            
'conditionally format the percent change column
    'set LastCond
        LastCond = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    'the conditionals
        For i = 2 To LastCond
            If ws.Cells(i, 10).Value >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
    
'change cell widths to accomodate values
    ws.Range("A1:Q1").EntireColumn.AutoFit

'format the percent change column
    ws.Range("k1").EntireColumn.NumberFormat = "0.00%"

'fill in the values for the second table
    'assign the values to the needed cells
    ws.Cells(2, 17).Value = Excel.WorksheetFunction.Max(ws.Range(ws.Cells(2, 11), ws.Cells(LastCond, 11)))
    ws.Cells(3, 17).Value = Excel.WorksheetFunction.Min(ws.Range(ws.Cells(2, 11), ws.Cells(LastCond, 11)))
    ws.Cells(4, 17).Value = Excel.WorksheetFunction.Max(ws.Range(ws.Cells(2, 12), ws.Cells(LastCond, 12)))
    'assign the ticker
    For i = 2 To LastCond
        If ws.Cells(i, 11).Value = ws.Cells(2, 17) Then
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 11).Value = ws.Cells(3, 17).Value Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        ElseIf Cells(i, 12).Value = Cells(4, 17).Value Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        End If
    Next i
    'format the second table
    ws.Range("q2:q3").NumberFormat = "0.00%"

Next ws

End Sub

