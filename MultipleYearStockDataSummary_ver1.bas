Attribute VB_Name = "Module11"
' ================================================================================
' Module Name: Module 2 Challenge
' Description: This module provides functionality to create a data summary of
'              the ticker symbol, yearly change from opening price at beginning
'              of a given year to the closing price at the end of that year.
'              It also computes corresponding percentage change from opening price
'              and closing price, and the total stock volume.
'
' Created By: Michael Jardinico
' Created On: Sept 17, 2023
' Version: 1.0
' ================================================================================


Sub ClearResultValues()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ws.Range("I:R").Clear
    
End Sub

Sub GetResult()

    Dim lastRow As Long
    Dim ws As Worksheet
    Dim FirstValue As Double, LastValue As Double, ResultValue As Double
    Dim total As Double
    'Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer
    Dim i As Long, j As Long, k As Long, m As Long, n As Long
    Dim CurrentTicker As String
    Dim ticker As String
    Dim IsFirstValue As Boolean
    Dim lastTickerRow As Long
    Dim CurrentValue As String
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxTotalVolume As Double
    Dim MaxIncreaseTikcer As String
    Dim MaxDecreaseTicker As String
    Dim MaxTotalVolumeTicker As String
        

    'Set ws as the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    'Find the last row of data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Clear the destination column first to avoid any conflicts
    ws.Range("I:I").ClearContents
    
    
    
    '* * * * * * * * * * * * * * * * * * * * *
    '*     Display Unique Ticker Symbols     *
    '* * * * * * * * * * * * * * * * * * * * *
    
    'Use AdvanceFilter to get the Ticker symbol and display on column I
    ws.Range("A1:A" & lastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I1:I" & lastRow), Unique:=True
    
    'Display the header name Ticker on column I
    Range("I1").Value = "Ticker"
    
    
    
    '* * * * * * * * * * * * * * * * * * * * * * * * * *
    '*  Compute for Yearly Change and Percent Change   *
    '* * * * * * * * * * * * * * * * * * * * * * * * * *
    
    'Display the header name "Yearly Change" on column J
    Range("J1").Value = "Yearly Change"
    
    'Display the header name "Percent Change" on column K
    Range("K1").Value = "Percent Change"
    
    
    'This will loop through the ticker in column I
    For i = 1 To lastRow
        ' Check if the cell in column 8 has a value and not blank
        If ws.Cells(i + 1, 9).Value <> "" Then
            CurrentTicker = ws.Cells(i + 1, "I").Value
            IsFirstValue = True
            
            'Now, this will loop through the column A to match with column I value
            For j = 1 To lastRow
                If ws.Cells(j + 1, "A").Value = CurrentTicker Then
                    If IsFirstValue Then
                        FirstValue = ws.Cells(j + 1, "C").Value
                        IsFirstValue = False
                    End If
                    LastValue = ws.Cells(j + 1, "F").Value
                End If
            Next j
            
            'This will populate the value in column I
            ws.Cells(i + 1, "J").Value = LastValue - FirstValue
            
            'Change the cell color of resulting value in column J to red if
            'value is less than or equal to 0, while green if greater than 0
            If ws.Cells(i + 1, "J").Value <= 0 Then
                ws.Cells(i + 1, "J").Interior.Color = RGB(255, 0, 0)
            ElseIf ws.Cells(i + 1, "J").Value > 0 Then
                ws.Cells(i + 1, "J").Interior.Color = RGB(0, 255, 0)
            End If
            
            
            'Assign variable as ResultValue to difference between LastValue and FirstValue
            ResultValue = ws.Cells(i + 1, "J").Value
            
            'Compute for the Yearly Change and display on column K. Use the equation (LastValue / ResultValue) x 100%
            ws.Cells(i + 1, "K").Value = Round((ResultValue / FirstValue) * 100, 2) & "%"
            
         End If
    Next i
    
    
    
    '* * * * * * * * * * * * * * * * * * * * *
    '*  Compute for the Total Stock Volume   *
    '* * * * * * * * * * * * * * * * * * * * *
    
    'Display the header name "Total Stock Volume"
    Range("L1").Value = "Total Stock Volume"
    
    'Determine the last row of the Ticker column
    lastTickerRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'Loop through the unique value in Ticker column
    For k = 2 To lastTickerRow
        ticker = ws.Cells(k, "I")
        
        'Assign an initial value to sum of common ticker called total
        total = 0
        
        'skip if cell is empty
        If ws.Cells(k, "I").Value <> "" Then
            ticker = ws.Cells(k, "I").Value
        Else
            GoTo NextIteration     'This will go to NextIteration line:
        End If
        
        'Loop through the main table to accumulate the sum of each Ticker symbol
        For m = 2 To lastRow
            CurrentValue = ws.Cells(m, "A").Value
            If CurrentValue = ticker Then
                total = total + ws.Cells(m, "G").Value
            End If
        
        Next m
        
        'Display the total in column L
        ws.Cells(k, "L").Value = total
        
NextIteration:
    Next k
    
    
    
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    '*  Greatest % Increase and % Decrease, Greatest Total Volume  *
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    'Display Names Headers
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Initialize values for MaxIncrease, MaxDecrease, MaxTotalVolume
    MaxIncrease = ws.Cells(2, "K").Value
    MaxDecrease = ws.Cells(2, "K").Value
    MaxTotalVolume = ws.Cells(2, "L").Value
    MaxIncreaseTicker = ws.Cells(2, "I").Value
    MaxDecreaseTicker = ws.Cells(2, "I").Value
    MaxTotalVolumeTicker = ws.Cells(2, "I").Value
    
    
    For n = 2 To lastTickerRow
        If ws.Cells(n, "K").Value > MaxIncrease Then
            MaxIncrease = ws.Cells(n, "K").Value
            MaxIncreaseTicker = ws.Cells(n, "I").Value
            
        ElseIf ws.Cells(n, "K").Value < MaxDecrease Then
            MaxDecrease = ws.Cells(n, "K").Value
            MaxDecreaseTicker = ws.Cells(n, "I").Value
        
        End If
        
        If ws.Cells(n, "L").Value > MaxTotalVolume Then
            MaxTotalVolume = ws.Cells(n, "L").Value
            MaxTotalVolumeTicker = ws.Cells(n, "I").Value
        End If
        
    Next n
            
    'Display the ticker symbols in column P
    ws.Range("P2").Value = MaxIncreaseTicker
    ws.Range("P3").Value = MaxDecreaseTicker
    ws.Range("P4").Value = MaxTotalVolumeTicker
    
    'Display the Ticker Values in colomn Q
    ws.Range("Q2").Value = (MaxIncrease * 100) & "%"
    ws.Range("Q3").Value = (MaxDecrease * 100) & "%"
    ws.Range("Q4").Value = MaxTotalVolume
    
    
End Sub
