
Sub StockTracker()

' Initialize Variables
    Dim Ticker_Sign As String
    Dim YearlyChange As Double
        YearlyChange = 0
    Dim Summary_Table_Row As Double
        Summary_Table_Row = 2
    Dim VolumeTotal As Double
        VolumeTotal = 0
    Dim TotalPercentage As Double
        TotalPercentage = 0
    Dim start As Double
    Dim ws As Worksheet
    Dim MaxVol As Double
    Dim MaxTic As String
    Dim GreatPerc As Double
    Dim MaxTic2 As String
    Dim GPD As Double
    
For Each ws In Worksheets
ws.Activate
   Summary_Table_Row = 2
    start = 2
'Label Columns
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest Total Volume"
    

    
' Loop through all stock transactions
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        start = 2
            For i = 2 To LastRow
            
' Check if we are still within the same Ticker_Sign, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Set the StockSymbol
    Ticker_Sign = ws.Cells(i, 1).Value

' Calculate YearlyChange
    YearlyChange = ws.Cells(i, 6) - ws.Cells(start, 3)

' Calculate TotalPercentage
    TotalPercentage = (ws.Cells(i, 6) / ws.Cells(start, 3) - 1)
    
    GreatPerc = Application.WorksheetFunction.max(Range("K:K"))
    GPD = Application.WorksheetFunction.Min(Range("K:K"))
'MaxTic2 = ws.Range("K:K").Find(MaxVol).Offset(, -2)...why does this not work...?
        
'Calculate Volume
    VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
    
    MaxVol = Application.WorksheetFunction.max(Range("L:L"))
'...When this does...?
    MaxTic = ws.Range("L:L").Find(MaxVol).Offset(, -3)


' Print to the Summary Table
    ws.Range("I" & Summary_Table_Row).Value = Ticker_Sign
    ws.Range("J" & Summary_Table_Row).Value = YearlyChange
    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    ws.Range("K" & Summary_Table_Row).Value = TotalPercentage
    ws.Range("L" & Summary_Table_Row).Value = VolumeTotal
    ws.Range("P4") = MaxVol
    ws.Range("O4") = MaxTic
    ws.Range("P2") = GreatPerc
    ws.Range("P2").NumberFormat = "0.00%"
    'ws.Range("O2") = MaxTic2
    ws.Range("P3") = GPD
    ws.Range("P3").NumberFormat = "0.00%"
    
' Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
    

      
' Reset
    VolumeTotal = 0
    YearlyChange = 0
    TotalPercentage = 0
    start = i + 1
    

    
' If the cell immediately following a row is the stock...
Else
    VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value

End If



' Color Cells
    If ws.Cells(i, 10) > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 10) <= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
    
        Next i
    Next ws
End Sub
