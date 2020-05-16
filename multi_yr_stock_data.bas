Attribute VB_Name = "Module1"
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

For Each ws In Worksheets
ws.Activate
   Summary_Table_Row = 2
    start = 2
'Label Columns
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
' Loop through all stock transactions
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        start = 2
            For I = 2 To LastRow
            
' Check if we are still within the same Ticker_Sign, if it is not...
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

' Set the StockSymbol
    Ticker_Sign = ws.Cells(I, 1).Value

' Calculate YearlyChange
    YearlyChange = ws.Cells(I, 6) - ws.Cells(start, 3)

' Calculate TotalPercentage
On Error Resume Next
    TotalPercentage = (ws.Cells(I, 6) / ws.Cells(start, 3) - 1)

'Calculate Volume
    VolumeTotal = VolumeTotal + ws.Cells(I, 7).Value

' Print to the Summary Table
    ws.Range("I" & Summary_Table_Row).Value = Ticker_Sign
    ws.Range("J" & Summary_Table_Row).Value = YearlyChange
    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    ws.Range("K" & Summary_Table_Row).Value = TotalPercentage
    ws.Range("L" & Summary_Table_Row).Value = VolumeTotal

' Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
      
' Reset
    VolumeTotal = 0
    YearlyChange = 0
    TotalPercentage = 0
    start = I + 1
    
' If the cell immediately following a row is the stock...
Else
    VolumeTotal = VolumeTotal + ws.Cells(I, 7).Value
End If

' Color Cells
    If ws.Cells(I, 10) > 0 Then
        ws.Cells(I, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(I, 10) <= 0 Then
        ws.Cells(I, 10).Interior.ColorIndex = 3
    End If
    
        Next I
    Next ws
End Sub

