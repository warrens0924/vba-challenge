Attribute VB_Name = "Module1"
Sub stockmarket()

    
 'run through all worksheets
 
 For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N2").Value = "Greatest % increase"
        ws.Range("N3").Value = "Greatest % decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        
'declaring variables

Dim TickerName As String

Dim LastRow As Long

Dim LastRowValue As Long
        
Dim TotalTickerVolume As Double
TotalTickerVolume = 0
        
Dim SummaryTableRow As Long
SummaryTableRow = 2

Dim YearlyOpen As Double

Dim YearlyClose As Double
        
Dim YearlyChange As Double
        
Dim PreviousAmount As Long
PreviousAmount = 2
        
Dim PercentChange As Double
        
Dim GreatestIncrease As Double
GreatestIncrease = 0
        
Dim GreatestDecrease As Double
GreatestDecrease = 0
        
Dim GreatestTotalVolume As Double
GreatestTotalVolume = 0

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


        'loop
        
        For i = 2 To LastRow
            TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
            'different ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TickerName = ws.Cells(i, 1).Value
                'print tickername
                ws.Range("I" & SummaryTableRow).Value = TickerName
                'print volume total
                ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
                TotalTickerVolume = 0
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                'print yearly change
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                
                If YearlyOpen = 0 Then
                PercentChange = 0
                
                 Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                    
            End If
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange
                
                'postive gain green
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                    
                'loss red
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                    
                End If
                
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                
                End If
                
            Next i
            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
            
            ' start loop for final results
            For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("P2").Value Then
                    ws.Range("P2").Value = ws.Range("K" & i).Value
                    ws.Range("O2").Value = ws.Range("I" & i).Value
                End If
                If ws.Range("K" & i).Value < ws.Range("P3").Value Then
                    ws.Range("P3").Value = ws.Range("K" & i).Value
                    ws.Range("O3").Value = ws.Range("I" & i).Value
                End If
                If ws.Range("L" & i).Value > ws.Range("P4").Value Then
                    ws.Range("P4").Value = ws.Range("L" & i).Value
                    ws.Range("O4").Value = ws.Range("I" & i).Value
                End If
                
            Next i
            ws.Range("O2").NumberFormat = "0.00%"
            ws.Range("O3").NumberFormat = "0.00%"
            ws.Columns("I:O").AutoFit
        
    Next ws
   
End Sub

