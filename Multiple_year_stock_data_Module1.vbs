'Attribute VB_Name = "Module1"
Sub StockMarketSummary()

Dim TickerCode As String
Dim TotalVolume As Double
Dim SummaryCount As Integer
Dim YearlyChange As Double
Dim PercentChange As Double
Dim OpeningPrice() As Double
Dim ClosingPrice() As Double

Dim GreatestIncrTicker As String
Dim GreatestIncrVal As Double
Dim GreatestDecrTicker As String
Dim GreatestDecrVal As Double
Dim GreatestVolTicker As String
Dim GreatestVolVal As Double

' Go through each worksheet and apply the same changes
'
For Each ws In Worksheets
    
    GreatestIncrVal = 0
    GreatestDecrVal = 0
    GreatestVolVal = 0
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ReDim OpeningPrice(1 To lastrow) As Double
    ReDim ClosingPrice(1 To lastrow) As Double
    
    SummaryCount = 2
    TotalVolume = 0
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest% Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    OpeningPrice(2) = ws.Cells(2, 3)
    ClosingPrice(2) = 0
    
    For i = 2 To lastrow
        TickerCode = ws.Cells(i, 1)
        If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
            TotalVolume = TotalVolume + ws.Cells(i, 7)
        Else
            ClosingPrice(SummaryCount) = ws.Cells(i, 6)
            TotalVolume = TotalVolume + ws.Cells(i, 7)
            ws.Range("I" & SummaryCount).Value = TickerCode
            ws.Range("L" & SummaryCount).Value = TotalVolume
            If i <> lastrow Then
                OpeningPrice(SummaryCount + 1) = ws.Cells(i + 1, 3)
                SummaryCount = SummaryCount + 1
                TotalVolume = 0
            End If
        End If
    
    Next i
    
    
    For i = 2 To SummaryCount
        YearlyChange = (ClosingPrice(i) - OpeningPrice(i))
        ws.Range("J" & i).Value = YearlyChange
        If YearlyChange < 0 Then
            ws.Range("J" & i).Interior.ColorIndex = 3
        Else
            ws.Range("J" & i).Interior.ColorIndex = 4
        End If
        
        If OpeningPrice(i) > 0 Then
            PercentChange = ((ClosingPrice(i) - OpeningPrice(i)) / OpeningPrice(i))
        Else
            PercentChange = 0
        End If
        
        ws.Range("K" & i).Value = PercentChange
        ws.Range("K" & i).NumberFormat = "0.00%"
        
        ' Test for the Greatest Statistics
        If PercentChange > GreatestIncrVal Then
           GreatestIncrVal = PercentChange
            GreatestIncrTicker = ws.Range("I" & i)
        End If
        If PercentChange < GreatestDecrVal Then
            GreatestDecrVal = PercentChange
            GreatestDecrTicker = ws.Range("I" & i)
        End If
        If ws.Range("L" & i) > GreatestVolVal Then
            GreatestVolVal = ws.Range("L" & i)
            GreatestVolTicker = ws.Range("I" & i)
       End If
    
    Next i
    
    ws.Range("P2").Value = GreatestIncrTicker
    ws.Range("Q2").Value = GreatestIncrVal
    ws.Range("Q2").NumberFormat = "0.00%"
    
    ws.Range("P3").Value = GreatestDecrTicker
    ws.Range("Q3").Value = GreatestDecrVal
    ws.Range("Q3").NumberFormat = "0.00%"
    
    ws.Range("P4").Value = GreatestVolTicker
    ws.Range("Q4").Value = GreatestVolVal
    
    ws.Columns.AutoFit

Next ws

End Sub



