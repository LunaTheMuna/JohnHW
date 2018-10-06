    
    
Sub stock_volumes()

Dim ticker As String
Dim prevTicker As String
Dim SummaryTableRow As Integer
Dim i As Integer




Dim tickerAnnualVolume As Double
Dim beginPrice As Double
Dim endPrice As Double
Dim priceChange As Double
Dim priceChangeColor As Integer
Dim percentChange As Double
Dim gPercentIncrease As Double
Dim gPercentDecrease As Double
Dim gPercentIncreaseTicker As String
Dim gPercentDecreaseTicker As String
Dim gTotalVolume As String
Dim gTotalVolumeTicker As String


gPercentIncrease = 0
gPercentDecrease = 0
gTotalVolume = 0

tickerAnnualVolume = 0
i = 2
ticker = Cells(i, 1).Value
prevTicker = ticker


SummaryTableRow = 2

Range("I1:L70000").Clear

Range("I" & 1).Value = "Ticker"
Range("J" & 1).Value = "Total Stock Volume"
Range("K" & 1).Value = "Yearly Change"
Range("L" & 1).Value = "Percent Change"
beginPrice = Cells(i, 6).Value
While ticker <> "" And i < 30000

    If ticker <> prevTicker Then
        endPrice = Cells(i - 1, 6).Value
        priceChange = endPrice - beginPrice
        If priceChange < 0 Then
            priceChangeColor = 3
        Else: priceChangeColor = 4
        End If
        percentChange = (endPrice - beginPrice) / beginPrice
        If percentChange > gPercentIncrease Then
            gPercentIncrease = percentChange
            gPercentIncreaseTicker = prevTicker
        ElseIf percentChange < gPercentDecrease Then
            gPercentDecrease = percentChange
            gPercentDecreaseTicker = prevTicker
        End If
        
        If tickerAnnualVolume > gTotalVolume Then
            gTotalVolume = tickerAnnualVolume
            gTotalVolumeTicker = prevTicker
        End If
        
        
        beginPrice = Cells(i, 6).Value
        Range("I" & SummaryTableRow).Value = prevTicker
        Range("J" & SummaryTableRow).Value = tickerAnnualVolume
        Range("K" & SummaryTableRow).Value = priceChange
        Range("K" & SummaryTableRow).Interior.ColorIndex = priceChangeColor
        Range("L" & SummaryTableRow).Value = percentChange
        Range("L" & SummaryTableRow).NumberFormat = "0%"
        
        
        
        tickerAnnualVolume = Cells(i, 7)

        SummaryTableRow = SummaryTableRow + 1
    Else
      
        tickerAnnualVolume = tickerAnnualVolume + Cells(i, 7).Value
    End If
    
   i = i + 1
    prevTicker = ticker
    ticker = Cells(i, 1).Value
    
Wend

' print greatest increase,decrease,total volume

End Sub
    

