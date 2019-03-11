'Sanjay Mamidi
'Data Viz Class 3/3/2019
'Stock Calculations Assignment - Hard
'Global Variables
Dim GreatestPctIncrease As Double
Dim GreatestPctDecrease As Double
Dim GreatestTotalVol As Double
Dim GreatestPctIncreaseTkr As String
Dim GreatestPctDecreaseTkr As String
Dim GreatestTotVolTkr As String

Public Sub StockVolume()

Dim RowCount As Long
Dim StkVolume As Double
Dim TickerSummaryCurrentRow As Integer
Dim StkOpenPrice As Double
Dim StkClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
'Dim GreatestIncreaseTicker As String
'Dim GreatestDecreaseTicker As String


'Do this code for every worksheet
For Each WS In Worksheets
WS.Activate

RowCount = 0
StkVolume = 0
StkVolume = 0
StkOpenPrice = 0
StkClosePrice = 0
YearlyChange = 0
PercentChange = 0
GreatestPctIncrease = 0
GreatestPctDecrease = 0
GreatestTotalVol = 0
GreatestPctIncreaseTkr = ""
GreatestPctDecreaseTkr = ""
GreatestTotVolTkr = ""
TickerSummaryCurrentRow = 1 'Currently on HeaderRow

'Now set the summary result Column Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest Pct Increase"
Cells(3, 15).Value = "Greatest Pct Decrease"
Cells(4, 15).Value = "Greatest Total Volume"


'First row count of the range to detemine the for loop limits
For Each C In Range("A:A")
    If C.Value <> "" Then
        RowCount = RowCount + 1
    End If
Next C

'Now we can set the for loop for Stock Summary Calculations
For I = 2 To RowCount
    'Save opening price of Stock
    If I = 2 Then
        StkOpenPrice = Cells(I, 3).Value
    End If

    If ((Cells(I, 1).Value <> Cells(I + 1, 1).Value) Or (I = RowCount)) Then
            StkVolume = StkVolume + CDbl(Cells(I, 7).Value)
            Ticker = Cells(I, 1).Value
            Cells(TickerSummaryCurrentRow + 1, 9).Value = Cells(I, 1).Value ' for Tickername
            
            'Price Change Calculations
            StkClosePrice = Cells(I, 6).Value
            YearlyChange = YearChangeCalc(StkClosePrice, StkOpenPrice)
           
            Cells(TickerSummaryCurrentRow + 1, 10).Value = YearlyChange
            If YearlyChange > 0 Then
                Cells(TickerSummaryCurrentRow + 1, 10).Interior.ColorIndex = 4
            Else
                Cells(TickerSummaryCurrentRow + 1, 10).Interior.ColorIndex = 3
            End If
            
            'Percent Change Calculations
            PercentChange = PercentChangeCalc(StkClosePrice, StkOpenPrice)
            Cells(TickerSummaryCurrentRow + 1, 11).Value = CStr(PercentChange) & "%"
            
            If PercentChange > 0 Then
                Cells(TickerSummaryCurrentRow + 1, 11).Interior.ColorIndex = 4
            Else
                Cells(TickerSummaryCurrentRow + 1, 11).Interior.ColorIndex = 3
            End If
            Cells(TickerSummaryCurrentRow + 1, 12).Value = StkVolume ' for TickerTotal Volume
            
            'GreatestCalculations for last row ie 12/30/YYYY for this ticker
                 'GreatestCalculations
            If StkOpenPrice <> 0 And StkClosePrice <> 0 Then  'To take care of Divide by Zero Errors
                Call GreatestPctCalc(PercentChange, Ticker)
            End If
           
            If GreatestTotalVol < StkVolume Then
                GreatestTotalVol = StkVolume
                GreatestTotVolTkr = Ticker
            End If
            
            'Set these counters correctly for next Ticker Symbol
            TickerSummaryCurrentRow = TickerSummaryCurrentRow + 1
            StkVolume = 0
            StkOpenPrice = Cells(I + 1, 3).Value
    Else
            StkVolume = StkVolume + CDbl(Cells(I, 7).Value)
            Ticker = Cells(I, 1).Value
    End If
Next I
    'Update the Summary Location for this WS with Ticker Values
    Cells(2, 16).Value = GreatestPctIncreaseTkr
    Cells(2, 17).Value = GreatestPctIncrease
    Cells(3, 16).Value = GreatestPctDecreaseTkr
    Cells(3, 17).Value = GreatestPctDecrease
    Cells(4, 16).Value = GreatestTotVolTkr
    Cells(4, 17).Value = GreatestTotalVol
    


Next WS
End Sub

Function YearChangeCalc(ByVal ClosePrice As Double, OpenPrice As Double) As Double
        
        YearChangeCalc = ClosePrice - OpenPrice
    
End Function

Function PercentChangeCalc(ByVal ClosePrice As Double, ByVal OpenPrice As Double) As Double
        If ClosePrice = OpenPrice Then
             PercentChange = 0 'do nothing
         Else
             
        PercentChangeCalc = Round(((ClosePrice - OpenPrice) / OpenPrice) * 100, 2)
        End If
End Function

Sub GreatestPctCalc(ByVal PercentChange, ByVal Stock As String)
          
            If PercentChange > 0 Then 'Increase
                If PercentChange > GreatestPctIncrease Then
                    GreatestPctIncrease = PercentChange
                    GreatestPctIncreaseTkr = Stock
                End If
            End If
            If PercentChange < 0 Then 'Decrease
                If PercentChange < GreatestPctDecrease Then
                    GreatestPctDecrease = PercentChange
                    GreatestPctDecreaseTkr = Stock
                End If
            End If
        
        
End Sub

    

    







