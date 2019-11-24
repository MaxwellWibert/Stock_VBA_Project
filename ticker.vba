Private Function TickerVal(index As Integer, WS As Worksheet) As String
    TickerVal = WS.Cells(index, 1).Value
End Function


Private Function nextIndex(currentIndex As Integer, WS As Worksheet) As Integer
    Dim i As Integer
    i = currentIndex
    Dim currentTick As String
    currentTick = WS.Cells(currentIndex, 1).Value
    While WS.Cells(i, 1).Value = currentTick
        i = i + 1
    Wend
    nextIndex = i
End Function

Private Function Volume(fst As Integer, lst As Integer, WS As Worksheet) As Integer
    Dim total As Integer
    total = 0
    For i = fst To lst
        total = total + WS.Cells(7, i)
    Next i
    Volume = total
End Function

Private Function ArgMax(fst As Integer, lst As Integer, col As Integer, WS As Worksheet) As Integer
    Dim runningArgMax As Integer
    runningArgMax = fst
    Dim runningMax As Double
    runningMax = WS.Cells(fst, col).Value
    For i = fst To lst
        If WS.Cells(i, col).Value > runningMax Then
            runningMax = WS.Cells(i, col)
            runningArgMin = i
        End If
    Next i
    ArgMax = runningArgMax
End Function

Private Function ArgMin(fst As Integer, lst As Integer, col As Integer, WS As Worksheet) As Integer
    Dim runningArgMin As Integer
    runningArgMin = fst
    Dim runningMin As Double
    runningMin = WS.Cells(fst, col).Value
    For i = fst To lst
        If WS.Cells(i, col) < runningMin Then
            runningMin = WS.Cells(i, col)
            runningArgMin = i
        End If
    Next i
    ArgMin = runningArgMin
End Function

Private Function Delta(fst As Integer, lst As Integer, WS As Worksheet) As Double
    Delta = WS.Cells(lst, 6).Value - WS.Cells(fst, 3).Value
End Function

Private Function PercentDelta(fst As Integer, lst As Integer, WS As Worksheet) As Double
    PercentDelta = 100 * (WS.Cells(lst, 6).Value - WS.Cells(fst, 3).Value) / WS.Cells(fst, 3).Value
End Function

Public Sub StockAnalysis()
    Dim WS As Worksheet
    Dim Ticker As String

    Dim Vol As Integer

    Dim Summary_Table_Row As Integer

    Dim yearlyChange As Double

    Dim percentChange As Double

    Dim Tick_Begin As Integer
    Dim Tick_End As Integer

    Dim changeMindex As Integer
    Dim changeMaxdex As Integer

    Dim volMaxdex As Integer

    For Each WS In Worksheets
        Vol = 0
        WS.Cells(1, 10).Value = "Ticker"
        WS.Cells(1, 11).Value = "Yearly Change"
        WS.Cells(1, 12).Value = "Percent Change"
        WS.Cells(1, 13).Value = "Total Stock Volume"

        Tick_Begin = 2
        Tick_End = nextIndex(2, WS) - 1
        Summary_Table_Row = 2

        While Not IsEmpty(WS.Cells(Tick_Begin, 1))
            Ticker = TickerVal(Tick_Begin, WS)
            Vol = Volume(Tick_Begin, Tick_End, WS)
            yearlyChange = Delta(Tick_Begin, Tick_End, WS)
            percentChange = PercentDelta(Tick_Begin, Tick_End, WS)
            
            WS.Cells(Summary_Table_Row, 10).Value = Ticker
            WS.Cells(Summary_Table_Row, 11).Value = yearlyChange
            WS.Cells(Summary_Table_Row, 12).Value = percentChange
            WS.Cells(Summary_Table_Row, 13).Value = Vol
            
            If (yearlyChange > 0) Then
                WS.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
            ElseIf (yearlyChange < 0) Then
                WS.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
            End If

            Tick_Begin = Tick_End + 1
            Tick_End = nextIndex(Tick_Begin, WS) - 1
            Summary_Table_Row = Summary_Table_Row + 1
        Wend

        Summary_Table_Row = Summary_Table_Row - 1

        WS.Cells(2, 15).Value = "Greatest Percent Increase"
        WS.Cells(3, 15).Value = "Greatest Percent Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"

        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"


        changeMindex = ArgMin(2, Summary_Table_Row, 12, WS)
        changeMaxdex = ArgMax(2, Summary_Table_Row, 12, WS)

        WS.Cells(2, 16).Value = WS.Cells(changeMaxdex, 10).Value
        WS.Cells(2, 17).Value = WS.Cells(changeMaxdex, 12).Value

        WS.Cells(3, 16).Value = WS.Cells(changeMindex, 10).Value
        WS.Cells(3, 17).Value = WS.Cells(changeMindex, 12).Value

        volMaxdex = ArgMax(2, Summary_Table_Row, 13, WS)
        WS.Cells(4, 16).Value = WS.Cells(volMaxdex, 10).Value
        WS.Cells(4, 17).Value = WS.Cells(volMaxdex, 13).Value
        
    Next WS
End Sub
