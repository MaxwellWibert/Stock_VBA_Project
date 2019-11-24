Private Function TickerVal(index As Long, WS As Worksheet) As String
    TickerVal = WS.Cells(index, 1).value
End Function


Private Function nextIndex(currentIndex As Long, WS As Worksheet) As Long
    Dim i As Long
    i = currentIndex
    Dim currentTick As String
    currentTick = WS.Cells(currentIndex, 1).value
    While WS.Cells(i, 1).value = currentTick
        i = i + 1
    Wend
    nextIndex = i
End Function

Private Function Volume(fst As Long, lst As Long, WS As Worksheet) As Long
    Dim total As Long
    total = 0
    For i = fst To lst
        total = total + CLng(WS.Cells(7, i).value)
    Next i
    Volume = total
End Function

Private Function ArgMax(fst As Long, lst As Long, col As Long, WS As Worksheet) As Long
    Dim runningArgMax As Long
    runningArgMax = fst
    Dim runningMax As Double
    runningMax = WS.Cells(fst, col).value
    For i = fst To lst
        If WS.Cells(i, col).value > runningMax Then
            runningMax = WS.Cells(i, col)
            runningArgMin = i
        End If
    Next i
    ArgMax = runningArgMax
End Function

Private Function ArgMin(fst As Long, lst As Long, col As Long, WS As Worksheet) As Long
    Dim runningArgMin As Long
    runningArgMin = fst
    Dim runningMin As Double
    runningMin = WS.Cells(fst, col).value
    For i = fst To lst
        If WS.Cells(i, col) < runningMin Then
            runningMin = WS.Cells(i, col)
            runningArgMin = i
        End If
    Next i
    ArgMin = runningArgMin
End Function

Private Function Delta(fst As Long, lst As Long, WS As Worksheet) As Double
    Delta = WS.Cells(lst, 6).value - WS.Cells(fst, 3).value
End Function

Private Function PercentDelta(fst As Long, lst As Long, WS As Worksheet) As Double
    PercentDelta = (WS.Cells(lst, 6).value - WS.Cells(fst, 3).value) / WS.Cells(fst, 3).value
End Function

Public Sub StockAnalysis()
    Dim WS As Worksheet
    Dim Ticker As String

    Dim vol As Long

    Dim Summary_Table_Row As Long

    Dim yearlyChange As Double

    Dim percentChange As Double

    Dim Tick_Begin As Long
    Dim Tick_End As Long

    Dim changeMindex As Long
    Dim changeMaxdex As Long

    Dim volMaxdex As Long

    For Each WS In Worksheets
        vol = 0
        WS.Cells(1, 10).value = "Ticker"
        WS.Cells(1, 11).value = "Yearly Change"
        WS.Cells(1, 12).value = "Percent Change"
        WS.Cells(1, 13).value = "Total Stock Volume"

        Tick_Begin = 2
        Tick_End = nextIndex(2, WS) - 1
        Summary_Table_Row = 2

        While Not IsEmpty(WS.Cells(Tick_Begin, 1))
            Ticker = TickerVal(Tick_Begin, WS)
            vol = Volume(Tick_Begin, Tick_End, WS)
            yearlyChange = Delta(Tick_Begin, Tick_End, WS)
            percentChange = PercentDelta(Tick_Begin, Tick_End, WS)
            
            WS.Cells(Summary_Table_Row, 10).value = Ticker
            WS.Cells(Summary_Table_Row, 11).value = yearlyChange
            WS.Cells(Summary_Table_Row, 12).value = percentChange
            WS.Cells(Summary_Table_Row, 13).value = vol
            
            If (yearlyChange > 0) Then
                WS.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
            ElseIf (yearlyChange < 0) Then
                WS.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
            End If
            
            WS.Cells(Summary_Table_Row, 12).NumberFormat = "0.00%"

            Tick_Begin = Tick_End + 1
            Tick_End = nextIndex(Tick_Begin, WS) - 1
            Summary_Table_Row = Summary_Table_Row + 1
        Wend

        Summary_Table_Row = Summary_Table_Row - 1

        WS.Cells(2, 15).value = "Greatest Percent Increase"
        WS.Cells(3, 15).value = "Greatest Percent Decrease"
        WS.Cells(4, 15).value = "Greatest Total Volume"

        WS.Cells(1, 16).value = "Ticker"
        WS.Cells(1, 17).value = "Value"


        changeMindex = ArgMin(2, Summary_Table_Row, 12, WS)
        changeMaxdex = ArgMax(2, Summary_Table_Row, 12, WS)

        WS.Cells(2, 16).value = WS.Cells(changeMaxdex, 10).value
        WS.Cells(2, 17).value = WS.Cells(changeMaxdex, 12).value
        WS.Cells(2, 17).NumberFormat = "0.00%"

        WS.Cells(3, 16).value = WS.Cells(changeMindex, 10).value
        WS.Cells(3, 17).value = WS.Cells(changeMindex, 12).value
        WS.Cells(3, 17).NumberFormat = "0.00%"

        volMaxdex = ArgMax(2, Summary_Table_Row, 13, WS)
        WS.Cells(4, 16).value = WS.Cells(volMaxdex, 10).value
        WS.Cells(4, 17).value = WS.Cells(volMaxdex, 13).value
        
    Next WS
End Sub
