Private Function TickerVal(index as Int, WS as Worksheet) as String
    TickerVal = WS.Cells(index,1).value
End Function 


Private Function nextIndex(currentIndex as Int, WS as Worksheet) as Int
    Dim i as Int
    i = currentIndex
    Dim currentTick as String
    currentTick= WS.Cells(currentIndex,1).value
    While WS.Cells(i,1).value = currentTick
        i = i + 1
    Wend
    nextIndex = i
End Function

Private Function Volume(fst as Int, lst as Int, WS as Worksheet) as Long
    dim total as Long
    total = 0
    For i =fst to lst
        total +=WS.Cells(7,i)
    Next i
    Volume = total
End Function

Private Function ArgMax(fst as Int, lst as Int, col as Int, WS as Worksheet) as Int
    dim runningArgMax as Int
    runningArgmax = fst
    dim runningMax as Double
    runningMax = WS.Cells(fst,col).value
    For i = fst to lst
        If WS.Cells(i,col).value > runningMax Then 
            runningMax = WS.Cells(i,col)
            runningArgMin = i
        End If
    Next i
    ArgMax = runningArgMax
End Function

Private Function ArgMin(fst as Int, lst as Int, col as Int, WS as WorkSheet) as Int
    dim runningArgMin as Int
    runningArgMin = fst
    dim runningMin as Double
    runningMin= WS.Cells(fst,col).Value
    For i = fst to lst
        If WS.Cells(i,col) < runningMin Then 
            runningMin = WS.Cells(i,col)
            runningArgMin = i
        End If
    Next i 
    ArgMin = runningArgMin
End Function

Private Function Delta(fst as Int, lst as Int, WS as Worksheet) as Double
    Delta = WS.Cells(lst,6).value - WS.Cells(fst,3).value
End Function

Private Function PercentDelta(fst as Int, lst as Int, WS as Worksheet) as Double
    PercentDelta = 100*(WS.Cells(lst,6).value - WS.Cells(fst,3).value)/WS.Cells(fst,3).value
End Function

Public Sub StockAnalysis
    Dim Ticker as String

    Dim Vol as Long

    Dim Summary_Table_Row as Int

    Dim yearlyChange as Double

    Dim percentChange as Double

    Dim Tick_Begin as Int
    Dim Tick_End as Int

    Dim changeMindex as Int
    Dim changeMaxdex as Int

    Dim volMaxdex as Int

    For Each WS in Worksheets
        WS.activate()
        vol = 0
        WS.Cells(1,10).value = "Ticker"
        WS.Cells(1,11).value = "Yearly Change"
        WS.Cells(1,12).value = "Percent Change"
        WS.Cells(1,13).value = "Total Stock Volume"

        Tick_Begin = 2
        Tick_End = nextIndex(2,WS) - 1
        Summary_Table_Row = 2

        While Not isEmpty(WS.Cells(Tick_Begin,1))
            Ticker = TickerVal(Tick_Begin,WS)
            Vol = Volume(Tick_Begin,Tick_End, WS)
            yearlyChange = Delta(Tick_Begin,Tick_End,WS)
            percentChange = PercentDelta(Tick_Begin,Tick_End,WS)
            
            WS.Cells(Summary_Table_Row,10).value = Ticker
            WS.Cells(Summary_Table_Row,11).value = yearlyChange
            WS.Cells(Summary_Table_Row,12).value = percentChange
            WS.Cells(Summary_Table_Row,13).value = Vol
            
            If(yearlyChange > 0 ) Then
            WS.Cells(Summary_Table_Row,11).Interior.ColorIndex = 4
            Else If(yearlyChange < 0) Then
            WS.Cells(Summary_Table_Row,11).Interior.ColorIndex = 3
            End If

            Tick_Begin = Tick_End+1
            Tick_End = nextIndex(Tick_Begin,WS) -1
            Summary_Table_Row = Summary_Table_Row+1
        Wend

        Summary_Table_Row = Summary_Table_Row - 1

        WS.Cells(2,15).value = "Greatest Percent Increase"
        WS.Cells(3,15).value = "Greatest Percent Decrease"
        WS.Cells(4,15).value = "Greatest Total Volume"

        WS.Cells(1,16).value = "Ticker"
        WS.Cells(1,17).value = "Value"


        changeMindex = ArgMin(2,Summary_Table_Row,12)
        changeMaxdex = ArgMax(2,Summary_Table_Row,12)

        WS.Cells(2,16).value = WS.Cells(changeMaxdex,10).value
        WS.Cells(2,17).value = WS.Cells(changeMaxdex,12).value

        WS.Cells(3,16).value = WS.Cells(changeMindex,10).value
        WS.Cells(3,17).value = WS.Cells(changeMindex,12).value

        volMaxdex = ArgMax(2, Summary_Table_Row, 13)
        WS.Cells(4,16).value = WS.Cells(volMaxdex,10).value
        WS.Cells(4,17).value = WS.Cells(volMaxdex,13).value
        
    Next WS
End Sub

