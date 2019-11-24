Private Function Ticker(index as Int) as String
    Ticker = WS.Cells(index,1).value
End Function 


Private Function nextIndex(currentIndex as Int) as Int
    Dim i as Int
    i = currentIndex
    Dim currentTick as String
    currentTick= Cells(currentIndex,1).value
    While Cells(i,1).value = currentTick
        i = i + 1
    Wend
    nextIndex = i
End Function

Private Function Volume(fst as Int, lst as Int) as Long
    dim total as Long
    total = 0
    For i =fst to lst
        total +=WS.Cells(7,i)
    Next i
    Volume = total
End Function

Private Function ArgMax(fst as Int, lst as Int, col as Int) as Int
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

Private Function ArgMin(fst as Int, lst as Int, col as Int) as Int
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

Private Function Min(index as Int) as Double
    Min = WS.Cells(index,5).value
End Function

Private Function Delta(fst as Int, lst as Int) as Double
    Delta = Cells(lst,6).value -Cells(fst,3).value
End Function

Private Function PercentDelta(fst as Int, lst as Int) as Double
    PercentDelta = (Cells(lst,6).value - Cells(fst,3).value)/Cells(fst,3).value
End Function

Public Sub StockAnalysis
    Dim Ticker as String
    Dim Vol as Long
    Dim Summary_Table_Row as Int
    Dim year_open as Double
    Dim Tick_Begin as Int
    Dim Tick_End as Int
    For Each WS in Worksheets
    WS.activate()
    vol = 0
    Tick_Begin = 2
    Tick_End = nextIndex(2) - 1
    WS.Cells(1,10).value = "Ticker"
    WS.Cells(1,10).value = "Yearly Change"
    WS.Cells(1,11).value = "Percent Change"
    WS.Cells(1,12).value = "Total Stock Volume"

    While Not WS.Cells(Tick_End + 1,1) = 

    Wend

    Next WS
End Sub

