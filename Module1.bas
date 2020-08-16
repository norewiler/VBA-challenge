Sub formatCols()
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 9).End(xlUp).Row
    Range("i1").Value = "Ticker"
    Range("i1").Font.Bold = True
    Range("j1").Value = "Yearly Change"
    Range("j1").Font.Bold = True
    Range("k1").Value = "Percent Change"
    Range("k1").Font.Bold = True
    Range("l1").Value = "Total Stock Volume"
    Range("l1").Font.Bold = True
    
    Range("O12").Value = "Greatest % Increase"
    Range("O13").Value = "Greatest % Decrease"
    Range("O14").Value = "Greatest Total Volume"
    Range("P11").Value = "Ticker"
    Range("Q11").Value = "Value"

    
    Range("J2:J" & lastRow).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = 4
    End With
    Range("J2:J" & lastRow).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = 3
    End With
    
    Columns("K:K").Select
    Selection.NumberFormat = "0.00%"
    
    
End Sub

Sub calculateChanges2()
    
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearChange As Double
    Dim changePct As Double
    Dim total As Double
    Dim greatIncPct As Double
    greatIncPct = -100
    Dim greatIncStock As String
    Dim greatDecPct As Double
    greatDecPct = 100
    Dim greatDecStock As String
    Dim greatTotal As Double
    Dim greatTotalStock As String
    
    Dim ticker As String
    ticker = ""
    Dim checkTicker As String
    Dim dataRows As Long
    dataRows = Cells(Rows.Count, 1).End(xlUp).Row
    Dim tickerIdx As Integer
    tickerIdx = 1
    Dim first As Boolean
    first = True
        
            
    For I = 2 To dataRows
        
        checkTicker = Cells(I, 1).Value
        If checkTicker <> ticker Then
            If first Then
                total = 0
                If I = 2 Then
                    openPrice = Cells(I, 3).Value
                Else
                    openPrice = Cells(I - 1, 3).Value
                End If
                
                first = False
                tickerIdx = tickerIdx + 1
                ticker = checkTicker
                Cells(tickerIdx, 9).Value = ticker
            Else
                closePrice = Cells(I - 1, 6).Value
                yearChange = closePrice - openPrice
                If openPrice <> 0 Then
                    changePct = yearChange / openPrice
                Else
                    changePct = yearChange / 1
                End If
                
                Cells(tickerIdx, 10).Value = yearChange
                Cells(tickerIdx, 11).Value = changePct
                first = True
                
            
            End If
            
        Else
            total = total + Cells(I, 7).Value
        End If
        
        If changePct > greatIncPct Then
            greatIncPct = changePct
            greatIncStock = ticker
        ElseIf changePct < greatDecPct Then
            greatDecPct = changePct
            greatDecStock = ticker
        End If
            
        Cells(tickerIdx, 12).Value = total
        If total > greatTotal Then
            greatTotal = total
            greatTotalStock = ticker
        End If
    
    Next I
    
    Cells(12, 16).Value = greatIncStock
    Cells(12, 17).Value = greatIncPct
    Cells(13, 16).Value = greatDecStock
    Cells(13, 17).Value = greatDecPct
    Cells(14, 16).Value = greatTotalStock
    Cells(14, 17).Value = greatTotal
    
    
   

End Sub

Sub allSheets()

         Dim ws As Worksheet
         Dim maxIncPct As Double
         maxIncPct = -100
         Dim maxIncStock As String
         Dim maxDecPct As Double
         maxDecPct = 100
         Dim maxDecStock As String
         Dim maxTotal As Double
         maxTotal = -100
         Dim maxTotalStock As String
         Dim first As Boolean
         first = True
                           
        
         For Each ws In ActiveWorkbook.Worksheets
            
            ws.Select
            Call calculateChanges2
            If ws.Cells(12, 17).Value > maxIncPct Then
                maxIncPct = ws.Cells(12, 17).Value
                maxIncStock = ws.Cells(12, 16).Value
            End If
                
            If ws.Cells(13, 17).Value < maxDecPct Then
                maxDecPct = ws.Cells(13, 17).Value
                maxDecStock = ws.Cells(13, 16).Value
            End If
            
            If ws.Cells(14, 17).Value > maxTotal Then
                maxTotal = ws.Cells(14, 17).Value
                maxTotalStock = ws.Cells(14, 16).Value
            End If
            
            Call formatCols
            
            
         Next ws
                
         Sheets(1).Select
         Range("P1").Value = "Ticker"
         Range("P1").Font.Bold = True
         Range("Q1").Value = "Value"
         Range("Q1").Font.Bold = True
         
         Range("O2").Value = "Greatest % Increase"
         Range("O2").Font.Bold = True
         Range("O3").Value = "Greatest % Decrease"
         Range("O3").Font.Bold = True
         Range("O4").Value = "Greatest Total Volume"
         Range("O4").Font.Bold = True
         Cells(2, 16).Value = maxIncStock
         Cells(2, 17).Value = maxIncPct
         Cells(2, 17).NumberFormat = "0.00%"
         Cells(3, 16).Value = maxDecStock
         Cells(3, 17).Value = maxDecPct
         Cells(3, 17).NumberFormat = "0.00%"
         Cells(4, 16).Value = maxTotalStock
         Cells(4, 17).Value = maxTotal
         
         

End Sub






'Sub filterTickers()
'    Dim lastRow As Long
'
'
'    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
'
'    Range("A2:A" & lastRow).AdvancedFilter Action:=xlFilterCopy, copytorange:=Range("I1"), unique:=True
'
'
'
'
'
'End Sub




'Sub yearlyChange()
'    Dim openPrice As Double
'    Dim closePrice As Double
'    Dim yearChange As Double
'    Dim changePct As Double
'    Dim total As Double
'    Dim ticker As String
'    Dim checkTicker As String
'    Dim entryDate As String
'    Dim tickerRows As Long
'    Dim dataRows As Long
'
'    tickerRows = Cells(Rows.Count, 9).End(xlUp).Row
'    dataRows = Cells(Rows.Count, 1).End(xlUp).Row
'
'    For I = 2 To tickerRows
'        ticker = Cells(I, 9).Value
'        total = 0
'        For j = 2 To dataRows
'            checkTicker = Cells(j, 1).Value
'            entryDate = Right(Cells(j, 2).Value, 4)
'            If checkTicker = ticker Then
'                total = total + Cells(j, 7).Value
'                If entryDate = "0101" Then
'                    openPrice = Cells(j, 3).Value
'                ElseIf entryDate = "1230" Then
'                    closePrice = Cells(j, 6).Value
'                End If
'
'
'            End If
'        Next j
'        yearChange = closePrice - openPrice
'        changePct = yearChange / openPrice
'        Cells(I, 10).Value = yearChange
'        Cells(I, 11).Value = changePct
'        Cells(I, 12).Value = total
'
'    Next I
'
'
'End Sub


'Sub calculateChanges()
'
'    Dim openPrice As Double 'price that a stock opened
'    Dim closePrice As Double 'price that a stock closed
'    Dim yearChange As Double 'change in a stock price over a year
'    Dim changePct As Double ' percent change in sa stock price over a year
'    Dim total As Double ' total stock volume in a year
'    Dim ticker As String ' ticker of a stock
'    ticker = "" ' initialize the ticker to be an empty string
'    Dim checkTicker As String ' variable to hold the ticker from the dat line being looked at
''    Dim entryDate As String
'    'Dim tickerRows As Long
'    Dim dataRows As Long 'number of rows of data in the table
'    dataRows = Cells(Rows.Count, 1).End(xlUp).Row ' count the num rows in the table
''    Dim step As Double
'    'Dim i As Double
'    Dim tickerIdx As Integer ' keep track of where you want each tickers data stored
'    tickerIdx = 1 ' initialize the ticker index to be 1
'    Dim first As Boolean ' flag to represent if it's the first time a new ticker has been seen
'    first = True ' initialize the flag to be tru
'
'    'tickerRows = Cells(Rows.Count, 9).End(xlUp).Row
'
'    For I = 2 To dataRows ' for all rows in the table
'
'        checkTicker = Cells(I, 1).Value ' get the value from the ticker column
'        If checkTicker <> ticker Then ' if it's not equal to the current ticker
'            If first Then ' if this is the first time seeing the new ticker
'                total = 0 ' reset the total count
'                If I = 2 Then ' this is just used since the very first ticker needs to be treated a little differently
'                    openPrice = Cells(I, 3).Value ' grab the value from row i, col C
'                Else
'                    openPrice = Cells(I - 1, 3).Value ' in all other cases, grab the value at the row above i, col C
'                End If
'
'                first = False ' indicate that this is no longer the first time the new ticker has been seen
'                tickerIdx = tickerIdx + 1 ' increase the ticker index (Next ticker wdat will be stored in the row below)
'                ticker = checkTicker ' set the current ticker to be the ticker you just found
'                Cells(tickerIdx, 9).Value = ticker ' place the new ticker in the output table
'            Else ' not the first time the ticker has been seen, now we are at end of ticker data - need to gab close value
'                closePrice = Cells(I - 1, 6).Value
'                first = True ' reset the first flag
'
'
'            End If
'
'        Else ' the current row has same ticker as the one we are grabbing data for
'            total = total + Cells(I, 7).Value ' add stock vol for this row to the running total
'        End If
'        yearChange = closePrice - openPrice ' calc year change
'        If openPrice <> 0 Then
'            changePct = yearChange / openPrice ' calc change pct
'        Else ' handle case where open price = 0 to avoid div by 0 error
'            changePct = yearChange / 1
'        End If
'        ' store values for that ticker
'        Cells(tickerIdx, 10).Value = yearChange
'        Cells(tickerIdx, 11).Value = changePct
'        Cells(tickerIdx, 12).Value = total
'
'    Next I
'
'
'
'
'End Sub