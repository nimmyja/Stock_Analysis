Sub Stock()

    For Each ws In Worksheets
    
'Grabbed the WorksheetName

        Dim WorksheetName As String
        WorksheetName = ws.Name
        Dim currentticker As String

'Initialized the variables

        Counter = 2
        currentticker = ws.Range("A2").Value
        min_date = ws.Range("B2").Value
        max_Date = ws.Range("B2").Value
        openvalue = ws.Range("C2").Value
        closevalue = ws.Range("F2").Value
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        volume = ws.Range("G2").Value

        For i = 3 To lastrow
            
            
            If currentticker = ws.Cells(i, 1).Value Then
                volume = volume + ws.Cells(i, 7).Value
                
                    If ws.Cells(i, 2).Value < min_date Then
                        min_date = ws.Cells(i, 2).Value
                        openvalue = ws.Cells(i, 3).Value
                    Else
                        max_Date = ws.Cells(i, 2).Value
                        closevalue = ws.Cells(i, 6).Value
                    End If
                
            Else
            
'Write Method call
                Call write_ticker_summary(ws, currentticker, openvalue, closevalue, Counter, volume)
                Counter = Counter + 1
                currentticker = ws.Cells(i, 1).Value
                min_date = ws.Cells(i, 2).Value
                max_Date = ws.Cells(i, 2).Value
                openvalue = ws.Cells(i, 3).Value
                closevalue = ws.Cells(i, 6).Value
                volume = ws.Cells(i, 7).Value
                
            End If
        Next i

    Call write_ticker_summary(ws, currentticker, openvalue, closevalue, Counter, volume)
    Call analysis_table(ws)
    
    Next ws
    
End Sub

Sub write_ticker_summary(ws, currentticker, openvalue, closevalue, Counter, volume)

'Column headers for new table

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
'Yearly(&)Percent change & Total Stock volume for each ticker

        ws.Range("I" & Counter).Value = currentticker
        ws.Range("J" & Counter).Value = closevalue - openvalue
        
'Division by Zero Error Correction

        If openvalue = 0 Then
            ws.Range("K" & Counter).Value = 0
        Else
            ws.Range("K" & Counter).Value = (closevalue - openvalue) / openvalue
        End If
        
'Number Format the percent change

        ws.Range("K" & Counter).NumberFormat = "0.00%"
        ws.Range("L" & Counter).Value = volume
        yearly_change = ws.Range("J" & Counter).Value
        
'Color Format the yearly change

        If yearly_change > 0 Then
            ws.Range("J" & Counter).Interior.ColorIndex = 4
        Else
            ws.Range("J" & Counter).Interior.ColorIndex = 3
        End If
        
 
                
End Sub


Sub analysis_table(ws)

lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
For i = 2 To lastrow
    ws.Range("Q2") = Application.WorksheetFunction.Max(ws.Range("K2:K" & i))
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3") = Application.WorksheetFunction.Min(ws.Range("K2:K" & i))
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4") = Application.WorksheetFunction.Max(ws.Range("L2:L" & i))
    
    If ws.Range("Q2").Value = ws.Range("K" & i) Then
        ws.Range("P2").Value = ws.Range("I" & i)
    End If
    
    If ws.Range("Q3").Value = ws.Range("K" & i) Then
        ws.Range("P3").Value = ws.Range("I" & i)
    End If
    
    If ws.Range("Q4").Value = ws.Range("L" & i) Then
        ws.Range("P4").Value = ws.Range("I" & i)
    End If
Next i
        
End Sub

