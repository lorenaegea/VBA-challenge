Attribute VB_Name = "Module1"
Sub stockData():

Dim ws As Worksheet

For Each ws In Worksheets
            
                ' add labels
                ws.Range("I1").Value = "Ticker"
                ws.Range("J1").Value = "Yearly Change"
                ws.Range("K1").Value = "Percent Change"
                ws.Range("L1").Value = "Total Stock Volume"
                
                ws.Range("N2").Value = "Greatest % Increase"
                ws.Range("N3").Value = "Greatest % Decrease"
                ws.Range("N4").Value = "Greatest Total Volume"
                
                ws.Range("O1").Value = "Ticker"
                ws.Range("P1").Value = "Value"
                
                ' declare variable to hold the last ticker row number
                 Dim lastRow As Long
                lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
                ' declare variable to hold the last summary table row number
                Dim lastRow2 As Long
                lastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

                Dim tickerRows, openRows, vtRows, findValue As Integer
                tickerRows = 2
                openRows = 2
                vtRows = 2
                
                Dim rowA As Long
                
                Dim ticker, matchIncrease, matchDecrease, matchVolume As String
                
                ' declare variable to hold the open price
                Dim openValue, closeValue, yearlyChange, percentChange, volumeTotal As Double
                volumeTotal = 0
                
                ' loop through the rows and check the changes in the ticker value
                For rowA = 2 To lastRow
                
                        ' conditional: check if the "next" row has a different ticker than the "current" row
                        If ws.Cells(rowA + 1, 1).Value <> ws.Cells(rowA, 1).Value Then
                
                                ' add ticker to column (I, tickerRow), then move down to next row
                                ticker = ws.Cells(rowA, 1).Value
                                
                                ws.Cells(tickerRows, 9).Value = ticker
                        
                                ' close price
                                closeValue = ws.Cells(rowA, 6).Value
                                
                                ' calculate yearlyChange (closeValue - openValue)
                                yearlyChange = closeValue - openValue
                                
                                ws.Cells(tickerRows, 10).Value = yearlyChange
                                
                                percentChange = Round((yearlyChange / openValue), 4)
                                
                                ws.Cells(tickerRows, 11).Value = percentChange
                                
                                ' pcRows = pcRows + 1
                                
                                tickerRows = tickerRows + 1
                                
                                ' calculate volumeTotal
                                volumeTotal = volumeTotal + ws.Cells(rowA, 7)
                                
                                ws.Cells(vtRows, 12).Value = volumeTotal
                                
                                vtRows = vtRows + 1
                                
                                volumeTotal = 0
                            
                        ' conditional: check if the "previous" row has a different ticker than the "current" row
                        ElseIf ws.Cells(rowA - 1, 1).Value <> ws.Cells(rowA, 1).Value Then
                        
                                 If ws.Cells(rowA, 3) = 0 Then
                                    
                                        For findValue = rowA To lastRow
                                            
                                            If ws.Cells(findValue, 3).Value <> 0 Then
                                            
                                                    openValue = ws.Cells(findValue, 3).Value
                                                    
                                                    Exit For
                                                    
                                            End If
                                            
                                        Next findValue
                                                    
                                  Else
                                  
                                        openValue = ws.Cells(rowA, 3).Value
                                    
                                End If
                                
                                openRows = openRows + 1
                                
                                volumeTotal = volumeTotal + ws.Cells(rowA, 7)
                                
                                ws.Cells(vtRows, 12).Value = volumeTotal
                                
                        Else
                        
                                volumeTotal = volumeTotal + ws.Cells(rowA, 7)
                                
                                ws.Cells(vtRows, 12).Value = volumeTotal
                                
                        End If
                        
                Next rowA
                
                        ' calculate greatest increase
                        ws.Range("P2").Value = WorksheetFunction.Max(ws.Range("K2:K" & lastRow2))
                       
                        ' calculate greatest decrease
                        ws.Range("P3").Value = WorksheetFunction.Min(ws.Range("K2:K" & lastRow2))
                
                        ' calculate greatest total volume
                        ws.Range("P4").Value = WorksheetFunction.Max(ws.Range("L2:L" & lastRow2))
                 
                        ' match greatest % increase value to ticker
                        matchIncrease = WorksheetFunction.Match(ws.Range("P2").Value, ws.Range("K2:K" & lastRow2), 0)
                        ws.Range("O2").Value = ws.Range("I" & matchIncrease + 1).Value
        
                        ' match greatest % decrease value to ticker
                        matchDecrease = WorksheetFunction.Match(ws.Range("P3").Value, ws.Range("K2:K" & lastRow2), 0)
                        ws.Range("O3").Value = ws.Range("I" & matchDecrease + 1).Value
                
                        ' match greatest volume increase value to ticker
                        matchVolume = WorksheetFunction.Match(ws.Range("P4").Value, ws.Range("L2:L" & lastRow2), 0)
                        ws.Range("O4").Value = ws.Range("I" & matchVolume + 1).Value
                
                        ' Autofit the columns
                        ws.Range("I1:L1").Columns.AutoFit
                        ws.Range("N:P").Columns.AutoFit
                
                For i = 2 To lastRow2
                
                ' color % increase green, and % decrease red
                        If ws.Range("J" & i).Value >= 0 Then
        
                                ws.Range("J" & i).Interior.ColorIndex = 4
                        
                        Else

                                ws.Range("J" & i).Interior.ColorIndex = 3
                        
                        End If
                
                Next i
                
                For p = 2 To lastRow2
                
                        ' show percent symbol
                        ws.Range("K" & p).NumberFormat = "0.00%"
                
                Next p
               
Next ws
                

End Sub


