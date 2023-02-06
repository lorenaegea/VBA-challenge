Attribute VB_Name = "Module1"
Sub stockData():

Dim ws As Worksheet

For Each ws In Worksheets
            
                ' add the word Ticker into cell I1
                ws.Range("I1").Value = "Ticker"
                
                ' add the word Yearly Chanve into cell "J1"
                ws.Range("J1").Value = "Yearly Change"
                
                ' add the word Percent Change into cell "K1"
                ws.Range("K1").Value = "Percent Change"
                
                ' add the word Total Stock Volume into cell "L1"
                ws.Range("L1").Value = "Total Stock Volume"
                
                ws.Range("N2").Value = "Greatest % Increase"
                ws.Range("N3").Value = "Greatest % Decrease"
                ws.Range("N4").Value = "Greatest Total Volume"
                
                ws.Range("O1").Value = "Ticker"
                ws.Range("P1").Value = "Value"
                
                ' declare variable to hold the last row number
                 Dim lastRow As Integer
                lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
                Dim lastRow2 As Integer
                lastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
                
                ' variable to hold the rows in the Ticker column (I)
                Dim tickerRows As Integer
                tickerRows = 2
                
                Dim openRows As Integer
                openRows = 2
                
                Dim closeRows As Integer
                closeRows = 2
                
                Dim ycRows As Integer
                ycRows = 2
                
                Dim pcRows As Integer
                pcRows = 2
                
                Dim vtRows As Integer
                vtRows = 2
                
                Dim start As Integer
                start = 2
                
                ' declare variable to hold the ticker
                Dim ticker As String
                
                ' declare variable to hold the open price
                Dim openValue As Double
                
                ' declare variable to hold the close price
                Dim closeValue As Double
                
                Dim yearlyChange As Double
                
                Dim percentChange As Double
                
                ' declare variable to hold the total volume
                Dim volumeTotal As Double
                volumeTotal = 0
                
                ' declare variable to hold the row in column A
                Dim rowA As Integer
                
                Dim matchIncrease, matchDecrease, matchVolume As String
                
                ' loop through the rows and check the changes in the ticker value
                For rowA = 2 To lastRow
                
                        ' conditional: check if the "next" row has a different ticker than the "current" row
                        If ws.Cells(rowA + 1, 1).Value <> ws.Cells(rowA, 1).Value Then
                
                                ' add ticker to column (I, tickerRow)
                                ticker = ws.Cells(rowA, 1).Value
                                
                                ws.Cells(tickerRows, 9).Value = ticker
                        
                                ' close price
                                closeValue = ws.Cells(rowA, 6).Value
                                
                                ' ws.Cells(tickerRows, 15).Value = closeValue
                                
                                tickerRows = tickerRows + 1
                                
                                'yearlyChange = closeValue - openValue
                        
                                'ws.Cells(ycRows, 10).Value = yearlyChange
                        
                                ' ycRows = ycRows + 1
                                
                                If ws.Cells(start, 3) = 0 Then
                                    
                                        For findValue = start To rowA
                                            
                                            If ws.Cells(findValue, 3).Value <> 0 Then
                                            
                                                    start = findValue
                                                    
                                                    Exit For
                                                    
                                            End If
                                            
                                      Next findValue
                                
                                End If
                                
                                yearlyChange = ws.Cells(rowA, 6) - ws.Cells(start, 3)
                                
                                ws.Cells(pcRows, 10).Value = yearlyChange
                                
                                percentChange = Round((yearlyChange / ws.Cells(start, 3) * 100), 2)
                                
                                ws.Cells(pcRows, 11).Value = percentChange
                                
                                ' ws.Cells(pcRows, 11).Value = Format(actCell, "#.####%")
                                
                                pcRows = pcRows + 1
                                
                                volumeTotal = volumeTotal + ws.Cells(rowA, 7)
                                
                                ws.Cells(vtRows, 12).Value = volumeTotal
                                
                                vtRows = vtRows + 1
                                
                                volumeTotal = 0
                                
                        ElseIf ws.Cells(rowA - 1, 1).Value <> ws.Cells(rowA, 1).Value Then
                        
                                openValue = ws.Cells(rowA, 3).Value
                                
                                ' ws.Cells(openRows, 14).Value = openValue
                                
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
                ' matchIncrease = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow2)), ws.Range("K2:K" & lastRow2), 0)
                
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
                
        Next ws

End Sub

