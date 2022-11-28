Attribute VB_Name = "Module1"
Sub vbachallenge():

'I believe the data was sorted by date so I could have left out the "search for" start and end values
' and just hard coded but I wanted it to work if the data was not sorted

                                            Dim worksheetname As String
                                            Dim outputrow As Integer
                                            Dim totalvolume As Double
                                            Dim firstdate As Long
                                            Dim lastdate As Long
                                            Dim firstdatevalue As Double
                                            Dim lastdatevalue As Double
                                            Dim annualchange As Double
                                            Dim percentchange As Double
                                            Dim highvolume As Double
                                            Dim maxincrease As Double
                                            Dim maxdecrease As Double
                                            Dim highvolumeticker As String
                                            Dim maxincreaseticker As String
                                            Dim maxdecreaseticker As String

For Each ws In Worksheets

            'setting output headings
                ws.Range("i1").Value = "Ticker"
                ws.Range("j1").Value = "Yearly Change"
                ws.Range("k1").Value = "Percent Change"
                ws.Range("l1").Value = "Total Stock Volume"
                ws.Range("o2").Value = "Greatest % Increase"
                ws.Range("o3").Value = "Greatest % Decrease"
                ws.Range("o4").Value = "Greatest Total Volume"
                ws.Range("p1").Value = "Ticker"
                ws.Range("q1").Value = "Value"
            
            'setting start values per sheet
                outputrow = 1
                totalvolume = 0
                firstdatevalue = 0
                lastdatevalue = 0
                 'using dates that would allow analysis for periods between 1900 and 2100
                firstdate = 21003112
                lastdate = 19000101
            
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
For i = 2 To lastrow

    If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
    
                            outputrow = outputrow + 1
                            ticker = ws.Cells(i, 1)
                             
                            
                            'total volume per ticker
                                totalvolume = totalvolume + ws.Cells(i, 7).Value
                                ws.Range("I" & outputrow).Value = ticker
                                ws.Range("L" & outputrow).Value = totalvolume
                            'BONUS
                                If totalvolume > highvolume Then
                                       highvolume = totalvolume
                                       highvolumeticker = ws.Cells(i, 1).Value
                                End If
                         
                             'reset total volume per ticker
                                totalvolume = 0
                                
                            'annual change per ticker
                
                            If firstdate > ws.Cells(i, 2).Value Then
                                firstdate = ws.Cells(i, 2).Value
                                firstdatevalue = ws.Cells(i, 3).Value
                            End If
                        
                            If lastdate < ws.Cells(i, 2).Value Then
                                lastdate = ws.Cells(i, 2).Value
                                lastdatevalue = ws.Cells(i, 6).Value
                            End If
                            
                            annualchange = lastdatevalue - firstdatevalue
                            percentchange = annualchange / firstdatevalue
                            ws.Cells(outputrow, 25).Value = lastdatevalue
                            ws.Cells(outputrow, 24).Value = firstdatevalue
                            
                            'reset change date values
                            firstdate = 21003112
                            lastdate = 19000101
                            
                            
                            'format yearly change output
                            ws.Range("J" & outputrow) = annualchange
                                    If ws.Range("J" & outputrow) > 0 Then
                                        ws.Range("J" & outputrow).Interior.ColorIndex = 4
                                    ElseIf ws.Range("J" & outputrow) < 0 Then
                                        ws.Range("J" & outputrow).Interior.ColorIndex = 3
                                    ElseIf ws.Range("J" & outputrow) = 0 Then
                                        ws.Range("J" & outputrow).Interior.ColorIndex = 2
                                    End If
                                                     
                             'format percentage change output
                             ws.Range("K" & outputrow) = percentchange
                             ws.Range("K" & outputrow).NumberFormat = "0.00%"
                                    FormatPercent (ws.Range("K" & outputrow))
                                    
                           'BONUS
                                   
                                            
                                    If percentchange > maxincrease Then
                                        maxincrease = percentchange
                                        maxincreaseticker = ws.Cells(i, 1).Value
                                    End If
                                            
                                     If percentchange < maxdecrease Then
                                        maxdecrease = percentchange
                                        maxdecreaseticker = ws.Cells(i, 1).Value
                                    End If
                                    
                             
    Else
    
                            totalvolume = totalvolume + ws.Cells(i, 7).Value
                            
                                    'start and end date find
                                      If firstdate > ws.Cells(i, 2).Value Then
                                        firstdate = ws.Cells(i, 2).Value
                                        firstdatevalue = ws.Cells(i, 3).Value
                        
            
                                    End If
                                    
                                    If lastdate < ws.Cells(i, 2).Value Then
                                        lastdate = ws.Cells(i, 2).Value
                                        lastdatevalue = ws.Cells(i, 6).Value
                                    End If
                                    
                                
                                                
                    
    
    End If
    
    
Next i

'BONUS
ws.Range("Q4").Value = highvolume
ws.Range("P4").Value = highvolumeticker
ws.Range("Q2").Value = maxincrease
ws.Range("P2").Value = maxincreaseticker
ws.Range("Q3").Value = maxdecrease
ws.Range("P3").Value = maxdecreaseticker
'FORMAT BONUS
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").NumberFormat = "0"

'Reset BONUS
highvolume = 0
maxincrease = 0
maxdecrease = 0

'set column width
Columns("A:Q").Select
Selection.Columns.AutoFit

Next ws


End Sub

