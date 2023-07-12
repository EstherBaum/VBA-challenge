Attribute VB_Name = "Module1"
Sub stock_data():


    Dim i As Long
    Dim j As Integer
    'j is rows for ticker symbols ect.
    Dim beg As Double
    Dim closing As Double
    Dim change As Double
    Dim pch As Double
    Dim stvol As Double
    Dim gtv As Double
    Dim gi As Double
    Dim gd As Double
    Dim lastrow As Long
    'to run macro on all worksheets
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    

    j = 1
    beg = 0
    change = 0
    closing = 0
    stvol = 0
    gtv = 0
    gi = 0
    gd = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    
    'label cells
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    
        For i = 1 To lastrow
        
             If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
             
                
                ' if i = 1 grab opening value and go up a row
                
                If i = 1 Then
                    ' go up a row
                     j = j + 1
                    'finding opening value
                    beg = ws.Cells(i + 1, 3).Value
                    ws.Cells(j, 10) = ws.Cells(i + 1, 1)
                
                
                ' if lastrow grab closing, Yearly Change, Perentage change and stvol
                
                ElseIf i = lastrow Then
                    ' go up a row
                    j = j + 1
                    'grab closing
                    closing = ws.Cells(i, 6).Value
                    'enter change in Yearly Change Column
                    change = closing - beg
                    ws.Cells(j - 1, 11).Value = change
                    ' enter percentage change
                    pch = (change) / (beg)
                    ws.Cells(j - 1, 12).Value = FormatPercent(pch, 2)
                    'enter Total Stock Volume
                    stvol = stvol + ws.Cells(i, 7).Value
                    ws.Cells(j - 1, 13).Value = stvol
                    
                    'Calculate greatest total volume
                
                       If stvol > gtv Then
                            gtv = stvol
                            ws.Cells(4, 16).Value = ws.Cells(i, 1).Value
                            ws.Cells(4, 17).Value = gtv
                        
                        End If
                        
                      'Calculate greatest percentage increase
                      
                    
                       If ws.Cells(j - 1, 12).Value > gi Then
                        gi = ws.Cells(j - 1, 12).Value
                        ws.Cells(2, 16).Value = ws.Cells(i, 1).Value
                        ws.Cells(2, 17).Value = FormatPercent(gi, 2)
                        
                        End If
                    
                'Calculate greatest percentage decrease
                    
                        If ws.Cells(j - 1, 12).Value < gd Then
                            gd = ws.Cells(j - 1, 12).Value
                            ws.Cells(3, 16).Value = ws.Cells(i, 1).Value
                            ws.Cells(3, 17).Value = FormatPercent(gd, 2)
                        
                         End If
                    
                    
                ' all other rows where Cells(i, 1).Value <> Cells(i + 1, 1).Value
                'grab Ticker, Yearly Change, Percentage Change, New opening value, total stock volume
                
                ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    ' go up a row
                    j = j + 1
                    'calculating ticker symbol
                    ws.Cells(j, 10) = ws.Cells(i + 1, 1)
                    'calculating Yearly Change
                    closing = ws.Cells(i, 6).Value
                    change = closing - beg
                    ws.Cells(j - 1, 11).Value = change
                    'calculate percentage change
                    pch = (change) / (beg)
                    ws.Cells(j - 1, 12).Value = FormatPercent(pch, 2)
                    'calculating stvol (Total Stock Volume)
                    stvol = stvol + ws.Cells(i, 7).Value
                    ws.Cells(j - 1, 13).Value = stvol
                    ' wait to grab beg til you have calculated the change
                    beg = ws.Cells(i + 1, 3).Value
                    
                    'Calculate greatest total volume
                
                       If stvol > gtv Then
                            gtv = stvol
                            ws.Cells(4, 16).Value = ws.Cells(i, 1).Value
                            ws.Cells(4, 17).Value = gtv
                        
                        End If
                        
                   'Calculate greatest percentage increase
                      
                    
                       If ws.Cells(j - 1, 12).Value > gi Then
                        gi = ws.Cells(j - 1, 12).Value
                        ws.Cells(2, 16).Value = ws.Cells(i, 1).Value
                        ws.Cells(2, 17).Value = FormatPercent(gi, 2)
                        
                        End If
                    
                   'Calculate greatest percentage decrease
                   
                    
                        If ws.Cells(j - 1, 12).Value < gd Then
                            gd = ws.Cells(j - 1, 12).Value
                            ws.Cells(3, 16).Value = ws.Cells(i, 1).Value
                            ws.Cells(3, 17).Value = FormatPercent(gd, 2)
                        
                         End If
                    
                    'set stvol to 0
                    stvol = 0
                    
                    
                End If
            
            Else
                stvol = stvol + ws.Cells(i, 7).Value
                    
                    
         End If
        
        Next i
        'format colors for cells
        For k = 2 To lastrow
            If ws.Cells(k, 11).Value > 0 Then
                ws.Cells(k, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(k, 11).Value < 0 Then
                ws.Cells(k, 11).Interior.ColorIndex = 3
            End If
            
        Next k
        
    Next ws
        
End Sub
