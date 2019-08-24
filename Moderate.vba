Sub stockvol1():

    For Each ws In Worksheets
    Dim totalvol As Double
    Dim j As Integer
    Dim r As Long
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearlychange As Double
    Dim percentchange As Double
        
        totalvol = 0
        j = 2
        r = 2
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastrow

           If ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value Then
                totalvol = totalvol + ws.Range("G" & i).Value
            
            Else
                openprice = ws.Range("C" & r)
                closeprice = ws.Range("F" & i)
                yearlychange = closeprice - openprice

                If openprice <> 0 Then
                    percentchange = yearlychange / openprice
                                  
                End If
            
                ws.Range("I" & j).Value = ws.Range("A" & i).Value
                ws.Range("J" & j).Value = totalvol + ws.Range("G" & i).Value
                ws.Range("K" & j).Value = yearlychange
                ws.Range("K" & j).NumberFormat = "0.00000000"
                ws.Range("L" & j).Value = percentchange
                ws.Range("L" & j).NumberFormat = "0.00%"
         
                If ws.Range("K" & j).Value > 0 Then
                    ws.Range("K" & j).Interior.ColorIndex = 4
                Else
                    ws.Range("K" & j).Interior.ColorIndex = 3
                End If
                
                r = i + 1
                j = j + 1
                totalvol = 0
             
            End If

        Next i
    
    Next ws
    
End Sub
