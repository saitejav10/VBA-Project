Sub stockvol():
    
    For Each ws In Worksheets
    Dim totalvol As Double
    Dim j As Integer
    
        totalvol = 0
        j = 2

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastrow

           If ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value Then
                totalvol = totalvol + Range("G" & i).Value
            
                Else
            
                  ws.Range("I" & j).Value = ws.Range("A" & i).Value
                  ws.Range("J" & j).Value = totalvol + ws.Range("G" & i).Value
                
                j = j + 1
                totalvol = 0
             
            End If

        Next i

    Next ws

End Sub
