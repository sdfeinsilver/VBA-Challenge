Sub StockAnalysis()

Dim ticker As String
Dim vol As Double
Dim j As Double
Dim closep As Double
Dim openp As Double



j = 2

vol = 0

Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

openp = Cells(2, 3).Value


For i = 2 To Lastrow
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        ticker = Cells(i, 1).Value
        vol = Cells(i, 7).Value + vol
        closep = Cells(i, 6).Value
        
        Cells(j, 12).Value = (closep - openp)
        Cells(j, 10).Value = ticker
        Cells(j, 11).Value = vol
        Cells(j, 13).Value = (closep - openp) / openp
        If Cells(j, 12).Value > 0 Then
            Cells(j, 12).Interior.ColorIndex = 4
        Else
            Cells(j, 12).Interior.ColorIndex = 3
        End If
        
        openp = Cells(i + 1, 3).Value
        j = j + 1
        vol = 0
        
    Else
        vol = Cells(i, 7).Value + vol
        
        
    End If
Next i


        
    
    
    


End Sub
