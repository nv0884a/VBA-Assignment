Attribute VB_Name = "Module1"
Sub Stock_Data()

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    
    total_vol = 0
    i_pointer = 2
    c_pointer = 2
    
    
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    
    For i = 2 To RowCount
    
        If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
            total_vol = total_vol + Cells(i, "G").Value
            
            open_price = Cells(c_pointer, "C").Value
            closing_price = Cells(i, "F").Value
            yearly_change = closing_price - open_price
            
            
            Cells(i_pointer, "I").Value = Cells(i, "A").Value
            Cells(i_pointer, "J").Value = yearly_change
            Cells(i_pointer, "K").Value = "%" & (yearly_change / open_price * 100)
            Cells(i_pointer, "L").Value = total_vol
            
            If yearly_change > 0 Then
                Cells(i_pointer, "J").Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(i_pointer, "J").Interior.ColorIndex = 3
            Else
                Cells(i_pointer, "J").Interior.ColorIndex = 2
            End If
                
            
            
            c_pointer = i + 1
            i_pointer = i_pointer + 1
            total_vol = 0

        Else
            total_vol = total_vol + Cells(i, "G").Value
        
        End If
        
        
        
    Next i



End Sub
