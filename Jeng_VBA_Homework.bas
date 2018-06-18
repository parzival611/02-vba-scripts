Attribute VB_Name = "Module1"
Sub CalculateTotal()

    Dim Brand_Name As String
    
    Dim Brand_Total As Double
    Brand_Total = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    For i = 2 To 43400

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
             
            Brand_Name = Cells(i, 1).Value
            
            Brand_Total = Brand_Total + Cells(i, 7).Value
            
            Range("I" & Summary_Table_Row).Value = Brand_Name
            
            Range("J" & Summary_Table_Row).Value = Brand_Total
            Summary_Table_Row = Summary_Table_Row + 1
            
            Brand_Total = 0
            
   
        Else
        
             Brand_Total = Brand_Total + Cells(i, 7).Value
            
        End If
        
    Next i

End Sub



