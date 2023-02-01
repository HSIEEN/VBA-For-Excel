Sub 右旋增益和特性自动着色()
ActiveSheet.Range(Cells(28, 3), Cells(40, 15)).Select
For Each OneCell In Selection
    'MsgBox (OneCell.Value)
    
    If OneCell.Value <= -20 Then

        OneCell.Interior.Color = RGB(190, 0, 0)
        
    ElseIf OneCell.Value > -20 And OneCell.Value <= -16 Then
    
        OneCell.Interior.Color = RGB(255, 0, 0)
    
    ElseIf OneCell.Value > -16 And OneCell.Value <= -14 Then
    
        OneCell.Interior.Color = RGB(255, 200, 0)
        
    ElseIf OneCell.Value > -14 And OneCell.Value <= -12 Then
    
        OneCell.Interior.Color = RGB(255, 255, 0)
    
    ElseIf OneCell.Value > -12 And OneCell.Value <= -10 Then
    
        OneCell.Interior.Color = RGB(216, 254, 154)
    
    ElseIf OneCell.Value > -10 And OneCell.Value <= -8 Then
    
        OneCell.Interior.Color = RGB(145, 218, 0)
    
    ElseIf OneCell.Value > -8 And OneCell.Value <= -6 Then
    
        OneCell.Interior.Color = RGB(0, 180, 0)
    
    ElseIf OneCell.Value > -6 Then
    
         OneCell.Interior.Color = RGB(0, 130, 0)
     
     End If

Next
ActiveSheet.Range(Cells(3, 3), Cells(15, 15)).Select
For Each OneCell In Selection
    'MsgBox (OneCell.Value)
    
    If OneCell.Value <= -9.5 Then

        OneCell.Interior.Color = RGB(192, 0, 0)
        
    ElseIf OneCell.Value > -9.5 And OneCell.Value <= -3.3 Then
    
        OneCell.Interior.Color = RGB(255, 0, 0)
    
    ElseIf OneCell.Value > -3.3 And OneCell.Value <= 0 Then
    
        OneCell.Interior.Color = RGB(255, 192, 0)
    
    ElseIf OneCell.Value > 0 And OneCell.Value <= 2.3 Then
    
        OneCell.Interior.Color = RGB(255, 255, 0)
    
    ElseIf OneCell.Value > 2.3 And OneCell.Value <= 5.9 Then
    
        OneCell.Interior.Color = RGB(216, 254, 154)
    
    ElseIf OneCell.Value > 5.9 And OneCell.Value <= 9.5 Then
    
        OneCell.Interior.Color = RGB(145, 218, 0)
    
    ElseIf OneCell.Value > 9.5 And OneCell.Value <= 15.4 Then
    
         OneCell.Interior.Color = RGB(0, 180, 0)
    
    ElseIf OneCell.Value > 15.4 Then
    
        OneCell.Interior.Color = RGB(0, 130, 0)

     
     End If

Next

End Sub
