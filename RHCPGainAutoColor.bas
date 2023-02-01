Sub 右旋增益自动着色()
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
    
        OneCell.Interior.Color = RGB(145,218, 0)
    
    ElseIf OneCell.Value > -8 And OneCell.Value <= -6 Then
    
        OneCell.Interior.Color = RGB(0, 180, 0)
    
    ElseIf OneCell.Value > -6 Then
    
         OneCell.Interior.Color = RGB(0, 130, 0)
     
     End If

Next

End Sub
