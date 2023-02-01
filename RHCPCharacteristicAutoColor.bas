Sub 右旋特征自动着色()
For Each OneCell In Selection
    'MsgBox (OneCell.Value)
    
    If OneCell.Value <=-9.5 Then

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
