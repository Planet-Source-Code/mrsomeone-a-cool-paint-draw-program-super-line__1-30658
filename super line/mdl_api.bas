Attribute VB_Name = "Module1"
Global Red As Integer
Global Green As Integer
Global Blue As Integer

Sub MirrorDraw(x As Single, y As Single, old1 As Integer, old2 As Integer, mirror As Byte)
    
    With Form1

        
        'Vertical
        If mirror = 0 Then
                .pic_paint.Line (.pic_paint.ScaleWidth - old1, old2)-(.pic_paint.ScaleWidth - x, y), RGB(Red, Green, Blue)
        End If
    
        'Horisontal
        If mirror = 1 Then
               .pic_paint.Line (old1, .pic_paint.ScaleHeight - old2)-(x, .pic_paint.ScaleWidth - y), RGB(Red, Green, Blue)
        End If
        
        'Diagonal
        If .chk_mirror_horisontal.value = 1 And .chk_mirror_vertical.value = 1 Then
            .pic_paint.Line (.pic_paint.ScaleWidth - old1, .pic_paint.ScaleHeight - old2)-(.pic_paint.ScaleWidth - x, .pic_paint.ScaleWidth - y), RGB(Red, Green, Blue)
        End If
    End With

End Sub

Sub GetRgb(Number As Long)
    
    Dim HCol As String
    
    On Error Resume Next
    HCol = Hex(Number)
    HCol = Space(6 - Len(HCol)) & HCol
    HCol = Replace(HCol, " ", "0")

    Blue = "&H" & Left(HCol, 2)
    Green = "&H" & Mid(HCol, 3, 2)
    Red = "&H" & Right(HCol, 2)
    
End Sub
Sub AutoUpdate()
        
    With Form1
        .scr_color(0).value = Red
        .scr_color(1).value = Green
        .scr_color(2).value = Blue
        
        .pic_color.BackColor = RGB(.scr_color(0).value, .scr_color(1).value, .scr_color(2).value)
        
        .lbl_color(0).Caption = Red
        .lbl_color(1).Caption = Green
        .lbl_color(2).Caption = Blue
    End With
End Sub
Sub UpdateColor(colorchange As Integer)
            
    Dim color, value As Single

    Randomize
    color = Int(Rnd * 3)

    Randomize
    value = Int(Rnd * 2)
                
    If color = 1 And value = 0 And Red + colorchange <= 255 Then Red = Red + colorchange
 
    If color = 1 And value = 1 And Red - colorchange >= 0 Then Red = Red - colorchange

    If color = 0 And value = 0 And Green + colorchange <= 255 Then Green = Green + colorchange

    If color = 0 And value = 1 And Green - colorchange >= 0 Then Green = Green - colorchange

    If color = 2 And value = 0 And Blue + colorchange <= 255 Then Blue = Blue + colorchange

    If color = 2 And value = 1 And Blue - colorchange >= 0 Then Blue = Blue - colorchange

End Sub
