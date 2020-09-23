VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Line BETA By MrSomeone Alias Jonas Persson"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Mirror Mode"
      Height          =   975
      Left            =   10920
      TabIndex        =   33
      Top             =   2760
      Width           =   2175
      Begin VB.CheckBox chk_mirror_vertical 
         Caption         =   "Mirror Mode  Vertical"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox chk_mirror_horisontal 
         Caption         =   "Mirror Mode Horisontal"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame fra_pen 
      Caption         =   "Pen"
      Height          =   945
      Left            =   8640
      TabIndex        =   28
      Top             =   7560
      Width           =   2175
      Begin VB.CheckBox chk_pen 
         Caption         =   "Pen On /Off"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "SavePicture"
      Height          =   255
      Left            =   10920
      TabIndex        =   32
      Top             =   8280
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar prgb_draw 
      Height          =   135
      Left            =   120
      TabIndex        =   31
      Top             =   8640
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmd_new 
      Caption         =   "New Picture"
      Height          =   255
      Left            =   10920
      TabIndex        =   30
      Top             =   7920
      Width           =   2175
   End
   Begin VB.HScrollBar scr_colorchange 
      Height          =   255
      Left            =   8760
      Max             =   50
      TabIndex        =   25
      Top             =   4560
      Value           =   10
      Width           =   1695
   End
   Begin VB.Frame fra_tools 
      Caption         =   "Background editing"
      Height          =   2415
      Left            =   8640
      TabIndex        =   14
      Top             =   5040
      Width           =   2175
      Begin VB.CommandButton cmd_draw 
         Caption         =   "Draw"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   1935
      End
      Begin VB.OptionButton opt_randcircle 
         Caption         =   "Randomized Circles"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1935
      End
      Begin VB.OptionButton opt_randline 
         Caption         =   "Randomized Lines"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton opt_vertical 
         Caption         =   "Vertical Lines"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton opt_horisontal 
         Caption         =   "Horisontal Lines"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton opt_circles 
         Caption         =   "Circles"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton opt_square 
         Caption         =   "Squares"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fra_settings 
      Caption         =   "Settings"
      Height          =   2175
      Left            =   8640
      TabIndex        =   11
      Top             =   2760
      Width           =   2175
      Begin VB.HScrollBar scr_drawwidth 
         Height          =   255
         Left            =   120
         Max             =   99
         Min             =   1
         TabIndex        =   23
         Top             =   1200
         Value           =   4
         Width           =   1695
      End
      Begin VB.CheckBox chk_autoredraw 
         Caption         =   "Auto Redraw"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chk_autoupdate 
         Caption         =   "Auto Update Colors"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lbl_colorchange 
         Alignment       =   2  'Center
         Caption         =   "10"
         Height          =   255
         Left            =   1850
         TabIndex        =   27
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label lbl_drawwidth 
         Alignment       =   2  'Center
         Caption         =   "4"
         Height          =   255
         Left            =   1850
         TabIndex        =   26
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lbl_topic 
         Caption         =   "Color change per step:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lbl_topic 
         Caption         =   "Drawwidth: (Rec. 4)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog comd_1 
      Left            =   10920
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pic_color 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8640
      ScaleHeight     =   345
      ScaleWidth      =   2145
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox pic_scale 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   9285
      MousePointer    =   2  'Cross
      ScaleHeight     =   225
      ScaleWidth      =   1500
      TabIndex        =   5
      Top             =   2040
      Width           =   1530
   End
   Begin VB.PictureBox pic_scale 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   9285
      MousePointer    =   2  'Cross
      ScaleHeight     =   225
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   1320
      Width           =   1530
   End
   Begin VB.PictureBox pic_scale 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   9285
      MousePointer    =   2  'Cross
      ScaleHeight     =   225
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   600
      Width           =   1530
   End
   Begin VB.HScrollBar scr_color 
      Height          =   255
      Index           =   2
      LargeChange     =   20
      Left            =   8640
      Max             =   255
      SmallChange     =   5
      TabIndex        =   2
      Top             =   2400
      Width           =   2175
   End
   Begin VB.HScrollBar scr_color 
      Height          =   255
      Index           =   1
      LargeChange     =   20
      Left            =   8640
      Max             =   255
      SmallChange     =   5
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.HScrollBar scr_color 
      Height          =   255
      Index           =   0
      LargeChange     =   20
      Left            =   8640
      Max             =   255
      SmallChange     =   5
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.PictureBox pic_paint 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   4
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   120
      ScaleHeight     =   559
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   559
      TabIndex        =   10
      Top             =   120
      Width           =   8415
      Begin VB.Line line_mirror_vertical 
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   280
         X2              =   280
         Y1              =   0
         Y2              =   560
      End
      Begin VB.Line Line_mirror_horisontal 
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   0
         X2              =   560
         Y1              =   264
         Y2              =   264
      End
   End
   Begin VB.Label lbl_color 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   8900
      TabIndex        =   9
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label lbl_color 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   8900
      TabIndex        =   8
      Top             =   1350
      Width           =   375
   End
   Begin VB.Label lbl_color 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   8900
      TabIndex        =   7
      Top             =   630
      Width           =   375
   End
   Begin VB.Shape shp_color 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   8640
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shp_color 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   8640
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shp_color 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   8640
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x1%, x2%, y1%, y2%

Dim placing As Boolean
Dim placed As Boolean

Dim circleposx%, circleposy%

Dim oldx%
Dim oldy%

Dim distance%



Private Sub chk_autoredraw_Click()

    If chk_autoredraw.value = 1 Then
        pic_paint.AutoRedraw = True
    Else:
        pic_paint.AutoRedraw = False
    End If

End Sub



Private Sub chk_mirror_horisontal_Click()

    If chk_mirror_horisontal.value = 1 Then
        Line_mirror_horisontal.Visible = True
    Else
        Line_mirror_horisontal.Visible = False
    End If

End Sub

Private Sub chk_mirror_vertical_Click()

    If chk_mirror_vertical.value = 1 Then
        line_mirror_vertical.Visible = True
    Else
        line_mirror_vertical.Visible = False
    End If
    
End Sub

Private Sub chk_pen_Click()

    If chk_pen.value = 1 Then
         fra_tools.Enabled = False
    Else
        fra_tools.Enabled = True
        
        chk_mirror_vertical.value = 0
        chk_mirror_horisontal.value = 0
        
        line_mirror_vertical.Visible = False
        Line_mirror_horisontal.Visible = False
    End If

End Sub

Private Sub cmd_draw_Click()

    pic_paint.Cls
    
    Call draw
    
End Sub
Private Sub draw()

    fra_pen.Enabled = False

    If opt_circles.value = True Or opt_square.value = True Then
        
        If placed = False Then
            MsgBox "Place then start position on the picture"
            placing = True
            cmd_draw.Enabled = False
            Exit Sub
        
        ElseIf placed = True Then
        
            cmd_draw.Enabled = True
            
            placed = False
            placing = False
            
        End If
        
    Else
        placing = False
        placed = False
        cmd_draw.Enabled = True
    
    End If
    
    If opt_randcircle.value = True Or opt_randline.value = True Then
        pic_paint.ScaleMode = 1
    End If
    
    Dim length As Double
    
    If opt_circles.value = True Or opt_square.value = True Then
        length = ((pic_paint.ScaleHeight ^ 2) + (pic_paint.ScaleWidth ^ 2)) ^ (1 / 2)
    Else
        length = pic_paint.Height
    End If
    
    Red = scr_color(0).value
    Green = scr_color(1).value
    Blue = scr_color(2).value
    
    prgb_draw.Min = 0
    prgb_draw.Max = length + 1
    
    For a% = 0 To length
        
        If opt_vertical.value = True Then
            pic_paint.Line (a%, 0)-(a%, pic_paint.Height), RGB(Red, Green, Blue)
        
        ElseIf opt_horisontal.value = True Then
            pic_paint.Line (0, a%)-(pic_paint.Width, a%), RGB(Red, Green, Blue)
        
        ElseIf opt_randline.value = True Then
            Randomize
            x1 = Int(Rnd * pic_paint.Width)
            x2 = Int(Rnd * pic_paint.Width)
            y1 = Int(Rnd * pic_paint.Height)
            y2 = Int(Rnd * pic_paint.Height)
            
            pic_paint.Line (x1, y1)-(x2, y2), RGB(Red, Green, Blue)
        
        ElseIf opt_randcircle.value = True Then
            Randomize
            x1 = Int(Rnd * pic_paint.Width)
            x2 = Int(Rnd * pic_paint.Width)
            y1 = Int(Rnd * pic_paint.Height)
            
            pic_paint.Circle (x1, y1), x2, RGB(Red, Green, Blue)
            
        ElseIf opt_square.value = True Then
            pic_paint.Line (circleposx + a%, circleposy + a%)-(circleposx + a%, circleposy - a%), RGB(Red, Green, Blue)
            pic_paint.Line (circleposx + a%, circleposy - a%)-(circleposx - a%, circleposy - a%), RGB(Red, Green, Blue)
            pic_paint.Line (circleposx - a%, circleposy - a%)-(circleposx - a%, circleposy + a%), RGB(Red, Green, Blue)
            pic_paint.Line (circleposx - a%, circleposy + a%)-(circleposx + a%, circleposy + a%), RGB(Red, Green, Blue)
                    
        ElseIf opt_circles.value = True Then
            pic_paint.Circle (circleposx, circleposy), a%, RGB(Red, Green, Blue)
        
        End If
    
        Call UpdateColor(scr_colorchange.value)
        
        prgb_draw.value = a%
        
        If chk_autoupdate.value = 1 Then
            Call AutoUpdate
        End If
        
    Next
    
    fra_pen.Enabled = True
    
    pic_paint.ScaleMode = 3
    
End Sub

Private Sub cmd_new_Click()

    pic_paint.Cls

End Sub

Private Sub Command2_Click()
pic_paint.CurrentX = Text2.Text
pic_paint.CurrentY = Text3.Text
pic_paint.Print Msg; Text1.Text
End Sub




Private Sub cmd_save_Click()

    With comd_1
        .DialogTitle = "Choose a filename to save"
        .Filter = "Bitmap files (*.BMP)|*.BMP"
        .FilterIndex = 1
        .FileName = ""
        .ShowSave
        
        If .FileName = "" Then Exit Sub
        
        SavePicture pic_paint.Image, .FileName
    End With
    
End Sub

Private Sub Form_Load()

    For a% = 1 To 255
        For b% = 0 To 5
            pic_scale(0).Line (a% * 6, 0)-(a% * 6 + b%, pic_scale(0).Height), RGB(a%, 0, 0)
            pic_scale(1).Line (a% * 6, 0)-(a% * 6 + b%, pic_scale(1).Height), RGB(0, a%, 0)
            pic_scale(2).Line (a% * 6, 0)-(a% * 6 + b%, pic_scale(2).Height), RGB(0, 0, a%)
        Next
    Next

    line_mirror_vertical.x1 = pic_paint.ScaleWidth / 2
    line_mirror_vertical.x2 = pic_paint.ScaleWidth / 2
    
    Line_mirror_horisontal.y1 = pic_paint.ScaleHeight / 2
    Line_mirror_horisontal.y2 = pic_paint.ScaleWidth / 2
    
    
End Sub
Private Sub pic_color_Click()

    comd_1.ShowColor
    
    pic_color.BackColor = comd_1.color
    
    GetRgb (comd_1.color)
    
    scr_color(0).value = Red
    scr_color(1).value = Green
    scr_color(2).value = Blue
    
    shp_color(0).BackColor = RGB(Red, 0, 0)
    shp_color(1).BackColor = RGB(0, Green, 0)
    shp_color(2).BackColor = RGB(0, 0, Blue)
    
    lbl_color(Index).Caption = scr_color(Index).value

End Sub

Private Sub pic_paint_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If placing = True Then

        placed = True
        
        circleposx = x
        circleposy = y
        
        Call draw
    
    ElseIf chk_pen.value = 1 Then
        
        oldx = x
        oldy = y
        
        pic_paint.Line (oldx, oldy)-(x, y), RGB(Red, Green, Blue)
            
            If chk_mirror_vertical.value = 1 Then
                Call MirrorDraw(x, y, oldx, oldy, 0)
            End If
            
            If chk_mirror_horisontal.value = 1 Then
                Call MirrorDraw(x, y, oldx, oldy, 1)
            End If
            
            Call UpdateColor(scr_colorchange.value)
         
            If chk_autoupdate.value = 1 Then
                Call AutoUpdate
            End If
    End If

End Sub

Private Sub pic_paint_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If placing = True Then
        pic_paint.Cls
        pic_paint.Line (x, 0)-(x, pic_paint.Height)
        pic_paint.Line (0, y)-(pic_paint.Width, y)
    
    ElseIf chk_pen.value = 1 And Button = 1 Then
        
        pic_paint.Line (oldx, oldy)-(x, y), RGB(Red, Green, Blue)
        
        If chk_mirror_vertical.value = 1 Then
             Call MirrorDraw(x, y, oldx, oldy, 0)
        End If
        
        If chk_mirror_horisontal.value = 1 Then
            Call MirrorDraw(x, y, oldx, oldy, 1)
        End If
        
        If chk_autoupdate.value = 1 Then
            Call AutoUpdate
        End If
   
        
        oldx = x
        oldy = y
                    

        
        Call UpdateColor(scr_colorchange.value)
    
    End If

End Sub

Private Sub pic_scale_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    
    Select Case Index
        
        Case 0
            shp_color(0).BackColor = RGB(x / 6, 0, 0)
            scr_color(0).value = x / 6
        Case 1
            shp_color(1).BackColor = RGB(0, x / 6, 0)
            scr_color(1).value = x / 6
        Case 2
            shp_color(2).BackColor = RGB(0, 0, x / 6)
            scr_color(2).value = x / 6
            
    End Select
    
    pic_color.BackColor = RGB(scr_color(0).value, scr_color(1).value, scr_color(2).value)
    lbl_color(Index).Caption = scr_color(Index).value
    
    Red = scr_color(0).value
    Green = scr_color(1).value
    Blue = scr_color(2).value

End Sub

Private Sub scr_color_Change(Index As Integer)
    
    Select Case Index
        
        Case 0
            shp_color(Index).BackColor = RGB(scr_color(0), 0, 0)
        Case 1
            shp_color(Index).BackColor = RGB(0, scr_color(1), 0)
        Case 2
            shp_color(Index).BackColor = RGB(0, 0, scr_color(2))
        
    End Select
            
    pic_color.BackColor = RGB(scr_color(0).value, scr_color(1).value, scr_color(2).value)
    lbl_color(Index).Caption = scr_color(Index).value
    
    Red = scr_color(0).value
    Green = scr_color(1).value
    Blue = scr_color(2).value
    
End Sub

Private Sub scr_colorchange_Change()

    lbl_colorchange.Caption = scr_colorchange.value

End Sub

Private Sub scr_drawwidth_Change()

    pic_paint.DrawWidth = scr_drawwidth.value
    lbl_drawwidth.Caption = scr_drawwidth.value

End Sub

Private Sub txt_find_Change(Index As Integer)

    On Error Resume Next
    If Not IsNumeric(txt_find(Index).Text) Or txt_find(Index).Text < 0 Or txt_find(Index).Text > 255 Then
        txt_find(Index).Text = ""
    End If

End Sub
