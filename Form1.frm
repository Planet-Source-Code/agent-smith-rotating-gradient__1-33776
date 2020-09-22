VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "James's sexy gradient"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   613
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   17
      Top             =   720
      Width           =   9015
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   15
      Left            =   4080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H009504FF&
      Height          =   255
      Index           =   14
      Left            =   3840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00C004FF&
      Height          =   255
      Index           =   13
      Left            =   3600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Index           =   12
      Left            =   3360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FF04B4&
      Height          =   255
      Index           =   11
      Left            =   3120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   2880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FF9B04&
      Height          =   255
      Index           =   9
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Index           =   8
      Left            =   2400
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0000FF00&
      Height          =   255
      Index           =   7
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0004FF8E&
      Height          =   255
      Index           =   6
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0004FFBA&
      Height          =   255
      Index           =   5
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0004DAFF&
      Height          =   255
      Index           =   3
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0004A7FF&
      Height          =   255
      Index           =   2
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H000482FF&
      Height          =   255
      Index           =   1
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CDbox 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar scrCount 
      Height          =   255
      Left            =   480
      Max             =   15
      Min             =   1
      TabIndex        =   1
      Top             =   120
      Value           =   15
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RGBset
    Angle As Integer
    r(0 To 15)
    g(0 To 15)
    b(0 To 15)
    Count As Integer
End Type
Dim gradtemp As RGBset
Dim mdown As Boolean

Private Sub Form_Load()

    gradtemp.Count = scrCount.Value
    For i = 0 To scrCount.Value
        gradtemp.r(i) = Red(picColors(i).BackColor)
        gradtemp.g(i) = Green(picColors(i).BackColor)
        gradtemp.b(i) = Blue(picColors(i).BackColor)
    Next i

End Sub

Private Sub picColors_Click(Index As Integer)

    On Error GoTo ChoseCancel
    CDbox.Color = picColors(Index).BackColor
    CDbox.ShowColor
    picColors(Index).BackColor = CDbox.Color
    gradtemp.r(Index) = Red(CDbox.Color)
    gradtemp.g(Index) = Green(CDbox.Color)
    gradtemp.b(Index) = Blue(CDbox.Color)
    DrawGrad picMain, gradtemp
ChoseCancel:

End Sub

Private Sub picMain_Click()
    
    DrawGrad picMain, gradtemp

End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mdown = True

End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mdown = True Then
        gradtemp.Angle = GetAngle(picMain.ScaleWidth / 2, picMain.ScaleHeight / 2, X, Y)
        picMain.Cls
        'DrawGrad picMain, gradtemp
        DrawAngleMark picMain
    End If

End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mdown = False
    picMain.Cls
    DrawGrad picMain, gradtemp

End Sub

Private Sub scrCount_Change()

    gradtemp.Count = scrCount
    picMain.Cls
    DrawGrad picMain, gradtemp

End Sub

Private Function Red(ByVal Color As Long) As Integer
    Red = Color Mod &H100
End Function
Private Function Green(ByVal Color As Long) As Integer
    Green = (Color \ &H100) Mod &H100
End Function
Private Function Blue(ByVal Color As Long) As Integer
    Blue = (Color \ &H10000) Mod &H100
End Function

Private Sub DrawGrad(obj As Object, rgbs As RGBset)

    rang = rgbs.Angle - 180
    MainWidth = obj.ScaleWidth
    MainHeight = obj.ScaleHeight
    obj.DrawWidth = 2
    r2 = rgbs.r(0)
    g2 = rgbs.g(0)
    b2 = rgbs.b(0)
    h = Int(obj.ScaleHeight / rgbs.Count)
    For c = 0 To rgbs.Count - 1
        rd = (rgbs.r(c + 1) - r2) / h
        gd = (rgbs.g(c + 1) - g2) / h
        bd = (rgbs.b(c + 1) - b2) / h
        For Y = 0 To h - 1
            r2 = r2 + rd
            g2 = g2 + gd
            b2 = b2 + bd
            
            rad_ang = ((rang - 90) / 360 * 6.28318)
            cx = MainWidth / 2 + Cos(rad_ang) * ((c * h + Y) - MainHeight / 2)
            cy = MainHeight / 2 + Sin(rad_ang) * ((c * h + Y) - MainHeight / 2)
            
            rad_ang = ((rang - 180) / 360 * 6.28318)
            X1 = cx + Cos(rad_ang) * MainWidth / 2
            Y1 = cy + Sin(rad_ang) * MainWidth / 2
            
            rad_ang = (rang / 360 * 6.28318)
            X2 = cx + Cos(rad_ang) * MainWidth / 2
            Y2 = cy + Sin(rad_ang) * MainWidth / 2
            
            obj.Line (X1, Y1)-(X2, Y2), RGB(r2, g2, b2)
            'obj.Line (0, cx)-(obj.Width, c * h + Y), RGB(r2, g2, b2)
        Next Y
    Next c
    obj.DrawWidth = 1

End Sub

Private Sub DrawAngleMark(obj As Object)

    obj.DrawMode = 6
    rad_ang = ((gradtemp.Angle - 90) / 360 * 6.28318)
    INX = obj.ScaleWidth / 2 + Cos(rad_ang) * 10
    INY = obj.ScaleHeight / 2 + Sin(rad_ang) * 10
    OUTX = obj.ScaleWidth / 2 + Cos(rad_ang) * 30
    OUTY = obj.ScaleHeight / 2 + Sin(rad_ang) * 30
    obj.Circle (obj.ScaleWidth / 2, obj.ScaleHeight / 2), 5
    obj.Line (INX, INY)-(OUTX, OUTY)
    obj.DrawMode = 13

End Sub

Private Function GetAngle(X1, Y1, X2, Y2) As Currency

    XD = X2 - X1
    YD = Y2 - Y1
    
    On Error Resume Next
    ang = Atn(YD / XD) * (180 / 3.14159)
    
    If X2 < X1 And Y2 < Y1 Then ang = ang + 270
    If X2 > X1 And Y2 < Y1 Then ang = 90 - -ang
    If X2 < X1 And Y2 > Y1 Then ang = 90 - -ang + 180
    If X2 > X1 And Y2 > Y1 Then ang = ang + 90
    
    If X2 = X1 And Y2 > Y1 Then ang = 180
    If Y2 = Y1 And X2 > X1 Then ang = 90
    If Y2 = Y1 And X2 < X1 Then ang = 270
    If X2 = X1 And Y2 < Y1 Then ang = 0
    
    GetAngle = ang

End Function
