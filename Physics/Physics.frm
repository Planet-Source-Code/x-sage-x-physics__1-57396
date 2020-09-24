VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Physics"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   2280
      Picture         =   "Physics.frx":0000
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   8
      Tag             =   "0"
      Top             =   360
      Width           =   585
   End
   Begin VB.PictureBox Ball 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   480
      Picture         =   "Physics.frx":0ECA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Tag             =   "0"
      Top             =   600
      Width           =   240
   End
   Begin VB.PictureBox Picture5 
      Height          =   975
      Left            =   3600
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   7
      Top             =   360
      Width           =   975
      Begin VB.Line lnWind 
         X1              =   -120
         X2              =   360
         Y1              =   240
         Y2              =   360
      End
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1920
      Picture         =   "Physics.frx":120C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1680
      Picture         =   "Physics.frx":154E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1920
      Picture         =   "Physics.frx":1890
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1680
      Picture         =   "Physics.frx":1BD2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer Timer2 
      Left            =   3240
      Top             =   1080
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   19
      SelStart        =   4
      Value           =   4
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2760
      Top             =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bouncyness:"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4560
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Force As Integer
Dim Wind As Integer
Dim Bounc As Integer
Dim Bounce As String
Dim PiC As Integer
Dim bForce As Integer
Dim bWind As Integer


Private Sub Form_Load()
MsgBox "To move the ball around click and hold in the box and let go to make the ball move in that direction."
lnWind.X1 = Picture5.Width / 2
lnWind.Y1 = Picture5.Height / 2
lnWind.X2 = lnWind.X1 + 1000
lnWind.Y2 = lnWind.Y1
Bounce = 0.3
Wind = (lnWind.X1 - lnWind.X2) / 15
Force = (lnWind.Y1 - lnWind.Y2) / 15
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lnWind.X2 = X
lnWind.Y2 = Y
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 0 Then
lnWind.X2 = X
lnWind.Y2 = Y
End If
End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lnWind.X2 = X
lnWind.Y2 = Y
Wind = (lnWind.X1 - lnWind.X2) / 15
Force = (lnWind.Y2 - lnWind.Y1) / 15
End Sub

Private Sub Timer1_Timer()
Timer2.Interval = Abs(Wind) * 5

If Box.Left + Box.Width > Me.Width Then
bWind = -(bWind * Bounce)
Box.Left = Me.Width - Box.Width
End If

If Box.Left < 0 Then
bWind = -(bWind * Bounce)
Box.Left = 0
End If

If Ball.Left + Ball.Width > Box.Left And Ball.Left < Box.Left + Box.Width And Ball.Top + Ball.Height > Box.Top And Ball.Top < Box.Top + Box.Height Then

If Box.Top + Box.Height < Ball.Top + Box.Height / 8 Then
bForce = -(bForce * Bounce)
Box.Top = Ball.Top - Box.Height - 1
Exit Sub
End If

If Ball.Top + Ball.Height < Box.Top + Box.Height / 3 Then
Force = -(Force * Bounce)
If Wind > 0 Then
Wind = Wind - 1
End If
If Wind < 0 Then
Wind = Wind + 1
End If

Ball.Top = Box.Top - Ball.Height - 1
Exit Sub
End If


'Hit the LEFT side;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
If Abs(Ball.Left + Ball.Width - Box.Left) < Abs(Ball.Left - Box.Left - Box.Width) Then
Ball.Left = Box.Left - Ball.Width - 2
Wind = -(Wind * Bounce) / 2
bWind = bWind + Wind * 3
If Force > 0 Then
bForce = (bForce + Force) * 2
Else
bForce = (bForce - Force) / 2
End If
Exit Sub
End If


'Hits the RIGHT side;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
If Ball.Left < Box.Left + Box.Width Then
Ball.Left = Box.Left + Box.Width + 2
Wind = -(Wind * Bounce)
bWind = bWind + Wind * 3
bForce = -((bForce + Force) * 0.2)
If Force > 0 Then
bForce = (bForce + Force) * 2
Else
bForce = (bForce - Force) / 2
End If
Exit Sub
End If

End If

If Box.Top + Box.Height >= Line1.Y1 Then
bForce = -(bForce * Bounce)
Box.Top = Line1.Y1 - Box.Height
Else
bForce = bForce + 1
End If

Box.Top = Box.Top + bForce
Box.Left = Box.Left + bWind

If bWind > 0 Then
bWind = bWind - 1
End If
If bWind < 0 Then
bWind = bWind + 1
End If
If Slider1.Value = 10 Then Slider1.Value = 9
If Slider1.Value < 10 Then
Bounce = "0." & Slider1.Value
Else
Bounce = Slider1.Value
Bounce = Right(Bounce, Len(Bounce) - 1)
End If
If Ball.Left + Ball.Width > Me.Width Then
Ball.Left = Me.Width - Ball.Width
Wind = -(Wind * Bounce)
End If
If Ball.Left < 0 Then
Ball.Left = 0
Wind = -(Wind * Bounce)
End If
Force = Force + 1
Ball.Left = Ball.Left - Wind
Ball.Top = Ball.Top + Force

If Ball.Top + Ball.Height >= Line1.Y1 Then
If Force <= 1 Then
Force = 0
If Wind > 0 Then
Wind = Wind - 1
End If
If Wind < 0 Then
Wind = Wind + 1
End If
Ball.Top = Line1.Y1 - Ball.Height
Exit Sub
End If

Ball.Top = Line1.Y1 - Ball.Height
Force = -(Force * Bounce)
End If
End Sub

Private Sub Timer2_Timer()
If PiC = 1 Then
PiC = 2
Ball.Picture = Picture3.Picture
Exit Sub
End If
If PiC = 2 Then
PiC = 3
Ball.Picture = Picture4.Picture
Exit Sub
End If
If PiC = 3 Then
PiC = 0
Ball.Picture = Picture1.Picture
Exit Sub
End If
If PiC = 0 Then
PiC = 1
Ball.Picture = Picture2.Picture
End If
End Sub
