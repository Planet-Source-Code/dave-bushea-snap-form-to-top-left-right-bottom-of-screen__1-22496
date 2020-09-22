VERSION 5.00
Begin VB.Form Snap 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   885
   ClientLeft      =   6300
   ClientTop       =   3975
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   885
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.Timer snaptimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snaping example by Dave Bushea http://www.rapta.net"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label lbldrag 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Drag Me i'll snap!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5640
   End
End
Attribute VB_Name = "Snap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub lbldrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mouseY = Y
mouseX = X
snaptimer.Enabled = True
End Sub

Private Sub lbldrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
snaptimer.Enabled = False
End Sub

Private Sub snaptimer_Timer()
Dim curpos As POINTAPI
Dim pixelsnap As Integer
retval = GetCursorPos(curpos)
'amount of pixels to snap to
pixelsnap = 30

'snap to right
If (curpos.X - (mouseX / Screen.TwipsPerPixelX)) <= pixelsnap Then
curpos.X = mouseX / Screen.TwipsPerPixelX
End If

'snap to top
If (curpos.Y - (mouseY / Screen.TwipsPerPixelY)) <= pixelsnap Then
curpos.Y = mouseY / Screen.TwipsPerPixelY
End If

'stap to left
If (curpos.X - (mouseX / Screen.TwipsPerPixelX)) >= ((Screen.Width / Screen.TwipsPerPixelX) - (Me.Width / Screen.TwipsPerPixelX)) - pixelsnap Then
curpos.X = Screen.Width / Screen.TwipsPerPixelX - (Me.Width / Screen.TwipsPerPixelX) + (mouseX / Screen.TwipsPerPixelX)
End If

'snap to bottom
If (curpos.Y - (mouseY / Screen.TwipsPerPixelY)) >= ((Screen.Height / Screen.TwipsPerPixelY) - (Me.Height / Screen.TwipsPerPixelY)) - pixelsnap Then
curpos.Y = Screen.Height / Screen.TwipsPerPixelY - (Me.Height / Screen.TwipsPerPixelY) + (mouseY / Screen.TwipsPerPixelY)
End If

'moves the form
Me.Top = (curpos.Y * Screen.TwipsPerPixelY) - mouseY
Me.Left = (curpos.X * Screen.TwipsPerPixelX) - mouseX
End Sub
