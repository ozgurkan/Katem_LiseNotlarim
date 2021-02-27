VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6345
   DrawWidth       =   5
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2880
      Top             =   3240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Show
End Sub

Private Sub Form_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
Timer1.Enabled = True

End If
End Sub

Private Sub Timer1_Timer()
Randomize
For sayac = 1 To 25
a = a + 100
Line (5000 - a, 5000 - a)-(X + a, Y + a), RGB(Rnd * 265, Rnd * 256, Rnd * 256), B

Next sayac
End Sub
