VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Randomize
If Button = 1 Then
For sayac = 1 To 5
a = a + 100
Line (X - a, Y - a)-(X + a, Y + a), RGB(Rnd * 265, Rnd * 256, Rnd * 256), B

Next sayac
End If
End Sub
