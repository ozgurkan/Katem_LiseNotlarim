VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Show
Randomize
FillStyle = 1
FillColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Randomize
If Button = 1 Then
For sayac = 1 To 5
Circle (X, Y), 100 * sayac, RGB(Rnd * 255, Rnd * 255, Rnd * 255)

Next sayac
End If
End Sub

