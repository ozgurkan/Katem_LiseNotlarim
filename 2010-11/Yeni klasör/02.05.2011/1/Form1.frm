VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.ForeColor = vbRed
If Button = 1 Then Print "farenin sol tuþuna bastýnýz"
If Button = 2 Then Print "farenin sað tuþuna bastýnýz"
If Button = 4 Then Print "farenin orta tuþuna bastýnýz"
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.ForeColor = vbBlue
If Button = 1 Then Print "farenin sol tuþunu býraktýnýz"
If Button = 2 Then Print "farenin sað tuþunu býraktýnýz"
If Button = 4 Then Print "farenin orta tuþunu býraktýnýz"
End Sub
