VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   13716
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   13716
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CurrentX = X
CurrentY = Y
If Button = 1 Then
Print "GÜL&ÖZGÜR"
End If
End Sub
