VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   2160
      ScaleHeight     =   915
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyUp
    Picture1.Top = Picture1.Top - 50
Case vbKeyDown
    Picture1.Top = Picture1.Top + 50
Case vbKeyLeft
    Picture1.Left = Picture1.Left - 50
Case vbKeyRight
    Picture1.Left = Picture1.Left + 50
End Select
End Sub


