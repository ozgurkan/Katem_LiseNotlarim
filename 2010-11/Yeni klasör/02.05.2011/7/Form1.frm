VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5040
      Top             =   2880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SARS"
      Height          =   735
      Left            =   2160
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      Height          =   1815
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sars, x As Integer

Private Sub Command1_Click()
Timer1.Enabled = True
sars = 1
x = 150
End Sub

Private Sub Timer1_Timer()
sars = sars + 1
If sars = 31 Then
MsgBox "sarsma baba yorgun"
Timer1.Enabled = False
End If
Shape1.Left = Shape1.Left + x
x = -x

End Sub
