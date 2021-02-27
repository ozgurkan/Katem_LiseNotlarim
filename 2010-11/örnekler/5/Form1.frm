VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As Integer

Dim a As Double
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub
Private Sub Form_Load()
DrawWidth = 5
a = 0
b = 0
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
a = a + 0.1
Circle (5000, 3500), b, , , a
If a = 6 Then
Circle (5000, 3500), 1000, , , 6.4
b = b + 50

a = 0
End If
End Sub

