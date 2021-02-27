VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1575
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label1.FontSize = 36
Label1.BorderStyle = 0

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Label1.ForeColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
If Button = 2 Then Form1.BackColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
End Sub

