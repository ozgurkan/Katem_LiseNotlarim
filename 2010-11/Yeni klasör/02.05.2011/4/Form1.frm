VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.ToolTipText = "SHIFT+CTRL+ALT tu�lar�na basarak FARE �LE BU ALANA TIKLAYAB�L�RS�N�Z"
End Sub


Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = 1 Then Text1.Text = "shift'e bas�ld�"
If Shift = 2 Then Text1.Text = "ctrl'e bas�ld�"
If Shift = 4 Then Text1.Text = "Alt'a bas�ld�"
If Shift = 3 Then Text1.Text = "Shift+Ctrl'ye bas�ld�"
If Shift = 7 Then Text1.Text = "Alt+Shift+Ctrl'ye bas�ld�"
If Shift = 6 Then Text1.Text = "Ctrl+Alt'a bas�ld�"
If Shift = 5 Then Text1.Text = "Shift+Alt'a bas�ld�"
End Sub
