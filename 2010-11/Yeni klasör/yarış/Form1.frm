VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4065
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Yarýþý Baþlat"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4920
      Top             =   1200
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   0
      MousePointer    =   1  'Arrow
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      Picture         =   "Form1.frx":E872B
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "Form1.frx":FCFA2
      Top             =   360
      Width           =   1500
   End
   Begin VB.Line Line1 
      X1              =   11520
      X2              =   11520
      Y1              =   360
      Y2              =   2880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Image1.Left = 0
Image2.Left = 0
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
Randomize

Image1.Left = Image1.Left + (Rnd * 200)
Image2.Left = Image2.Left + (Rnd * 200)

If Image1.Left >= Line1.X1 Then
Timer1.Enabled = False
a = MsgBox("Kýrmýzý takým kazandý")
End If

If Image2.Left >= Line1.X1 Then
Timer1.Enabled = False
a = MsgBox("Mavi takým kazandý")

End If


End Sub
