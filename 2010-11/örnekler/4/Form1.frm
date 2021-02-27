VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "temizle"
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "tam"
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   3960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "üç çeyrek"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "yarým"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "çeyrek"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   5000
      X2              =   5015
      Y1              =   3500
      Y2              =   3515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pi As Integer

Private Sub Command1_Click()



Circle (5000, 3500), 500, , , 1.57

End Sub

Private Sub Command2_Click()

Circle (5000, 3500), 500, , 1.57, pi
End Sub

Private Sub Command3_Click()

Circle (5000, 3500), 500, , pi, 3 * pi / 2
End Sub

Private Sub Command4_Click()

Circle (5000, 3500), 500, , 4.5, 0
End Sub

Private Sub Form_Load()
pi = 3.14
End Sub
