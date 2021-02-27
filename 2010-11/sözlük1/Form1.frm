VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Height          =   495
      Left            =   5280
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Timer Timer2 
      Left            =   1920
      Top             =   8160
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   2040
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2520
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   8040
      Picture         =   "Form1.frx":2F86
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "yeni kelime eklemek için týklayýn"
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   6960
      Picture         =   "Form1.frx":4590
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "kelime aramak için týklayýn"
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   8040
      Picture         =   "Form1.frx":58F2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "kelimenin anlamýný yazdýrmak için týklayýn"
      Top             =   7560
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      ItemData        =   "Form1.frx":7E54
      Left            =   0
      List            =   "Form1.frx":7E56
      TabIndex        =   3
      Top             =   3240
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2520
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   240
      Picture         =   "Form1.frx":7E58
      ScaleHeight     =   1785
      ScaleWidth      =   9585
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.CommandButton Command6 
         Height          =   495
         Left            =   8760
         Picture         =   "Form1.frx":4047A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "çýkýþ yapmak için týklayýn"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Height          =   495
         Left            =   8160
         Picture         =   "Form1.frx":4111C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "simge durumunda küçültmek için týklayýn"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Height          =   495
         Left            =   7440
         Picture         =   "Form1.frx":41C9E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "sözlük hakkýnda bilgi almak için týklayýn"
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00004040&
      Caption         =   "programý hazýrlayan=özgür kan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   8160
      Width           =   4815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Lütfen kelimeyi küçük harfle yazýn!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   2040
      Width           =   8055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   4560
      TabIndex        =   8
      Top             =   3240
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Ara:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command2_Click()
Set db = OpenDatabase("c:\sözlük1\sözlük.mdb")
Set Rs = db.OpenRecordset("tablo", dbOpenSnapshot)
Do While Not Rs.EOF
If Text1.Text = Rs!kelime Then
Text1.Text = "aradýðýnýz kelime sözlüðümüzde bulundu"
Text3.Text = Rs!kelime
Else:
Rs.MoveNext
End If
Loop
If Text1.Text <> "aradýðýnýz kelime sözlüðümüzde bulundu" Then Text1.Text = "aradýðýnýz kelime sözlüðümüzde bulunamadi"
If Text1.Text = "aradýðýnýz kelime sözlüðümüzde bulundu" Then
For sayac = Rs.MoveFirst To Rs.EOF
For i = 0 To List1.ListCount
If Text3.Text = Rs!kelime Then
a = Rs!anlam
Else
Rs.MoveNext
End If
Next i
Next sayac
Label2.Caption = a
Else
Text1.Text = "sözlüðümüzde bulunmayan bir kelime girdiniz"
Text3.Text = "aradýðýnýz kelime bulunamadý."
End If
End Sub


Private Sub Command3_Click()
Form1.Hide
Form2.Show
Form2.Text1.Text = ""
Form2.Text2.Text = ""
End Sub

Private Sub Command5_Click()
Form1.WindowState = vbMinimized
End Sub

Private Sub command6_click()
cevap = MsgBox("Programdan çýkmak istiyormusunuz?", 36, "onay kutusu")
If cevap = 6 Then
MsgBox ("Çýkýþ iþlemi yapýlýyor.")
Cancel = False
End
ElseIf cevap = 7 Then
MsgBox ("Çýkýþ iþlemi gerçekleþtirilemedi.")
Cancel = True
End If

End Sub

Private Sub Command7_Click()
Call ShellExecute(&O0, vbNullString, "www.google.com", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim db As Database
Dim Rs As Recordset
Set db = OpenDatabase("c:\sözlük1\sözlük.mdb")
Set Rs = db.OpenRecordset("tablo", dbOpenSnapshot)
For sayac = 0 To 100
If Rs!kelime <> "" And Rs!anlam <> "" Then
a = Rs!kelime
List1.AddItem a
Rs.MoveNext
End If
Next sayac
On Error Resume Next
For i = 0 To 100
If List1.List(i) = "" Then
List1.RemoveItem i
End If
Next i



Label3.Caption = "                                                  Lütfen kelimeyi küçük harfle yazýn!!!!   "
Label3.Left = ScaleLeft
Label3.Width = Me.ScaleWidth
Timer1.Interval = 100

Label4.Caption = "                                                   programý hazýrlayan=özgür kan       "
Label4.Left = ScaleLeft
Label4.Width = Me.ScaleWidth
Timer2.Interval = 100


End Sub

Private Sub List1_Click()
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
Text1 = List1.List(i)
End If
Next i
End Sub

Private Sub Text1_Change()
Label2.Caption = clean
Text3.Text = clean
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label3 = Mid(Label3, 2, Len(Label3) - 1) + Left(Label3, 1)
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Label4 = Mid(Label4, 2, Len(Label4) - 1) + Left(Label4, 1)
End Sub
