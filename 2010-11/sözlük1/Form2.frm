VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10335
   LinkTopic       =   "Form2"
   ScaleHeight     =   8475
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   240
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   1785
      ScaleWidth      =   9585
      TabIndex        =   5
      Top             =   240
      Width           =   9615
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   7440
         Picture         =   "Form2.frx":3916E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   8160
         Picture         =   "Form2.frx":39BC0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Height          =   495
         Left            =   8760
         Picture         =   "Form2.frx":3A5DA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "KELÝMEYÝ SÖZLÜÐE KAYDETMEK ÝÇÝN TIKLAYIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   4
      Top             =   7080
      Width           =   6015
   End
   Begin VB.TextBox Text2 
      Height          =   2895
      Left            =   4440
      TabIndex        =   3
      Top             =   3840
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   2400
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   " ANLAMI GÝRÝNÝZ    =>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "KELÝMEYÝ GÝRÝNÝZ  =>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
cevap = MsgBox("Kelime ekleme alanýný kapatmak istiyormusunuz?", 36, "onay kutusu")
If cevap = 6 Then
MsgBox ("Kelime ekleme alaný kapanýyor.")
Cancel = False
Form2.Hide
Form1.Show
ElseIf cevap = 7 Then
MsgBox ("Kelime ekleme alaný kapatýlamadý.")
Cancel = True
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Text1 = "" And Text2 = "" Then
MsgBox ("Lütfen boþ alan býrakmayýnýz")
ElseIf Text1 = "" Or Text2 = "" Then
MsgBox ("Lütfen boþ alan býrakmayýnýz")
Else
Set db = OpenDatabase(App.Path & "\sözlük.mdb")
Set Rs = db.OpenRecordset("tablo")
Rs.AddNew
Rs.Fields("kelime") = Text1.Text
Rs.Fields("anlam") = Text2.Text
Form1.List1.AddItem Text1
Rs.Update
Rs.Close
MsgBox ("kelimeniz sözlüðe eklenmiþtir.")
Form2.Hide
Form1.Show
s = s + 1
End If
End Sub

