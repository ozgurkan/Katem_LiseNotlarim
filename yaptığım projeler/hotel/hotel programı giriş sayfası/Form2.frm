VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5850
   ClientLeft      =   3195
   ClientTop       =   2385
   ClientWidth     =   7920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   5835
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         TabIndex        =   4
         Top             =   3360
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         TabIndex        =   3
         Top             =   2040
         Width           =   4095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Giri� Sayfas�na D�n"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Giri� sayfas�na d�nmek i�in t�klay�n."
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "KULLANICI ADI VE PAROLA SATIN AL"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Sat�n almak i�in t�klay�n."
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Parola giriniz===>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Kullan�c� ad� giriniz===>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text1 = "" And Text2 = "" Then
MsgBox ("L�tfen bo� alan b�rakmay�n�z")
ElseIf Text1 = "" Or Text2 = "" Then
MsgBox ("L�tfen bo� alan b�rakmay�n�z")
Else
Set db = OpenDatabase("c:\hotel\hotel program� giri� sayfas�\�ifre.mdb")
Set Rs = db.OpenRecordset("tablo", dbOpenSnapshot)
Rs.AddNew
Rs.Fields("ad�") = Text1.Text
Rs.Fields("�ifre") = Text2.Text
Rs.Update
Rs.Close
MsgBox "kullan�c� ad� ve �ifre sat�n al�nd�."
Text1 = ""
Text2 = ""
End If
End Sub
Private Sub Command2_Click()
On Error Resume Next
Form1.Show
Form2.Hide
End Sub
Private Sub Form_Load()
Text1.Text = Clear
Text2.Text = Clear
End Sub

