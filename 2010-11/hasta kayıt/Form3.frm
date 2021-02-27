VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF0000&
   Caption         =   "Form3"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form3"
   ScaleHeight     =   4605
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "GÝRÝÞ SAYFASINA DÖN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "SATIN AL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text2 
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
      Left            =   3840
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
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
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "YENÝ ÞÝFRE===>"
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
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "YENÝ KULLANICI ADI===>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text1 = "" And Text2 = "" Then
MsgBox ("Lütfen boþ alan býrakmayýnýz")
ElseIf Text1 = "" Or Text2 = "" Then
MsgBox ("Lütfen boþ alan býrakmayýnýz")
Else
Set db = OpenDatabase(App.Path & "\hasta kayýt.mdb")
Set rs = db.OpenRecordset("þifre")
rs.AddNew
rs.Fields("adý") = Text1.Text
rs.Fields("þifre") = Text2.Text
rs.Update
rs.Close
MsgBox "kullanýcý adý ve þifre satýn alýndý."
Text1 = ""
Text2 = ""
End If
End Sub

Private Sub Command2_Click()
Form3.Hide
Form2.Show
End Sub
