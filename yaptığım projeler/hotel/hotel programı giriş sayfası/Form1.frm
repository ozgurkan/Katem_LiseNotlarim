VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   1710
   ClientTop       =   645
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   9255
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   9195
      ScaleWidth      =   11475
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   11040
         Picture         =   "Form1.frx":155E9E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "çýkmak için týklayýn."
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "YENÝ KULLANICI ADI VE PAROLA AL"
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Kullanýcý adý ve þifre almak için týklayýn."
         Top             =   6000
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "TAMAM"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Programa giriþ yapmak için týklayýn."
         Top             =   6000
         Visible         =   0   'False
         Width           =   1695
      End
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
         IMEMode         =   3  'DISABLE
         Left            =   5280
         PasswordChar    =   "*"
         TabIndex        =   4
         ToolTipText     =   "parola giriniz"
         Top             =   5280
         Visible         =   0   'False
         Width           =   3735
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
         Left            =   5280
         TabIndex        =   3
         ToolTipText     =   "kullanýcý adý giriniz"
         Top             =   4560
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   480
         Top             =   6960
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   735
         Left            =   1200
         TabIndex        =   2
         Top             =   7560
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   1296
         _Version        =   393216
         Appearance      =   1
         Max             =   20
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Parola giriniz===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   11
         Top             =   5280
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Kullanýcý adý giriniz===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   10
         Top             =   4560
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "HOÞ GELDÝNÝZ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   9
         Top             =   3360
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   8400
         Width           =   6015
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "KULLANICI ADINI VEYA ÞÝFREYÝ YANLIÞ GÝRDÝNÝZ."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   7080
         Visible         =   0   'False
         Width           =   7215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "    OTEL  TAKÝP     PROGRAMI"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   24
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1395
         Left            =   2520
         TabIndex        =   1
         Top             =   120
         Width           =   6465
      End
   End
End
Attribute VB_Name = "Form1"
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
Set db = OpenDatabase("c:\hotel\hotel programý giriþ sayfasý\þifre.mdb")
Set Rs = db.OpenRecordset("tablo", dbOpenSnapshot)
a = Rs.RecordCount
For sayac = 0 To a
If Text1.Text <> Rs!Adý Or Text2.Text <> Rs!þifre Then
Label6.Visible = True
Rs.MoveNext
ElseIf Text1.Text = Rs!Adý And Text2.Text = Rs!þifre Then
Label6.Visible = False
Form1.Hide
Form3.Show
End If
Next sayac
Rs.Close
End If
End Sub

Private Sub Command2_Click()
Form1.Hide
Form2.Show
Form2.Text2.Text = ""
Form2.Text1.Text = ""
Text1 = ""
Text2 = ""
End Sub

Private Sub Command3_Click()
cevap = MsgBox("programdan çýkmak istiyormusunuz?", 36, "onay butonu")
If cevap = 6 Then
End
Else
MsgBox "çýkýþ iþlemi iptal edildi."
End If
End Sub

Private Sub Text1_Click()
Text1.BackColor = vbYellow
End Sub
Private Sub Text2_Click()
Text2.BackColor = vbYellow
End Sub

Private Sub Timer1_Timer()
'Timer1.Interval = Timer1.Interval + 100
ProgressBar1.Value = ProgressBar1.Value + 1
Label1.Caption = "program yükleniyor    %" & ProgressBar1.Value * 5
If ProgressBar1.Value = 20 Then
Timer1.Enabled = False
Label1.Caption = "program yüklendi..... "
ProgressBar1.Visible = False
Label1.Visible = False
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Text1.Visible = True
Text2.Visible = True
Command1.Visible = True
Command2.Visible = True

End If
End Sub

