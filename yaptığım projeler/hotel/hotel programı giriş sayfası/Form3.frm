VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   9255
   ClientLeft      =   1575
   ClientTop       =   990
   ClientWidth     =   12465
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   12465
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   9255
      Left            =   0
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   9195
      ScaleWidth      =   12435
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080C0FF&
         Caption         =   "AYRILAN M��TER�LER"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ayr�lan m��terilere ula�mak i�in t�klay�n."
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0080C0FF&
         Caption         =   "ODA DURUMU"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Oda durumuna ula�mak i�in t�klay�n."
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   4200
         Top             =   840
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "M��TER� ARA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Otelde bulunan m��terileri aramak i�in t�klay�n."
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "M��TER� KAYIT"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "M��teri giri�i i�in t�klay�n."
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "M��TER� AYRILI�"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "M��teri ��k��� i�in t�klay�n."
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080C0FF&
         Caption         =   "M��TER� B�LG�LER�"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "M��teri bilgilerine ula�mak i�in t�klay�n."
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080C0FF&
         Caption         =   "�IKI�"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Program� kapatmak i�in t�klay�n."
         Top             =   7200
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0080C0FF&
         Caption         =   "G�R�� SAYFASINA D�N"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Giri� sayfas�na d�nmek i�in t�klay�n."
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   " ANA SAYFA "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   4800
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   6240
         X2              =   6000
         Y1              =   840
         Y2              =   1200
      End
   End
   Begin VB.Menu program 
      Caption         =   "PROGRAM"
      Begin VB.Menu de�i�tir 
         Caption         =   "KULLANICI DE���T�R"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu m��teri 
      Caption         =   "M��TER�LER"
      Begin VB.Menu kay�t 
         Caption         =   "M��TER� KAYIT"
         Shortcut        =   ^K
      End
      Begin VB.Menu ayr�l�� 
         Caption         =   "M��TER� AYRILI�"
         Shortcut        =   ^A
      End
      Begin VB.Menu bilgileri 
         Caption         =   "M��TER� B�LG�LER�"
         Shortcut        =   ^B
      End
      Begin VB.Menu ara 
         Caption         =   "M��TER� ARA"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu oda 
      Caption         =   "ODALAR"
      Begin VB.Menu DURUMU 
         Caption         =   "ODA DURUMU"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu yard�m 
      Caption         =   "YARDIM"
      Index           =   1
      Begin VB.Menu YARDIM 
         Caption         =   "YARDIM"
         Shortcut        =   ^Y
      End
      Begin VB.Menu hakk�nda 
         Caption         =   "HAKKINDA"
         Shortcut        =   ^H
      End
      Begin VB.Menu ��k�� 
         Caption         =   "�IKI�"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub ara_Click()
Form3.Hide
Form7.Show
End Sub

Private Sub ayr�l��_Click()
Form3.Hide
Form6.Show
End Sub

Private Sub bilgileri_Click()
Form3.Hide
Form4.Show
End Sub

Private Sub Command1_Click()
Form7.Show
Form3.Hide
End Sub
Private Sub Command2_Click()
Form3.Hide
Form5.Show
End Sub
Private Sub Command3_Click()
Form3.Hide
Form6.Show
End Sub
Private Sub Command4_Click()
Form3.Hide
Form4.Show
End Sub
Private Sub Command5_Click()
cevap = MsgBox("programdan ��kmak istiyormusunuz?", 36, "onay butonu")
If cevap = 6 Then
End
Else
MsgBox "��k�� i�lemi iptal edildi."
End If
End Sub
Private Sub Command6_Click()
Form3.Hide
Form1.Show
End Sub

Private Sub Command7_Click()
Form9.Show
Form3.Hide

End Sub

Private Sub ��k��_Click()
cevap = MsgBox("programdan ��kmak istiyormusunuz?", 36, "onay butonu")
If cevap = 6 Then
End
Else
MsgBox "��k�� i�lemi iptal edildi."
End If
End Sub

Private Sub de�i�tir_Click()
Form3.Hide
Form1.Show
Form1.Text1.Text = ""
Form1.Text2.Text = ""
End Sub

Private Sub DURUMU_Click()
Form3.Hide
Form9.Show
End Sub

Private Sub Form_Load()
Label2.Caption = Date
Label3.Caption = Time
End Sub


Private Sub hakk�nda_Click()
Form8.Show
End Sub

Private Sub kay�t_Click()
Form3.Hide
Form5.Show
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Date
Label3.Caption = Time
End Sub

Private Sub YARDIM_Click()
On Error Resume Next
Shell "cmd /c notepad.exe " & "C:\hotel\hotel program� giri� sayfas�\yard�m"
End Sub
