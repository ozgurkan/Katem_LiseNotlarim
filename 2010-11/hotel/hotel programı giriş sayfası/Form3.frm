VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ANA SAYFA"
   ClientHeight    =   9255
   ClientLeft      =   1575
   ClientTop       =   870
   ClientWidth     =   12465
   FillColor       =   &H00404040&
   ForeColor       =   &H00404040&
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
         Top             =   960
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
         Caption         =   "ANA SAYFA "
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
         Top             =   960
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
         Top             =   960
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   6240
         X2              =   6000
         Y1              =   960
         Y2              =   1320
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
      Begin VB.Menu ayr�lan 
         Caption         =   "AYRILAN M��TER�LER"
         Shortcut        =   ^G
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

Private Sub ayr�lan_Click()
On Error Resume Next
Set db = OpenDatabase(App.Path & "\�ifre.mdb")
Set rs = db.OpenRecordset("tablo2")
rs.MoveFirst
Form10.Label2.Caption = rs!OdaNo
Form10.Label4.Caption = rs!Ad�
Form10.Label6.Caption = rs!soyad�
Form10.Label8.Caption = rs!Tc
Form10.Label10.Caption = rs!�kametgah_Adresi
Form10.Label12.Caption = rs!Telefon
Form10.Label14.Caption = rs!Geli�_Tarihi
Form10.Label16.Caption = rs!Geli�_Saati
Form10.Label18.Caption = rs!Ayr�l��_Tarihi
Form10.Label20.Caption = rs!Ayr�l��_Saati
Form10.Label22.Caption = rs!G�n
Form10.Label24.Caption = rs!Fiyat
rs.Update
rs.Close
If Form10.Label2.Caption = "" And Form10.Label8.Caption = "" Then
MsgBox "OTEL�M�ZDEN M��TER� AYRILMAMI�TIR."
Form10.Hide
Form3.Show
Else
Form3.Hide
Form10.Show
End If
End Sub

Private Sub ayr�l��_Click()
Form3.Hide
Form6.Show
End Sub

Private Sub bilgileri_Click()
On Error Resume Next
If Form4.Text5.Text = "" And Form4.Text13.Text = "" Then
MsgBox "OTEL�M�ZDE KAYITLI M��TER� YOKTUR."
Form4.Hide
Form3.Show
Else
Form3.Hide
Form4.Show
End If
Set db = OpenDatabase(App.Path & "\�ifre.mdb")
Set rs = db.OpenRecordset("tablo1")
rs.MoveFirst
Form4.Text1.Text = rs!Ad�
Form4.Text2.Text = rs!soyad�
Form4.Text3.Text = rs!Baba_ad�
Form4.Text4.Text = rs!Anne_ad�
Form4.Text5.Text = rs!Tc
Form4.Text6.Text = rs!il
Form4.Text7.Text = rs!il�e
Form4.Text8.Text = rs!Do�um_Yeri
Form4.Text9.Text = rs!Do�um_Tarih
Form4.Text10.Text = rs!Cinsiyet
Form4.Text11.Text = rs!Medeni_Hali
Form4.Text12.Text = rs!Mahalle_K�y
Form4.Label18.Caption = rs!�kametgah_Adresi
Form4.Label20.Caption = rs!E_Posta
Form4.Text14.Text = rs!Telefon
Form4.Text15.Text = rs!Mesle�i
Form4.Text13.Text = rs!OdaNo
Form4.Text16.Text = rs!Geli�_Tarihi
Form4.Text17.Text = rs!Geli�_Saati
rs.Update
rs.Close
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
On Error Resume Next
If Form4.Text5.Text = "" And Form4.Text13.Text = "" Then
MsgBox "OTEL�M�ZDE KAYITLI M��TER� YOKTUR."
Form4.Hide
Form3.Show
Else
Form3.Hide
Form4.Show
End If
Set db = OpenDatabase(App.Path & "\�ifre.mdb")
Set rs = db.OpenRecordset("tablo1")
rs.MoveFirst
Form4.Text1.Text = rs!Ad�
Form4.Text2.Text = rs!soyad�
Form4.Text3.Text = rs!Baba_ad�
Form4.Text4.Text = rs!Anne_ad�
Form4.Text5.Text = rs!Tc
Form4.Text6.Text = rs!il
Form4.Text7.Text = rs!il�e
Form4.Text8.Text = rs!Do�um_Yeri
Form4.Text9.Text = rs!Do�um_Tarih
Form4.Text10.Text = rs!Cinsiyet
Form4.Text11.Text = rs!Medeni_Hali
Form4.Text12.Text = rs!Mahalle_K�y
Form4.Label18.Caption = rs!�kametgah_Adresi
Form4.Label20.Caption = rs!E_Posta
Form4.Text14.Text = rs!Telefon
Form4.Text15.Text = rs!Mesle�i
Form4.Text13.Text = rs!OdaNo
Form4.Text16.Text = rs!Geli�_Tarihi
Form4.Text17.Text = rs!Geli�_Saati
rs.Update
rs.Close
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

Private Sub Command8_Click()
On Error Resume Next
Set db = OpenDatabase(App.Path & "\�ifre.mdb")
Set rs = db.OpenRecordset("tablo2")
rs.MoveFirst
Form10.Label2.Caption = rs!OdaNo
Form10.Label4.Caption = rs!Ad�
Form10.Label6.Caption = rs!soyad�
Form10.Label8.Caption = rs!Tc
Form10.Label10.Caption = rs!�kametgah_Adresi
Form10.Label12.Caption = rs!Telefon
Form10.Label14.Caption = rs!Geli�_Tarihi
Form10.Label16.Caption = rs!Geli�_Saati
Form10.Label18.Caption = rs!Ayr�l��_Tarihi
Form10.Label20.Caption = rs!Ayr�l��_Saati
Form10.Label22.Caption = rs!G�n
Form10.Label24.Caption = rs!Fiyat
rs.Update
rs.Close
If Form10.Label2.Caption = "" And Form10.Label8.Caption = "" Then
MsgBox "OTEL�M�ZDEN M��TER� AYRILMAMI�TIR."
Form10.Hide
Form3.Show
Else
Form3.Hide
Form10.Show
End If
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
