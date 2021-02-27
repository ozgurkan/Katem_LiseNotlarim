VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   8190
   ClientLeft      =   3330
   ClientTop       =   765
   ClientWidth     =   8925
   LinkTopic       =   "Form6"
   ScaleHeight     =   8190
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   8175
      Left            =   0
      Picture         =   "Form6.frx":0000
      ScaleHeight     =   8115
      ScaleWidth      =   8835
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3360
         Top             =   120
      End
      Begin VB.TextBox Text1 
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
         Left            =   3960
         TabIndex        =   4
         ToolTipText     =   "T.C KÝMLÝK NUMARASI GÝRÝNÝZ"
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "BUL"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "MÜÞTERÝ ARAMAK ÝÇÝN TIKLAYIN"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "ANA SAYFAYA DÖN"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "ANA SAYFAYA DÖNMEK ÝÇÝN TIKLAYIN"
         Top             =   7320
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "MÜÞTERÝ AYRILIÞ"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "MÜÞTERÝ AYRILIÞI ÝÇÝN TIKLAYIN"
         Top             =   7320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TL"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   26
         Top             =   6840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "FÝYAT===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   6840
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   24
         Top             =   6840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   23
         Top             =   6360
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "KONAKLAMA GÜNÜ===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   6360
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Left            =   3960
         TabIndex        =   21
         Top             =   120
         Width           =   1215
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
         Height          =   375
         Left            =   5400
         TabIndex        =   20
         Top             =   120
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   5400
         X2              =   5160
         Y1              =   120
         Y2              =   480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "AD===>"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "SOYAD===>"
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
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "T.C KÝMLÝK NUMARASI===>"
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
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "ODA NUMARASI===>"
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
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4080
         TabIndex        =   15
         Top             =   3000
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "ÝKAMETGAH ADRESÝ===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   4080
         TabIndex        =   13
         Top             =   3600
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "TELEFON NUMARASI===>"
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
         Left            =   120
         TabIndex        =   12
         Top             =   4560
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4080
         TabIndex        =   11
         Top             =   4560
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "GELÝÞ TARÝHÝ===>"
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
         Left            =   120
         TabIndex        =   10
         Top             =   5160
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4080
         TabIndex        =   9
         Top             =   5160
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "GELÝÞ SAATÝ===>"
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
         Left            =   120
         TabIndex        =   8
         Top             =   5760
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4080
         TabIndex        =   7
         Top             =   5760
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4080
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4080
         TabIndex        =   5
         Top             =   2400
         Visible         =   0   'False
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Len(Text1) < 11 Then
MsgBox "T.C KÝMLÝK NUMARASI 11 HANELÝ OLMALIDIR."
Text1.Text = ""
End If
If Text1 = "" Then
MsgBox ("LÜTFEN T.C KÝMLÝK NUMARASI GÝRÝNÝZ")
Else
Label3.Visible = True
Label4.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label13.Visible = True
Label14.Visible = True
Label15.Visible = True
Label16.Visible = True
Label17.Visible = True
Command3.Visible = True
Label18.Visible = True
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True
Label22.Visible = True
Set db = OpenDatabase("c:\hotel\hotel programý giriþ sayfasý\þifre.mdb")
Set rs = db.OpenRecordset("tablo1", dbOpenSnapshot)
a = rs.RecordCount
For sayac = 0 To a
If Text1 = rs!Tc Then

Label7.Caption = rs!OdaNo
Label9.Caption = rs!ikametgah_Adresi
Label11.Caption = rs!Telefon
Label13.Caption = rs!Geliþ_Tarihi
Label15.Caption = rs!Geliþ_Saati
Label16.Caption = rs!Adý
Label17.Caption = rs!soyadý
Label19.Caption = rs!Gün
Label21.Caption = rs!Fiyat
Exit For
Else
rs.MoveNext
End If
Next sayac
If Label16.Caption = "" And Label16.Visible = True Then
MsgBox "MÜÞTERÝ BULUNAMADI."
Text1.Text = ""
Command3.Visible = False
Label16.Visible = False
Label17.Visible = False
Label7.Visible = False
Label9.Visible = False
Label11.Visible = False
Label13.Visible = False
Label15.Visible = False
Label3.Visible = False
Label4.Visible = False
Label6.Visible = False
Label8.Visible = False
Label10.Visible = False
Label12.Visible = False
Label14.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
End If
End If
End Sub

Private Sub Command2_Click()
Text1 = ""
Label16.Caption = ""
Label17.Caption = ""
Label7.Caption = ""
Label9.Caption = ""
Label11.Caption = ""
Label13.Caption = ""
Label15.Caption = ""
Label19.Caption = ""
Label21.Caption = ""
Command3.Visible = False
Label16.Visible = False
Label17.Visible = False
Label7.Visible = False
Label9.Visible = False
Label11.Visible = False
Label13.Visible = False
Label15.Visible = False
Label3.Visible = False
Label4.Visible = False
Label6.Visible = False
Label8.Visible = False
Label10.Visible = False
Label12.Visible = False
Label14.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False

Form6.Hide
Form3.Show
End Sub

Private Sub Command3_Click()
On Error Resume Next
cevap = MsgBox("müþteri ayrýlmasýný onaylýyormusunuz?", 36, "onay iþlemi")
If cevap = 6 Then
If Text1 = "" Then
MsgBox "MÜÞTERÝNÝN T.C KÝMLÝK NUMARASINI GÝRÝNÝZ"
End If
If Label16.Caption <> "" And Label16.Visible = True Then
MsgBox "MÜÞTERÝ AYRILIÞ ÝÞLEMLERÝ YAPILIYOR."
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo2")
rs.AddNew
rs.Fields("Tc") = Text1.Text
rs.Fields("Adý") = Label16.Caption
rs.Fields("Soyadý") = Label17.Caption
rs.Fields("OdaNo") = Label7.Caption
rs.Fields("Ýkametgah_Adresi") = Label9.Caption
rs.Fields("Telefon") = Label11.Caption
rs.Fields("Geliþ_Tarihi") = Label13.Caption
rs.Fields("Geliþ_Saati") = Label15.Caption
rs.Fields("Ayrýlýþ_Tarihi") = Label1.Caption
rs.Fields("Ayrýlýþ_Saati") = Label2.Caption
rs.Fields("Gün") = Label19.Caption
rs.Fields("Fiyat") = Label21.Caption
rs.Update
rs.Close


Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If Text1 = rs!Tc Then
rs.Delete
Exit For
Else
rs.MoveNext
End If
Next sayac
rs.Update
rs.Close





Label16.Caption = ""
Label17.Caption = ""
Label7.Caption = ""
Label9.Caption = ""
Label11.Caption = ""
Label13.Caption = ""
Label15.Caption = ""
Text1 = ""
Label19.Caption = ""
Label21.Caption = ""
Command3.Visible = False
Label16.Visible = False
Label17.Visible = False
Label7.Visible = False
Label9.Visible = False
Label11.Visible = False
Label13.Visible = False
Label15.Visible = False
Label3.Visible = False
Label4.Visible = False
Label6.Visible = False
Label8.Visible = False
Label10.Visible = False
Label12.Visible = False
Label14.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
MsgBox "MÜÞTERÝ OTELÝMÝZDEN AYRILMIÞTIR."
End If
ElseIf cevap = 7 Then
MsgBox "ayrýlýþ iþlemi onaylanmadý."
End If
End Sub

Private Sub Form_Load()
Label1.Caption = Date
Label2.Caption = Time
End Sub

Private Sub Text1_Change()
If Len(Text1) > 11 Then
MsgBox "T.C KÝMLÝK NUMARASI 11 HANELÝ OLMALIDIR."
Text1.Text = ""
End If
Label16.Caption = ""
Label17.Caption = ""
Label7.Caption = ""
Label9.Caption = ""
Label11.Caption = ""
Label13.Caption = ""
Label15.Caption = ""
Label19.Caption = ""
Label21.Caption = ""
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Date
Label2.Caption = Time
End Sub
