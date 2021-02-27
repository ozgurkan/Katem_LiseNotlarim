VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   7095
   ClientLeft      =   1200
   ClientTop       =   1185
   ClientWidth     =   11745
   LinkTopic       =   "Form10"
   ScaleHeight     =   7095
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   7095
      Left            =   0
      Picture         =   "Form10.frx":0000
      ScaleHeight     =   7035
      ScaleWidth      =   11715
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "ANA SAYFAYA DÖN"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5160
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "ilk kayýt"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ýlk kayda gitmek için týklayýn."
         Top             =   5880
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "ileri"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Sonraki kayda gitmek için týklayýn."
         Top             =   5880
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080C0FF&
         Caption         =   "geri"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Önceki kayda gitmek için týklayýn."
         Top             =   5880
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080C0FF&
         Caption         =   "son kayýt"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Son kayda gitmek için týklayýn."
         Top             =   5880
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "ODA NUMARASI===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3120
         TabIndex        =   31
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5760
         TabIndex        =   30
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "ADI===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   29
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   28
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "SOYADI===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   27
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   26
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   "T.C KÝMLÝK NUMARASI===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   25
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         TabIndex        =   24
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         Caption         =   "ÝKAMETGAH ADRESÝ===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   23
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         TabIndex        =   22
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         Caption         =   "TELEFON NUMARASI===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   21
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         TabIndex        =   20
         Top             =   4200
         Width           =   2775
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
         Caption         =   "GELÝÞ TARÝHÝ===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   19
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         TabIndex        =   18
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080C0FF&
         Caption         =   "GELÝÞ SAATÝ===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   17
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         TabIndex        =   16
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080C0FF&
         Caption         =   "AYRILIÞ TARÝHÝ===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   15
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         TabIndex        =   14
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackColor       =   &H0080C0FF&
         Caption         =   "AYRILIÞ SAATÝ===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   13
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         TabIndex        =   12
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label21 
         BackColor       =   &H0080C0FF&
         Caption         =   "KONAKLAMA GÜNÜ===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6000
         TabIndex        =   11
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9480
         TabIndex        =   10
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label23 
         BackColor       =   &H0080C0FF&
         Caption         =   "FÝYAT===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6000
         TabIndex        =   9
         Top             =   4440
         Width           =   3255
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9480
         TabIndex        =   8
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "KAYIT NUMARASI===>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   7
         Top             =   6480
         Width           =   1935
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FF8080&
         Caption         =   "0 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   6
         Top             =   6480
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Form10.Hide
Form3.Show
Label26.Caption = 0
On Error Resume Next
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo2")
rs.MoveFirst
Label2.Caption = rs!OdaNo
Label4.Caption = rs!Adý
Label6.Caption = rs!soyadý
Label8.Caption = rs!Tc
Label10.Caption = rs!Ýkametgah_Adresi
Label12.Caption = rs!Telefon
Label14.Caption = rs!Geliþ_Tarihi
Label16.Caption = rs!Geliþ_Saati
Label18.Caption = rs!Ayrýlýþ_Tarihi
Label20.Caption = rs!Ayrýlýþ_Saati
Label22.Caption = rs!Gün
Label24.Caption = rs!Fiyat
rs.Update
rs.Close
End Sub

Private Sub Command2_Click()
On Error Resume Next
Label26.Caption = "0"
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo2")
rs.MoveFirst
Label2.Caption = rs!OdaNo
Label4.Caption = rs!Adý
Label6.Caption = rs!soyadý
Label8.Caption = rs!Tc
Label10.Caption = rs!Ýkametgah_Adresi
Label12.Caption = rs!Telefon
Label14.Caption = rs!Geliþ_Tarihi
Label16.Caption = rs!Geliþ_Saati
Label18.Caption = rs!Ayrýlýþ_Tarihi
Label20.Caption = rs!Ayrýlýþ_Saati
Label22.Caption = rs!Gün
Label24.Caption = rs!Fiyat
MsgBox "ilk kayýttasýnýz"
rs.Update
rs.Close
End Sub

Private Sub Command3_Click()
On Error Resume Next
Label26.Caption = Label26.Caption + 1
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo2")
For sayac = 1 To Label26.Caption
rs.MoveNext
Label2.Caption = rs!OdaNo
Label4.Caption = rs!Adý
Label6.Caption = rs!soyadý
Label8.Caption = rs!Tc
Label10.Caption = rs!Ýkametgah_Adresi
Label12.Caption = rs!Telefon
Label14.Caption = rs!Geliþ_Tarihi
Label16.Caption = rs!Geliþ_Saati
Label18.Caption = rs!Ayrýlýþ_Tarihi
Label20.Caption = rs!Ayrýlýþ_Saati
Label22.Caption = rs!Gün
Label24.Caption = rs!Fiyat
Next sayac
If rs.EOF Then
MsgBox "zaten son kayýttasýnýz"
Label26.Caption = Label26.Caption - 1
End If
rs.Update
rs.Close
End Sub

Private Sub Command4_Click()
On Error Resume Next
Label26.Caption = Label26.Caption - 1
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo2")
For sayac = 1 To Label26.Caption
rs.MoveNext
Next sayac
Label2.Caption = rs!OdaNo
Label4.Caption = rs!Adý
Label6.Caption = rs!soyadý
Label8.Caption = rs!Tc
Label10.Caption = rs!Ýkametgah_Adresi
Label12.Caption = rs!Telefon
Label14.Caption = rs!Geliþ_Tarihi
Label16.Caption = rs!Geliþ_Saati
Label18.Caption = rs!Ayrýlýþ_Tarihi
Label20.Caption = rs!Ayrýlýþ_Saati
Label22.Caption = rs!Gün
Label24.Caption = rs!Fiyat
If Label26.Caption < 0 Then
MsgBox "zaten ilk kayýtttasýnýz."
Label26.Caption = "0"
End If
rs.Update
rs.Close
End Sub

Private Sub Command5_Click()
On Error Resume Next
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo2")
rs.MoveLast
a = rs.RecordCount
Label2.Caption = rs!OdaNo
Label4.Caption = rs!Adý
Label6.Caption = rs!soyadý
Label8.Caption = rs!Tc
Label10.Caption = rs!Ýkametgah_Adresi
Label12.Caption = rs!Telefon
Label14.Caption = rs!Geliþ_Tarihi
Label16.Caption = rs!Geliþ_Saati
Label18.Caption = rs!Ayrýlýþ_Tarihi
Label20.Caption = rs!Ayrýlýþ_Saati
Label22.Caption = rs!Gün
Label24.Caption = rs!Fiyat
MsgBox "son kayýttasýnýz"
Label26.Caption = a - 1
rs.Update
rs.Close
End Sub

Private Sub Form_Load()
On Error Resume Next
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo2")
rs.MoveFirst
Label2.Caption = rs!OdaNo
Label4.Caption = rs!Adý
Label6.Caption = rs!soyadý
Label8.Caption = rs!Tc
Label10.Caption = rs!Ýkametgah_Adresi
Label12.Caption = rs!Telefon
Label14.Caption = rs!Geliþ_Tarihi
Label16.Caption = rs!Geliþ_Saati
Label18.Caption = rs!Ayrýlýþ_Tarihi
Label20.Caption = rs!Ayrýlýþ_Saati
Label22.Caption = rs!Gün
Label24.Caption = rs!Fiyat
rs.Update
rs.Close
End Sub

Private Sub Label16_Click()

End Sub

Private Sub Picture1_Click()

End Sub
