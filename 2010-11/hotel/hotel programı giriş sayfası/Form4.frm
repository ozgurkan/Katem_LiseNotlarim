VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   9240
   ClientLeft      =   1275
   ClientTop       =   645
   ClientWidth     =   12150
   LinkTopic       =   "Form4"
   ScaleHeight     =   9240
   ScaleWidth      =   12150
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   9255
      Left            =   0
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   9195
      ScaleWidth      =   12075
      TabIndex        =   0
      Top             =   0
      Width           =   12135
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
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text2 
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
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
         Height          =   735
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Sonraki kayda gitmek için týklayýn."
         Top             =   7920
         Width           =   855
      End
      Begin VB.TextBox Text3 
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2280
         Width           =   2895
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
         Height          =   735
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ýlk kayda gitmek için týklayýn."
         Top             =   7920
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Height          =   735
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Son kayda gitmek için týklayýn."
         Top             =   7920
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
         Height          =   735
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Önceki kayda gitmek için týklayýn."
         Top             =   7920
         Width           =   855
      End
      Begin VB.TextBox Text4 
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text6 
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4440
         Width           =   2895
      End
      Begin VB.TextBox Text7 
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4920
         Width           =   2895
      End
      Begin VB.TextBox Text8 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text9 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text10 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox Text11 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox Text12 
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   5400
         Width           =   2895
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   4560
         Width           =   2895
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   120
         Width           =   2895
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ana Sayfaya Dön"
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ana sayfaya dönmek için týklayýn."
         Top             =   7320
         Width           =   1695
      End
      Begin VB.TextBox Text16 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   6240
         Width           =   2895
      End
      Begin VB.TextBox Text17 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   6840
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "ADI===>"
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
         Left            =   240
         TabIndex        =   45
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Soyadý===>"
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
         Left            =   240
         TabIndex        =   44
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Baba Adý===>"
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
         Left            =   240
         TabIndex        =   43
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label4 
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
         Left            =   6000
         TabIndex        =   42
         Top             =   8640
         Width           =   615
      End
      Begin VB.Label Label5 
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
         Left            =   4080
         TabIndex        =   41
         Top             =   8640
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Anne Adý==>"
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
         Left            =   240
         TabIndex        =   40
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   "T.C Kimlik Numarasý===>"
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
         Left            =   240
         TabIndex        =   39
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ýl===>"
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
         Left            =   240
         TabIndex        =   38
         Top             =   4440
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ýlçe===>"
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
         Left            =   240
         TabIndex        =   37
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         Caption         =   "Doðum Yeri===>"
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
         Left            =   6480
         TabIndex        =   36
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         Caption         =   "Doðum Tarihi===>"
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
         Left            =   6480
         TabIndex        =   35
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cinsiyeti===>"
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
         Left            =   6480
         TabIndex        =   34
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
         Caption         =   "Medeni Hali===>"
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
         Left            =   6480
         TabIndex        =   33
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ýkametgah Adresi===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   32
         Top             =   5880
         Width           =   2415
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080C0FF&
         Caption         =   "E-Posta Adresi===>"
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
         Left            =   6480
         TabIndex        =   31
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080C0FF&
         Caption         =   "Telefon Numarasý===>"
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
         Left            =   6480
         TabIndex        =   30
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080C0FF&
         Caption         =   "Mesleði===>"
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
         Left            =   6480
         TabIndex        =   29
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2880
         TabIndex        =   28
         Top             =   5880
         Width           =   2895
      End
      Begin VB.Label Label19 
         BackColor       =   &H0080C0FF&
         Caption         =   "Mahalle/Köy===>"
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
         Left            =   240
         TabIndex        =   27
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   9120
         TabIndex        =   26
         Top             =   5400
         Width           =   2895
      End
      Begin VB.Label Label21 
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
         Height          =   735
         Left            =   3720
         TabIndex        =   25
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label22 
         BackColor       =   &H0080C0FF&
         Caption         =   "Geliþ Tarihi===>"
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
         Left            =   6480
         TabIndex        =   24
         Top             =   6240
         Width           =   2415
      End
      Begin VB.Label Label23 
         BackColor       =   &H0080C0FF&
         Caption         =   "Geliþ Saati===>"
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
         Left            =   6480
         TabIndex        =   23
         Top             =   6840
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Label4.Caption = Label4.Caption + 1
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
For sayac = 1 To Label4.Caption
rs.MoveNext
Text1.Text = rs!Adý
Text2.Text = rs!soyadý
Text3.Text = rs!Baba_adý
Text4.Text = rs!Anne_adý
Text5.Text = rs!Tc
Text6.Text = rs!il
Text7.Text = rs!ilçe
Text8.Text = rs!Doðum_Yeri
Text9.Text = rs!Doðum_Tarih
Text10.Text = rs!Cinsiyet
Text11.Text = rs!Medeni_Hali
Text12.Text = rs!Mahalle_Köy
Label18.Caption = rs!Ýkametgah_Adresi
Label20.Caption = rs!E_Posta
Text14.Text = rs!Telefon
Text15.Text = rs!Mesleði
Text13.Text = rs!OdaNo
Text16.Text = rs!Geliþ_Tarihi
Text17.Text = rs!Geliþ_Saati
Next sayac
If rs.EOF Then
MsgBox "zaten son kayýttasýnýz"
Label4.Caption = Label4.Caption - 1
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Label4.Caption = "0"
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
rs.MoveFirst
Text1.Text = rs!Adý
Text2.Text = rs!soyadý
Text3.Text = rs!Baba_adý
Text4.Text = rs!Anne_adý
Text5.Text = rs!Tc
Text6.Text = rs!il
Text7.Text = rs!ilçe
Text8.Text = rs!Doðum_Yeri
Text9.Text = rs!Doðum_Tarih
Text10.Text = rs!Cinsiyet
Text11.Text = rs!Medeni_Hali
Text12.Text = rs!Mahalle_Köy
Label18.Caption = rs!Ýkametgah_Adresi
Label20.Caption = rs!E_Posta
Text14.Text = rs!Telefon
Text15.Text = rs!Mesleði
Text13.Text = rs!OdaNo
Text16.Text = rs!Geliþ_Tarihi
Text17.Text = rs!Geliþ_Saati
MsgBox "ilk kayýttasýnýz"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
rs.MoveLast
a = rs.RecordCount
Text1.Text = rs!Adý
Text2.Text = rs!soyadý
Text3.Text = rs!Baba_adý
Text4.Text = rs!Anne_adý
Text5.Text = rs!Tc
Text6.Text = rs!il
Text7.Text = rs!ilçe
Text8.Text = rs!Doðum_Yeri
Text9.Text = rs!Doðum_Tarih
Text10.Text = rs!Cinsiyet
Text11.Text = rs!Medeni_Hali
Text12.Text = rs!Mahalle_Köy
Label18.Caption = rs!Ýkametgah_Adresi
Label20.Caption = rs!E_Posta
Text14.Text = rs!Telefon
Text15.Text = rs!Mesleði
Text13.Text = rs!OdaNo
Text16.Text = rs!Geliþ_Tarihi
Text17.Text = rs!Geliþ_Saati
MsgBox "son kayýttasýnýz"
Label4.Caption = a - 1

End Sub

Private Sub Command4_Click()
On Error Resume Next
Label4.Caption = Label4.Caption - 1
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
For sayac = 1 To Label4.Caption
rs.MoveNext
Next sayac
Text1.Text = rs!Adý
Text2.Text = rs!soyadý
Text3.Text = rs!Baba_adý
Text4.Text = rs!Anne_adý
Text5.Text = rs!Tc
Text6.Text = rs!il
Text7.Text = rs!ilçe
Text8.Text = rs!Doðum_Yeri
Text9.Text = rs!Doðum_Tarih
Text10.Text = rs!Cinsiyet
Text11.Text = rs!Medeni_Hali
Text12.Text = rs!Mahalle_Köy
Label18.Caption = rs!Ýkametgah_Adresi
Label20.Caption = rs!E_Posta
Text14.Text = rs!Telefon
Text15.Text = rs!Mesleði
Text13.Text = rs!OdaNo
Text16.Text = rs!Geliþ_Tarihi
Text17.Text = rs!Geliþ_Saati
If Label4.Caption < 0 Then
MsgBox "zaten ilk kayýtttasýnýz."
Label4.Caption = "0"
End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
Form4.Hide
Form3.Show
Label4.Caption = 0
On Error Resume Next
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
rs.MoveFirst
Text1.Text = rs!Adý
Text2.Text = rs!soyadý
Text3.Text = rs!Baba_adý
Text4.Text = rs!Anne_adý
Text5.Text = rs!Tc
Text6.Text = rs!il
Text7.Text = rs!ilçe
Text8.Text = rs!Doðum_Yeri
Text9.Text = rs!Doðum_Tarih
Text10.Text = rs!Cinsiyet
Text11.Text = rs!Medeni_Hali
Text12.Text = rs!Mahalle_Köy
Label18.Caption = rs!Ýkametgah_Adresi
Label20.Caption = rs!E_Posta
Text14.Text = rs!Telefon
Text15.Text = rs!Mesleði
Text13.Text = rs!OdaNo
Text16.Text = rs!Geliþ_Tarihi
Text17.Text = rs!Geliþ_Saati
End Sub
Private Sub Form_Load()
On Error Resume Next
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
rs.MoveFirst
Text1.Text = rs!Adý
Text2.Text = rs!soyadý
Text3.Text = rs!Baba_adý
Text4.Text = rs!Anne_adý
Text5.Text = rs!Tc
Text6.Text = rs!il
Text7.Text = rs!ilçe
Text8.Text = rs!Doðum_Yeri
Text9.Text = rs!Doðum_Tarih
Text10.Text = rs!Cinsiyet
Text11.Text = rs!Medeni_Hali
Text12.Text = rs!Mahalle_Köy
Label18.Caption = rs!Ýkametgah_Adresi
Label20.Caption = rs!E_Posta
Text14.Text = rs!Telefon
Text15.Text = rs!Mesleði
Text13.Text = rs!OdaNo
Text16.Text = rs!Geliþ_Tarihi
Text17.Text = rs!Geliþ_Saati
End Sub

