VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   9000
   ClientLeft      =   1275
   ClientTop       =   510
   ClientWidth     =   12000
   LinkTopic       =   "Form5"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   9015
      Left            =   0
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   8955
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin MSACAL.Calendar Calendar1 
         Height          =   3255
         Left            =   8160
         TabIndex        =   45
         Top             =   2280
         Visible         =   0   'False
         Width           =   3735
         _Version        =   524288
         _ExtentX        =   6588
         _ExtentY        =   5741
         _StockProps     =   1
         BackColor       =   65280
         Year            =   2011
         Month           =   5
         Day             =   27
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   255
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   0
         GridLinesColor  =   255
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   12582912
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Left            =   9240
         TabIndex        =   42
         Top             =   6120
         Width           =   2655
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "Form5.frx":16156A
         Left            =   9240
         List            =   "Form5.frx":161574
         TabIndex        =   40
         Top             =   3240
         Width           =   2655
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Form5.frx":161585
         Left            =   9240
         List            =   "Form5.frx":16158F
         TabIndex        =   39
         Top             =   2760
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2880
         TabIndex        =   38
         Top             =   5040
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form5.frx":1615A1
         Left            =   2880
         List            =   "Form5.frx":161698
         TabIndex        =   37
         Top             =   4560
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Kaydet"
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Müþteri giriþi yapmak için týklayýn."
         Top             =   7920
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ana Sayfaya Dön"
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
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ana sayfaya dönmek için týklayýn."
         Top             =   7920
         Width           =   1815
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
         TabIndex        =   14
         Top             =   3720
         Width           =   2655
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
         TabIndex        =   13
         Top             =   3240
         Width           =   2655
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
         Height          =   975
         IMEMode         =   3  'DISABLE
         Left            =   2880
         TabIndex        =   12
         Top             =   6000
         Width           =   2655
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
         Height          =   735
         Left            =   9240
         TabIndex        =   11
         Top             =   5280
         Width           =   2655
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
         Left            =   2880
         TabIndex        =   10
         Top             =   5520
         Width           =   2655
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
         Left            =   9240
         TabIndex        =   9
         Top             =   1800
         Width           =   2655
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
         Left            =   9240
         TabIndex        =   8
         Top             =   2280
         Width           =   2655
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
         Height          =   855
         Left            =   9240
         TabIndex        =   7
         Top             =   3720
         Width           =   2655
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
         Height          =   495
         Left            =   9240
         TabIndex        =   6
         Top             =   4680
         Width           =   2655
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
         Height          =   735
         Left            =   6120
         TabIndex        =   5
         Top             =   960
         Width           =   2775
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
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   1800
         Width           =   2655
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
         TabIndex        =   3
         Top             =   2280
         Width           =   2655
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
         TabIndex        =   2
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   4200
         Top             =   240
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Müþteri Bilgilerini Göster"
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Müþteri bilgilerini görmek istiyorsanýz týklayýn."
         Top             =   7200
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label23 
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
         Left            =   9240
         TabIndex        =   44
         Top             =   6600
         Width           =   2655
      End
      Begin VB.Label Label22 
         BackColor       =   &H0080C0FF&
         Caption         =   "Fiyat===>"
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
         Left            =   6600
         TabIndex        =   43
         Top             =   6600
         Width           =   2415
      End
      Begin VB.Label Label21 
         BackColor       =   &H0080C0FF&
         Caption         =   "Konaklama Günü===>"
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
         Left            =   6600
         TabIndex        =   41
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Label Label17 
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
         Left            =   3480
         TabIndex        =   36
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label8 
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
         Index           =   1
         Left            =   240
         TabIndex        =   35
         Top             =   5520
         Width           =   2415
      End
      Begin VB.Label Label15 
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
         Height          =   495
         Index           =   1
         Left            =   6600
         TabIndex        =   34
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label14 
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
         Left            =   6600
         TabIndex        =   33
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label16 
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
         Index           =   0
         Left            =   6600
         TabIndex        =   32
         Top             =   5280
         Width           =   2415
      End
      Begin VB.Label Label9 
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
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Top             =   6000
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
         Left            =   6600
         TabIndex        =   30
         Top             =   3240
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
         Left            =   6600
         TabIndex        =   29
         Top             =   2760
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
         Left            =   6600
         TabIndex        =   28
         Top             =   2280
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
         Left            =   6600
         TabIndex        =   27
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label7 
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
         TabIndex        =   26
         Top             =   5040
         Width           =   2415
      End
      Begin VB.Label Label6 
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
         TabIndex        =   25
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label Label5 
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
         TabIndex        =   24
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label4 
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
         TabIndex        =   23
         Top             =   3240
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
         TabIndex        =   22
         Top             =   2760
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
         TabIndex        =   21
         Top             =   2280
         Width           =   2415
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
         TabIndex        =   20
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "BOÞ ALANLAR SARI RENKTE  GÖSTERÝLMÝÞTÝR."
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
         Left            =   3000
         TabIndex        =   19
         Top             =   8520
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   495
         Left            =   4680
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   495
         Left            =   6480
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   6480
         X2              =   6240
         Y1              =   240
         Y2              =   720
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()
Text11.Text = Calendar1.Value
Calendar1.Visible = False
a = Label19.Caption
b = Text11.Text
s = Right(a, 4)
c = Right(b, 4)
d = s - c
If d < 18 Then
MsgBox "18 yaþýndan küçükler alýnmaz"
Text11 = ""
Calendar1.Visible = False
End If

End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
Combo2.Clear
Combo2.AddItem "ALADAÐ"
Combo2.AddItem "CEYHAN"
Combo2.AddItem "FEKE"
Combo2.AddItem "ÝMAMOÐLU"
Combo2.AddItem "KARAÝSALI"
Combo2.AddItem "KARATAÞ"
Combo2.AddItem "KOZAN"
Combo2.AddItem "POZANTI"
Combo2.AddItem "SAÝMBEYLÝ"
Combo2.AddItem "SEYHAN"
Combo2.AddItem "TUFANBEYLÝ"
Combo2.AddItem "YUMURTALIK"
Combo2.AddItem "YÜREÐÝR"
ElseIf Combo1.ListIndex = 1 Then
Combo2.Clear
Combo2.AddItem "ADIYAMAN MERKEZ"
Combo2.AddItem "BESNÝ"
Combo2.AddItem "ÇELÝKHAN"
Combo2.AddItem "GERGER"
Combo2.AddItem "GÖLBAÞI/ADIYAMAN"
Combo2.AddItem "KAHTA"
Combo2.AddItem "SAMSAT"
Combo2.AddItem "SÝNCÝK"
Combo2.AddItem "TUT"
ElseIf Combo1.ListIndex = 2 Then
Combo2.Clear
Combo2.AddItem "AFYONKARAHÝSAR MERKEZ"
Combo2.AddItem "BAÞMAKÇI"
Combo2.AddItem "BAYAT/AFYON"
Combo2.AddItem "BOLVADÝN"
Combo2.AddItem "ÇAY"
Combo2.AddItem "ÇOBANLAR"
Combo2.AddItem "DAZKIRI"
Combo2.AddItem "DÝNAR"
Combo2.AddItem "EMÝRDAÐ"
Combo2.AddItem "EVCÝLER"
Combo2.AddItem "HOCALAR"
Combo2.AddItem "ÝHSANÝYE"
Combo2.AddItem "ÝSCEHÝSAR"
Combo2.AddItem "KIZILÖREN"
Combo2.AddItem "SANDIKLI"
Combo2.AddItem "SÝNANPAÞA"
Combo2.AddItem "ÞUHUT"
Combo2.AddItem "SULTANDAÐI"
ElseIf Combo1.ListIndex = 3 Then
Combo2.Clear
Combo2.AddItem "AÐRI MERKEZ"
Combo2.AddItem "DÝYADÝN"
Combo2.AddItem "DOÐUBAYAZIT"
Combo2.AddItem "ELEÞKÝRT"
Combo2.AddItem "HAMUR"
Combo2.AddItem "PATNOS"
Combo2.AddItem "TAÞLIÇAY"
Combo2.AddItem "TUTAK"
ElseIf Combo1.ListIndex = 4 Then
Combo2.Clear
Combo2.AddItem "AÐAÇÖREN"
Combo2.AddItem "AKSARAY MERKEZ"
Combo2.AddItem "ESKÝL"
Combo2.AddItem "GÜLAÐAÇ"
Combo2.AddItem "GÜZELYURT"
Combo2.AddItem "ORTAKÖY"
Combo2.AddItem "SARIYAHÞÝ"
ElseIf Combo1.ListIndex = 5 Then
Combo2.Clear
Combo2.AddItem "AMASYA MERKEZ"
Combo2.AddItem "GÖYNÜCEK"
Combo2.AddItem "GÜMÜÞHACIKÖY"
Combo2.AddItem "HAMAMÖZÜ"
Combo2.AddItem "MERZÝFON"
Combo2.AddItem "SULUOVA"
Combo2.AddItem "TAÞOVA"
ElseIf Combo1.ListIndex = 6 Then
Combo2.Clear
Combo2.AddItem "AKYURT"
Combo2.AddItem "ALTINDAÐ"
Combo2.AddItem "ANKARA MERKEZ"
Combo2.AddItem "AYAÞ"
Combo2.AddItem "BALA"
Combo2.AddItem "BEYPAZARI"
Combo2.AddItem "ÇAMLIDERE"
Combo2.AddItem "ÇANKAYA"
Combo2.AddItem "ÇUBUK"
Combo2.AddItem "ELMADAÐ"
Combo2.AddItem "ETÝMESGUT"
Combo2.AddItem "EVREN"
Combo2.AddItem "GÖLBAÞI/ANKARA"
Combo2.AddItem "GÜDÜL"
Combo2.AddItem "HAYMANA"
Combo2.AddItem "KALECÝK"
Combo2.AddItem "KAZAN"
Combo2.AddItem "KEÇÝÖREN"
Combo2.AddItem "KIZILCAHAMAM"
Combo2.AddItem "MAMAK"
Combo2.AddItem "NALLIHAN"
Combo2.AddItem "POLATLI"
Combo2.AddItem "ÞEREFLÝKOÇHÝSAR"
Combo2.AddItem "SÝNCAN"
Combo2.AddItem "YENÝMAHALLE"
ElseIf Combo1.ListIndex = 7 Then
Combo2.Clear
Combo2.AddItem "AKSEKÝ"
Combo2.AddItem "ALANYA"
Combo2.AddItem "ANTALYA MERKEZ"
Combo2.AddItem "DEMRE"
Combo2.AddItem "ELMALI"
Combo2.AddItem "FÝNÝKE"
Combo2.AddItem "GAZÝPAÞA"
Combo2.AddItem "GÜNDOÐMUÞ"
Combo2.AddItem "ÝBRADI"
Combo2.AddItem "KAÞ"
Combo2.AddItem "KEMER/ANTALYA"
Combo2.AddItem "KORKUTELÝ"
Combo2.AddItem "KUMLUCA"
Combo2.AddItem "MANAVGAT"
Combo2.AddItem "SERÝK"
ElseIf Combo1.ListIndex = 8 Then
Combo2.Clear
Combo2.AddItem "ARDAHAN MERKEZ"
Combo2.AddItem "ÇILDIR"
Combo2.AddItem "DAMAL"
Combo2.AddItem "GÖLE"
Combo2.AddItem "HANAK"
Combo2.AddItem "POSOF"
ElseIf Combo1.ListIndex = 9 Then
Combo2.Clear
Combo2.AddItem "ARDANUÇ"
Combo2.AddItem "ARHAVÝ"
Combo2.AddItem "ARTVÝN MERKEZ"
Combo2.AddItem "BORÇKA"
Combo2.AddItem "HOPA"
Combo2.AddItem "MURGUL"
Combo2.AddItem "ÞAVÞAT"
Combo2.AddItem "YUSUFELÝ"
ElseIf Combo1.ListIndex = 10 Then
Combo2.Clear
Combo2.AddItem "AYDIN MERKEZ"
Combo2.AddItem "BOZDOÐAN"
Combo2.AddItem "BUHARKENT"
Combo2.AddItem "ÇÝNE"
Combo2.AddItem "DÝDÝM(YENÝHÝSAR)"
Combo2.AddItem "GERMENCÝK"
Combo2.AddItem "ÝNCÝRLÝOVA"
Combo2.AddItem "KARACASU"
Combo2.AddItem "KARPUZLU"
Combo2.AddItem "KOÇARLI"
Combo2.AddItem "KÖÞK"
Combo2.AddItem "KUÞADASI"
Combo2.AddItem "KUYUCAK"
Combo2.AddItem "NAZÝLLÝ"
Combo2.AddItem "SÖKE"
Combo2.AddItem "SULTANHÝSAR"
Combo2.AddItem "YENÝPAZAR/AYDIN"
ElseIf Combo1.ListIndex = 11 Then
Combo2.Clear
Combo2.AddItem "AYVALIK"
Combo2.AddItem "BALIKESÝR MERKEZ"
Combo2.AddItem "BALYA"
Combo2.AddItem "BANDIRMA"
Combo2.AddItem "BÝGADÝÇ"
Combo2.AddItem "BURHANÝYE"
Combo2.AddItem "DURSUNBEY"
Combo2.AddItem "EDREMÝT/BALIKESÝR"
Combo2.AddItem "ERDEK"
Combo2.AddItem "GÖMEÇ"
Combo2.AddItem "GÖNEN/BALIKESÝR"
Combo2.AddItem "HAVRAN"
Combo2.AddItem "ÝVRÝNDÝ"
Combo2.AddItem "KEPSUT"
Combo2.AddItem "MANYAS"
Combo2.AddItem "MARMARA"
Combo2.AddItem "SAVAÞTEPE"
Combo2.AddItem "SINDIRGI"
Combo2.AddItem "SUSURLUK"
ElseIf Combo1.ListIndex = 12 Then
Combo2.Clear
Combo2.AddItem "AMASRA"
Combo2.AddItem "BARTIN MERKEZ"
Combo2.AddItem "KURUCAÞÝLE"
Combo2.AddItem "ULUS"
ElseIf Combo1.ListIndex = 13 Then
Combo2.Clear
Combo2.AddItem "BATMAN MERKEZ"
Combo2.AddItem "BEÞÝRÝ"
Combo2.AddItem "GERCÜÞ"
Combo2.AddItem "HASANKEYF"
Combo2.AddItem "KOZLUK"
Combo2.AddItem "SASON"
ElseIf Combo1.ListIndex = 14 Then
Combo2.Clear
Combo2.AddItem "AYDINTEPE"
Combo2.AddItem "BAYBURT MERKEZ"
Combo2.AddItem "DEMÝRÖZÜ"
ElseIf Combo1.ListIndex = 15 Then
Combo2.Clear
Combo2.AddItem "BÝLECÝK MERKEZ"
Combo2.AddItem "BOZÜYÜK"
Combo2.AddItem "GÖLPAZARI"
Combo2.AddItem "ÝNHÝSAR"
Combo2.AddItem "OSMANELÝ"
Combo2.AddItem "PAZARYERÝ"
Combo2.AddItem "SÖÐÜT"
Combo2.AddItem "YENÝPAZAR/BÝLECÝK"
ElseIf Combo1.ListIndex = 16 Then
Combo2.Clear
Combo2.AddItem "ADAKLI"
Combo2.AddItem "BÝNGÖL MERKEZ"
Combo2.AddItem "GENÇ"
Combo2.AddItem "KARLIOVA"
Combo2.AddItem "KÝÐI"
Combo2.AddItem "SOLHAN"
Combo2.AddItem "YAYLADERE"
Combo2.AddItem "YEDÝSU"
ElseIf Combo1.ListIndex = 17 Then
Combo2.Clear
Combo2.AddItem "ADÝLCEVAZ"
Combo2.AddItem "AHLAT"
Combo2.AddItem "BÝTLÝS MERKEZ"
Combo2.AddItem "GÜROYMAK"
Combo2.AddItem "HÝZAN"
Combo2.AddItem "MUTKÝ"
Combo2.AddItem "TATVAN"
ElseIf Combo1.ListIndex = 18 Then
Combo2.Clear
Combo2.AddItem "BOLU MERKEZ"
Combo2.AddItem "DÖRTDÝVAN"
Combo2.AddItem "GEREDE"
Combo2.AddItem "GÖYNÜK"
Combo2.AddItem "KIBRISCIK"
Combo2.AddItem "MENGEN"
Combo2.AddItem "MUDURNU"
Combo2.AddItem "SEBEN"
Combo2.AddItem "YENÝÇAÐA"
ElseIf Combo1.ListIndex = 19 Then
Combo2.Clear
Combo2.AddItem "AÐLASUN"
Combo2.AddItem "ALTINYAYLA/BURDUR"
Combo2.AddItem "BUCAK"
Combo2.AddItem "BURDUR MERKEZ"
Combo2.AddItem "ÇAVDIR"
Combo2.AddItem "ÇELTÝKÇÝ"
Combo2.AddItem "GÖLHÝSAR"
Combo2.AddItem "KARAMANLI"
Combo2.AddItem "KEMER/BURDUR"
Combo2.AddItem "TEFENNÝ"
Combo2.AddItem "YEÞÝLOVA"
ElseIf Combo1.ListIndex = 20 Then
Combo2.Clear
Combo2.AddItem "BURSA MERKEZ"
Combo2.AddItem "BÜYÜKORHAN"
Combo2.AddItem "GEMLÝK"
Combo2.AddItem "GÜRSU"
Combo2.AddItem "HARMANCIK"
Combo2.AddItem "ÝNEGÖL"
Combo2.AddItem "ÝZNÝK"
Combo2.AddItem "KARACABEY"
Combo2.AddItem "KELES"
Combo2.AddItem "KESTEL"
Combo2.AddItem "MUDANYA"
Combo2.AddItem "MUSTAFA KEMAL PAÞA"
Combo2.AddItem "NÝLÜFER"
Combo2.AddItem "ORHANELÝ"
Combo2.AddItem "ORHANGAZÝ"
Combo2.AddItem "OSMANGAZÝ"
Combo2.AddItem "YENÝÞEHÝR"
Combo2.AddItem "YILDIRIM"
ElseIf Combo1.ListIndex = 21 Then
Combo2.Clear
Combo2.AddItem "AYVACIK/ÇANAKKALE"
Combo2.AddItem "BAYRAMÝÇ"
Combo2.AddItem "BÝGA"
Combo2.AddItem "BOZCAADA"
Combo2.AddItem "ÇAN"
Combo2.AddItem "ÇANAKKALE MERKEZ"
Combo2.AddItem "ECEABAT"
Combo2.AddItem "EZÝNE"
Combo2.AddItem "GELÝBOLU"
Combo2.AddItem "GÖKÇEADA"
Combo2.AddItem "LAPSEKÝ"
Combo2.AddItem "YENÝCE/ÇANAKKALE"
ElseIf Combo1.ListIndex = 22 Then
Combo2.Clear
Combo2.AddItem "ATKARACALAR"
Combo2.AddItem "BAYRAMÖREN"
Combo2.AddItem "ÇANKIRI MERKEZ"
Combo2.AddItem "ÇERKES"
Combo2.AddItem "ELDÝVAN"
Combo2.AddItem "ILGAZ"
Combo2.AddItem "KIZILIRMAK"
Combo2.AddItem "KORGUN"
Combo2.AddItem "KURÞUNLU"
Combo2.AddItem "ORTA"
Combo2.AddItem "ÞABANÖZÜ"
Combo2.AddItem "YAPRAKLI"
ElseIf Combo1.ListIndex = 23 Then
Combo2.Clear
Combo2.AddItem "ALACA"
Combo2.AddItem "BAYAT/ÇORUM"
Combo2.AddItem "BOÐAZKALE"
Combo2.AddItem "ÇORUM MERKEZ"
Combo2.AddItem "DODURGA"
Combo2.AddItem "ÝSKÝLÝP"
Combo2.AddItem "KARGI"
Combo2.AddItem "LAÇÝN"
Combo2.AddItem "MECÝTÖZÜ"
Combo2.AddItem "OÐUZLAR"
Combo2.AddItem "ORTAKÖY/ÇORUM"
Combo2.AddItem "OSMANCIK"
Combo2.AddItem "SUNGURLU"
Combo2.AddItem "UÐURLUDAÐ"
ElseIf Combo1.ListIndex = 24 Then
Combo2.Clear
Combo2.AddItem "ACIPAYAM"
Combo2.AddItem "AKKÖY"
Combo2.AddItem "BABADAÐ"
Combo2.AddItem "BAKLAN"
Combo2.AddItem "BEKÝLLÝ"
Combo2.AddItem "BEYAÐAÇ"
Combo2.AddItem "BOZKURT/DENÝZLÝ"
Combo2.AddItem "BULDAN"
Combo2.AddItem "ÇAL"
Combo2.AddItem "ÇAMELÝ"
Combo2.AddItem "ÇARDAK"
Combo2.AddItem "ÇÝVRÝL"
Combo2.AddItem "DENÝZLÝ MERKEZ"
Combo2.AddItem "GÜNEY"
Combo2.AddItem "HONAZ"
Combo2.AddItem "KALE/DENÝZLÝ"
Combo2.AddItem "SARAYKÖY"
Combo2.AddItem "SERÝNHÝSAR"
Combo2.AddItem "TAVAS"
ElseIf Combo1.ListIndex = 25 Then
Combo2.Clear
Combo2.AddItem "BÝSMÝL"
Combo2.AddItem "ÇERMÝK"
Combo2.AddItem "ÇINAR"
Combo2.AddItem "ÇÜNGÜÞ"
Combo2.AddItem "DÝCLE"
Combo2.AddItem "DÝYARBAKIR MERKEZ"
Combo2.AddItem "EÐÝL"
Combo2.AddItem "ERGANÝ"
Combo2.AddItem "HANÝ"
Combo2.AddItem "HAZRO"
Combo2.AddItem "KOCAKÖY"
Combo2.AddItem "KULP"
Combo2.AddItem "LÝCE"
Combo2.AddItem "SÝLVAN"
ElseIf Combo1.ListIndex = 26 Then
Combo2.Clear
Combo2.AddItem "AKÇAKOCA"
Combo2.AddItem "ÇÝLÝMLÝ"
Combo2.AddItem "CUMAYERÝ"
Combo2.AddItem "DÜZCE MERKEZ"
Combo2.AddItem "GÖLYAKA"
Combo2.AddItem "GÜMÜÞOVA"
Combo2.AddItem "KAYNAÞLI"
Combo2.AddItem "YIÐILCA"
ElseIf Combo1.ListIndex = 27 Then
Combo2.Clear
Combo2.AddItem "EDÝRNE MERKEZ"
Combo2.AddItem "ENEZ"
Combo2.AddItem "HAVSA"
Combo2.AddItem "ÝPSALA"
Combo2.AddItem "KEÞAN"
Combo2.AddItem "LALAPAÞA"
Combo2.AddItem "MERÝÇ"
Combo2.AddItem "SÜLOÐLU"
Combo2.AddItem "UZUNKÖPRÜ"
ElseIf Combo1.ListIndex = 28 Then
Combo2.Clear
Combo2.AddItem "AÐIN"
Combo2.AddItem "ALACAKAYA"
Combo2.AddItem "ARICAK"
Combo2.AddItem "BASKÝL"
Combo2.AddItem "ELAZIÐ MERKEZ"
Combo2.AddItem "KARAKOÇAN"
Combo2.AddItem "KEBAN"
Combo2.AddItem "KOVANCILAR"
Combo2.AddItem "MADEN"
Combo2.AddItem "PALU"
Combo2.AddItem "SÝVRÝCE"
ElseIf Combo1.ListIndex = 29 Then
Combo2.Clear
Combo2.AddItem "ÇAYIRLI"
Combo2.AddItem "ERZÝNCAN MERKEZ"
Combo2.AddItem "ÝLÝÇ"
Combo2.AddItem "KEMAH"
Combo2.AddItem "KEMALÝYE"
Combo2.AddItem "OTLUKBELÝ"
Combo2.AddItem "REFAHÝYE"
Combo2.AddItem "TERCAN"
Combo2.AddItem "ÜZÜMLÜ"
ElseIf Combo1.ListIndex = 30 Then
Combo2.Clear
Combo2.AddItem "AÞKALE"
Combo2.AddItem "ÇAT"
Combo2.AddItem "ERZURUM MERKEZ"
Combo2.AddItem "HINIS"
Combo2.AddItem "HORASAN"
Combo2.AddItem "ILICA"
Combo2.AddItem "ÝSPÝR"
Combo2.AddItem "KARAÇOBAN"
Combo2.AddItem "KARAYAZI"
Combo2.AddItem "KÖPRÜKÖY"
Combo2.AddItem "NARMAN"
Combo2.AddItem "OLTU"
Combo2.AddItem "OLUR"
Combo2.AddItem "PASÝNLER"
Combo2.AddItem "PAZARYOLU"
Combo2.AddItem "ÞENKAYA"
Combo2.AddItem "TEKMAN"
Combo2.AddItem "TORTUM"
Combo2.AddItem "UZUNDERE"
ElseIf Combo1.ListIndex = 31 Then
Combo2.Clear
Combo2.AddItem "ALPU"
Combo2.AddItem "BEYLÝKOVA"
Combo2.AddItem "ÇÝFTELER"
Combo2.AddItem "ESKÝÞEHÝR MERKEZ"
Combo2.AddItem "GÜNYÜZÜ"
Combo2.AddItem "HAN"
Combo2.AddItem "ÝNÖNÜ"
Combo2.AddItem "MAHMUDÝYE"
Combo2.AddItem "MÝHALGAZÝ"
Combo2.AddItem "MÝHALIÇÇIK"
Combo2.AddItem "SARICAKAYA"
Combo2.AddItem "SEYÝTGAZÝ"
Combo2.AddItem "SÝVRÝHÝSAR"
ElseIf Combo1.ListIndex = 32 Then
Combo2.Clear
Combo2.AddItem "ARABAN"
Combo2.AddItem "GAZÝANTEP MERKEZ"
Combo2.AddItem "ÝSLAHÝYE"
Combo2.AddItem "KARKAMIÞ"
Combo2.AddItem "NÝZÝP"
Combo2.AddItem "NURDAÐI"
Combo2.AddItem "OÐUZELÝ"
Combo2.AddItem "ÞAHÝNBEY"
Combo2.AddItem "ÞEHÝTKAMÝL"
Combo2.AddItem "YAVUZELÝ"
ElseIf Combo1.ListIndex = 33 Then
Combo2.Clear
Combo2.AddItem "ALUCRA"
Combo2.AddItem "BULANCAK"
Combo2.AddItem "ÇAMOLUK"
Combo2.AddItem "ÇANAKÇI"
Combo2.AddItem "DERELÝ"
Combo2.AddItem "DOÐANKENT"
Combo2.AddItem "ESPÝYE"
Combo2.AddItem "EYNESÝL"
Combo2.AddItem "GÝRESUN MERKEZ"
Combo2.AddItem "GÖRELE"
Combo2.AddItem "GÜCE"
Combo2.AddItem "KEÞAP"
Combo2.AddItem "PÝRAZÝZ"
Combo2.AddItem "ÞEBÝNKARAHÝSAR"
Combo2.AddItem "TÝREBOLU"
Combo2.AddItem "YAÐLIDERE"
ElseIf Combo1.ListIndex = 34 Then
Combo2.Clear
Combo2.AddItem "GÜMÜÞHANE MERKEZ"
Combo2.AddItem "KELKÝT"
Combo2.AddItem "KÖSE"
Combo2.AddItem "KÜRTÜN"
Combo2.AddItem "ÞÝRAN"
Combo2.AddItem "TORUL"
ElseIf Combo1.ListIndex = 35 Then
Combo2.Clear
Combo2.AddItem "ÇUKURCA"
Combo2.AddItem "HAKKARÝ MERKEZ"
Combo2.AddItem "ÞEMDÝNLÝ"
Combo2.AddItem "YÜKSEKOVA"
ElseIf Combo1.ListIndex = 36 Then
Combo2.Clear
Combo2.AddItem "ALTINÖZÜ"
Combo2.AddItem "BELEN"
Combo2.AddItem "DÖRTYOL"
Combo2.AddItem "ERZÝN"
Combo2.AddItem "HASSA"
Combo2.AddItem "HATAY MERKEZ"
Combo2.AddItem "ÝSKENDERUN"
Combo2.AddItem "KIRIKHAN"
Combo2.AddItem "KUMLU"
Combo2.AddItem "REYHANLI"
Combo2.AddItem "SAMANDAÐ"
Combo2.AddItem "YAYLADAÐ"
ElseIf Combo1.ListIndex = 37 Then
Combo2.Clear
Combo2.AddItem "ARALIK"
Combo2.AddItem "IÐDIR MERKEZ"
Combo2.AddItem "KARAKOYUNLU"
Combo2.AddItem "TUZLUCA"
ElseIf Combo1.ListIndex = 38 Then
Combo2.Clear
Combo2.AddItem "AKSU"
Combo2.AddItem "ATABEY"
Combo2.AddItem "EÐÝRDÝR"
Combo2.AddItem "GELENDOST"
Combo2.AddItem "GÖNEN/ISPARTA"
Combo2.AddItem "ISPARTA MERKEZ"
Combo2.AddItem "KEÇÝBORLU"
Combo2.AddItem "ÞARKÝKARAAÐAÇ"
Combo2.AddItem "SENÝRKENT"
Combo2.AddItem "SÜTÇÜLER"
Combo2.AddItem "ULUBORLU"
Combo2.AddItem "YALVAÇ"
Combo2.AddItem "YENÝÞARBADEMLÝ"
ElseIf Combo1.ListIndex = 39 Then
Combo2.Clear
Combo2.AddItem "ADALAR"
Combo2.AddItem "AVCILAR"
Combo2.AddItem "BAÐCILAR"
Combo2.AddItem "BAHÇELÝEVLER"
Combo2.AddItem "BAKIRKÖY"
Combo2.AddItem "BAYRAMPAÞA"
Combo2.AddItem "BEÞÝKTAÞ"
Combo2.AddItem "BEYKOZ"
Combo2.AddItem "BEYOÐLU"
Combo2.AddItem "BÜYÜKÇEKMECE"
Combo2.AddItem "ÇATALCA"
Combo2.AddItem "EMÝNÖNÜ"
Combo2.AddItem "ESENLER"
Combo2.AddItem "EYÜP"
Combo2.AddItem "FATÝH"
Combo2.AddItem "GAZÝOSMANPAÞA"
Combo2.AddItem "GÜNGÖREN"
Combo2.AddItem "ÝSTANBUL MERKEZ"
Combo2.AddItem "KADIKÖY"
Combo2.AddItem "KAÐITHANE"
Combo2.AddItem "KARTAL"
Combo2.AddItem "KÜÇÜKÇEKMECE"
Combo2.AddItem "MALTEPE"
Combo2.AddItem "PENDÝK"
Combo2.AddItem "SARIYER"
Combo2.AddItem "SÝLÝVRÝ"
Combo2.AddItem "SULTANBEYLÝ"
Combo2.AddItem "ÞÝLE"
Combo2.AddItem "ÞÝÞLÝ"
Combo2.AddItem "TUZLA"
Combo2.AddItem "ÜMRANÝYE"
Combo2.AddItem "ÜSKÜDAR"
Combo2.AddItem "ZEYTÝNBURNU"
ElseIf Combo1.ListIndex = 40 Then
Combo2.Clear
Combo2.AddItem "ALÝAÐA"
Combo2.AddItem "BALÇOVA"
Combo2.AddItem "BAYINDIR"
Combo2.AddItem "BERGAMA"
Combo2.AddItem "BEYDAÐ"
Combo2.AddItem "BORNOVA"
Combo2.AddItem "BUCA"
Combo2.AddItem "ÇEÞME"
Combo2.AddItem "ÇÝÐLÝ"
Combo2.AddItem "DÝKÝLÝ"
Combo2.AddItem "FOÇA"
Combo2.AddItem "GAZÝEMÝR"
Combo2.AddItem "GÜZELBAHÇE"
Combo2.AddItem "ÝZMÝR MERKEZ"
Combo2.AddItem "KARABURUN"
Combo2.AddItem "KARÞIYAKA"
Combo2.AddItem "KEMALPAÞA"
Combo2.AddItem "KINIK"
Combo2.AddItem "KÝRAZ"
Combo2.AddItem "KONAK"
Combo2.AddItem "MENDERES"
Combo2.AddItem "MENEMEN"
Combo2.AddItem "NARLIDERE"
Combo2.AddItem "ÖDEMÝÞ"
Combo2.AddItem "SEFERÝHÝSAR"
Combo2.AddItem "SELÇUK"
Combo2.AddItem "TÝRE"
Combo2.AddItem "TORBALI"
Combo2.AddItem "URLA"
ElseIf Combo1.ListIndex = 41 Then
Combo2.Clear
Combo2.AddItem "AFÞÝN"
Combo2.AddItem "ANDIRIN"
Combo2.AddItem "ÇAÐLIYANCERÝT"
Combo2.AddItem "EKÝNÖZÜ"
Combo2.AddItem "ELBÝSTAN"
Combo2.AddItem "GÖKSUN"
Combo2.AddItem "KAHRAMANMARAÞ MERKEZ"
Combo2.AddItem "NURHAK"
Combo2.AddItem "PAZARCIK"
Combo2.AddItem "TÜRKOÐLU"
ElseIf Combo1.ListIndex = 42 Then
Combo2.Clear
Combo2.AddItem "EFLANÝ"
Combo2.AddItem "ESKÝPAZAR"
Combo2.AddItem "KARABÜK MERKEZ"
Combo2.AddItem "OVACIK/KARABÜK"
Combo2.AddItem "SAFRANBOLU"
Combo2.AddItem "YENÝCE/KARABÜK"
ElseIf Combo1.ListIndex = 43 Then
Combo2.Clear
Combo2.AddItem "AYRANCI"
Combo2.AddItem "BAÞYAYLA"
Combo2.AddItem "ERMENEK"
Combo2.AddItem "KARAMAN MERKEZ"
Combo2.AddItem "KAZIMKARABEKÝR"
Combo2.AddItem "SARIVELÝLER"
ElseIf Combo1.ListIndex = 44 Then
Combo2.Clear
Combo2.AddItem "AKYAKA"
Combo2.AddItem "ARPAÇAY"
Combo2.AddItem "DÝGOR"
Combo2.AddItem "KAÐIZMAN"
Combo2.AddItem "KARS MERKEZ"
Combo2.AddItem "SARIKAMIÞ"
Combo2.AddItem "SELÝM"
Combo2.AddItem "SUSUZ"
ElseIf Combo1.ListIndex = 45 Then
Combo2.Clear
Combo2.AddItem "ABANA"
Combo2.AddItem "AÐLI"
Combo2.AddItem "ARAÇ"
Combo2.AddItem "AZDAVAY"
Combo2.AddItem "BOZKURT/KASTAMONU"
Combo2.AddItem "ÇATALZEYTÝN"
Combo2.AddItem "CÝDE"
Combo2.AddItem "DADAY"
Combo2.AddItem "DEVREKANÝ"
Combo2.AddItem "DOÐANYURT"
Combo2.AddItem "HANÖNÜ"
Combo2.AddItem "ÝHSANGAZÝ"
Combo2.AddItem "ÝNEBOLU"
Combo2.AddItem "KASTAMONU MERKEZ"
Combo2.AddItem "KÜRE"
Combo2.AddItem "PINARBAÞI/KASTAMONU"
Combo2.AddItem "SEYDÝLER"
Combo2.AddItem "ÞENPAZAR"
Combo2.AddItem "TAÞKÖPRÜ"
Combo2.AddItem "TOSYA"
ElseIf Combo1.ListIndex = 46 Then
Combo2.Clear
Combo2.AddItem "AKKIÞLA"
Combo2.AddItem "BÜNYAN"
Combo2.AddItem "DEVELÝ"
Combo2.AddItem "FELAHÝYE"
Combo2.AddItem "HACILAR"
Combo2.AddItem "ÝNCESU"
Combo2.AddItem "KAYSERÝ MERKEZ"
Combo2.AddItem "KOCASÝNAN"
Combo2.AddItem "MELÝKGAZÝ"
Combo2.AddItem "ÖZVATAN"
Combo2.AddItem "PINARBAÞI/KAYSERÝ"
Combo2.AddItem "SARIOÐLAN"
Combo2.AddItem "SARIZ"
Combo2.AddItem "TALAS"
Combo2.AddItem "TOMARZA"
Combo2.AddItem "YAHYALI"
Combo2.AddItem "YEÞÝLHÝSAR"
ElseIf Combo1.ListIndex = 47 Then
Combo2.Clear
Combo2.AddItem "BAHÞÝLÝ"
Combo2.AddItem "BALIÞEYH"
Combo2.AddItem "ÇELEBÝ"
Combo2.AddItem "DELÝCE"
Combo2.AddItem "KARAKEÇÝLÝ"
Combo2.AddItem "KESKÝN"
Combo2.AddItem "KIRIKKALE MERKEZ"
Combo2.AddItem "SULAKYURT"
Combo2.AddItem "YAHÞÝHAN"
ElseIf Combo1.ListIndex = 48 Then
Combo2.Clear
Combo2.AddItem "BABAESKÝ"
Combo2.AddItem "DEMÝRKÖY"
Combo2.AddItem "KIRKLARELÝ MERKEZ"
Combo2.AddItem "KOFÇAZ"
Combo2.AddItem "LÜLEBURGAZ"
Combo2.AddItem "PEHLÝVANKÖY"
Combo2.AddItem "PINARHÝSAR"
Combo2.AddItem "VÝZE"
ElseIf Combo1.ListIndex = 49 Then
Combo2.Clear
Combo2.AddItem "AKÇAKENT"
Combo2.AddItem "AKPINAR"
Combo2.AddItem "BOZTEPE"
Combo2.AddItem "ÇÝÇEKDAÐI"
Combo2.AddItem "KAMAN"
Combo2.AddItem "KIRÞEHÝR MERKEZ"
Combo2.AddItem "MUCUR"
ElseIf Combo1.ListIndex = 50 Then
Combo2.Clear
Combo2.AddItem "ELBEYLÝ"
Combo2.AddItem "KÝLÝS MERKEZ"
Combo2.AddItem "MUSABEYLÝ"
Combo2.AddItem "POLATELÝ"
ElseIf Combo1.ListIndex = 51 Then
Combo2.Clear
Combo2.AddItem "DERÝNCE"
Combo2.AddItem "GEBZE"
Combo2.AddItem "GÖLCÜK"
Combo2.AddItem "KANDIRA"
Combo2.AddItem "KARAMÜRSEL"
Combo2.AddItem "KOCAELÝ MERKEZ"
Combo2.AddItem "KÖRFEZ"
ElseIf Combo1.ListIndex = 52 Then
Combo2.Clear
Combo2.AddItem "AHIRLI"
Combo2.AddItem "AKÖREN"
Combo2.AddItem "AKÞEHÝR"
Combo2.AddItem "ALTINEKÝN"
Combo2.AddItem "BEYÞEHÝR"
Combo2.AddItem "BOZKIR"
Combo2.AddItem "ÇELTÝK"
Combo2.AddItem "CÝHANBEYLÝ"
Combo2.AddItem "ÇUMRA"
Combo2.AddItem "DERBENT"
Combo2.AddItem "DEREBUCAK"
Combo2.AddItem "DOÐANHÝSAR"
Combo2.AddItem "EMÝRGAZÝ"
Combo2.AddItem "EREÐLÝ/KONYA"
Combo2.AddItem "GÜNEYSINIR"
Combo2.AddItem "HADÝM"
Combo2.AddItem "HALKAPINAR"
Combo2.AddItem "HÜYÜK"
Combo2.AddItem "ILGIN"
Combo2.AddItem "KADINHANI"
Combo2.AddItem "KARAPINAR"
Combo2.AddItem "KARATAY"
Combo2.AddItem "KONYA MERKEZ"
Combo2.AddItem "KULU"
Combo2.AddItem "MERAM"
Combo2.AddItem "SARAYÖNÜ"
Combo2.AddItem "SELÇUKLU"
Combo2.AddItem "SEYDÝÞEHÝR"
Combo2.AddItem "TAÞKENT"
Combo2.AddItem "TUZLUKÇU"
Combo2.AddItem "YALIHÜYÜK"
Combo2.AddItem "YUNAK"
ElseIf Combo1.ListIndex = 53 Then
Combo2.Clear
Combo2.AddItem "ALTINTAÞ"
Combo2.AddItem "ASLANAPA"
Combo2.AddItem "ÇAVDARHÝSAR"
Combo2.AddItem "DOMANÝÇ"
Combo2.AddItem "DUMLUPINAR"
Combo2.AddItem "EMET"
Combo2.AddItem "GEDÝZ"
Combo2.AddItem "HÝSARCIK"
Combo2.AddItem "KÜTAHYA MERKEZ"
Combo2.AddItem "PAZARLAR"
Combo2.AddItem "SÝMAV"
Combo2.AddItem "ÞAPHANE"
Combo2.AddItem "TAVÞANLI"
ElseIf Combo1.ListIndex = 54 Then
Combo2.Clear
Combo2.AddItem "AKÇADAÐ"
Combo2.AddItem "ARAPGÝR"
Combo2.AddItem "ARGUVAN"
Combo2.AddItem "BATTALGAZÝ"
Combo2.AddItem "DARENDE"
Combo2.AddItem "DOÐANÞEHÝR"
Combo2.AddItem "DOÐANYOL"
Combo2.AddItem "HEKÝMHAN"
Combo2.AddItem "KALE/MALATYA"
Combo2.AddItem "KULUNCAK"
Combo2.AddItem "MALATYA MERKEZ"
Combo2.AddItem "PÜTÜRGE"
Combo2.AddItem "YAZIHAN"
Combo2.AddItem "YEÞÝLYURT/MALATYA"
ElseIf Combo1.ListIndex = 55 Then
Combo2.Clear
Combo2.AddItem "AHMETLÝ"
Combo2.AddItem "AKHÝSAR"
Combo2.AddItem "ALAÞEHÝR"
Combo2.AddItem "DEMÝRCÝ"
Combo2.AddItem "GÖLMARMARA"
Combo2.AddItem "GÖRDES"
Combo2.AddItem "KIRKAÐAÇ"
Combo2.AddItem "KÖPRÜBAÞI/MANÝSA"
Combo2.AddItem "KULA"
Combo2.AddItem "MANÝSA MERKEZ"
Combo2.AddItem "SALÝHLÝ"
Combo2.AddItem "SARIGÖL"
Combo2.AddItem "SARUHANLI"
Combo2.AddItem "SELENDÝ"
Combo2.AddItem "SOMA"
Combo2.AddItem "TURGUTLU"
ElseIf Combo1.ListIndex = 56 Then
Combo2.Clear
Combo2.AddItem "DARGEÇÝT"
Combo2.AddItem "DERÝK"
Combo2.AddItem "KIZILTEPE"
Combo2.AddItem "MARDÝN MERKEZ"
Combo2.AddItem "MAZIDAÐI"
Combo2.AddItem "MÝDYAT"
Combo2.AddItem "NUSAYBÝN"
Combo2.AddItem "ÖMERLÝ"
Combo2.AddItem "SAVUR"
Combo2.AddItem "YEÞÝLLÝ"
ElseIf Combo1.ListIndex = 57 Then
Combo2.Clear
Combo2.AddItem "ANAMUR"
Combo2.AddItem "AYDINCIK/MERSÝN"
Combo2.AddItem "BOZYAZI"
Combo2.AddItem "ÇAMLIYAYLA"
Combo2.AddItem "ERDEMLÝ"
Combo2.AddItem "GÜLNAR"
Combo2.AddItem "MERSÝN MERKEZ"
Combo2.AddItem "MUT"
Combo2.AddItem "SÝLÝFKE"
Combo2.AddItem "TARSUS"
ElseIf Combo1.ListIndex = 58 Then
Combo2.Clear
Combo2.AddItem "BODRUM"
Combo2.AddItem "DALAMAN"
Combo2.AddItem "DATÇA"
Combo2.AddItem "FETHÝYE"
Combo2.AddItem "KAVAKLIDERE"
Combo2.AddItem "KÖYCEÐÝZ"
Combo2.AddItem "MARMARÝS"
Combo2.AddItem "MÝLAS"
Combo2.AddItem "MUÐLA MERKEZ"
Combo2.AddItem "ORTACA"
Combo2.AddItem "ULA"
Combo2.AddItem "YATAÐAN"
ElseIf Combo1.ListIndex = 59 Then
Combo2.Clear
Combo2.AddItem "BULANIK"
Combo2.AddItem "HASKÖY"
Combo2.AddItem "KORKUT"
Combo2.AddItem "MALAZGÝRT"
Combo2.AddItem "MUÞ MERKEZ"
Combo2.AddItem "VARTO"
ElseIf Combo1.ListIndex = 60 Then
Combo2.Clear
Combo2.AddItem "ACIGÖL"
Combo2.AddItem "AVONOS"
Combo2.AddItem "DERÝNKUYU"
Combo2.AddItem "GÜLÞEHÝR"
Combo2.AddItem "HACIBEKTAÞ"
Combo2.AddItem "KOZAKLI"
Combo2.AddItem "NEVÞEHÝR MERKEZ"
Combo2.AddItem "ÜRGÜP"
ElseIf Combo1.ListIndex = 61 Then
Combo2.Clear
Combo2.AddItem "ALTUNHÝSAR"
Combo2.AddItem "BOR"
Combo2.AddItem "ÇAMARDI"
Combo2.AddItem "ÇÝFTLÝK"
Combo2.AddItem "NÝÐDE MERKEZ"
Combo2.AddItem "ULUKIÞLA"
ElseIf Combo1.ListIndex = 62 Then
Combo2.Clear
Combo2.AddItem "AKKUÞ"
Combo2.AddItem "AYBASTI"
Combo2.AddItem "ÇAMAÞ"
Combo2.AddItem "ÇATALPINAR"
Combo2.AddItem "ÇAYBAÞI"
Combo2.AddItem "FATSA"
Combo2.AddItem "GÖLKÖY"
Combo2.AddItem "GÜLYALI"
Combo2.AddItem "GÜRGENTEPE"
Combo2.AddItem "ÝKÝZCE"
Combo2.AddItem "KABADÜZ"
Combo2.AddItem "KABATAÞ"
Combo2.AddItem "KORGAN"
Combo2.AddItem "KUMRU"
Combo2.AddItem "MESUDÝYE"
Combo2.AddItem "ORDU MERKEZ"
Combo2.AddItem "PERÞEMBE"
Combo2.AddItem "ULUBEY/ORDU"
Combo2.AddItem "ÜNYE"
ElseIf Combo1.ListIndex = 63 Then
Combo2.Clear
Combo2.AddItem "BAHÇE"
Combo2.AddItem "DÜZÝÇÝ"
Combo2.AddItem "HASANBEYLÝ"
Combo2.AddItem "KADÝRLÝ"
Combo2.AddItem "OSMANÝYE MERKEZ"
Combo2.AddItem "SUMBAS"
Combo2.AddItem "TOPRAKKALE"
ElseIf Combo1.ListIndex = 64 Then
Combo2.Clear
Combo2.AddItem "ARDEÞEN"
Combo2.AddItem "ÇAMLIHEMÞÝN"
Combo2.AddItem "ÇAYELÝ"
Combo2.AddItem "DEREPAZARI"
Combo2.AddItem "FINDIKLI"
Combo2.AddItem "GÜNEYSU"
Combo2.AddItem "HEMÞÝN"
Combo2.AddItem "ÝKÝZDERE"
Combo2.AddItem "ÝYÝDERE"
Combo2.AddItem "KALKANDERE"
Combo2.AddItem "PAZAR/RÝZE"
Combo2.AddItem "RÝZE MERKEZ"
ElseIf Combo1.ListIndex = 65 Then
Combo2.Clear
Combo2.AddItem "AKYAZI"
Combo2.AddItem "FERÝZLÝ"
Combo2.AddItem "GEYVE"
Combo2.AddItem "HENDEK"
Combo2.AddItem "KARAPÜRÇEK"
Combo2.AddItem "KARASU"
Combo2.AddItem "KAYNARCA"
Combo2.AddItem "KOCAALÝ"
Combo2.AddItem "PAMUKOVA"
Combo2.AddItem "SAKARYA MERKEZ"
Combo2.AddItem "SAPANCA"
Combo2.AddItem "SÖÐÜTLÜ"
Combo2.AddItem "TARAKLI"
ElseIf Combo1.ListIndex = 66 Then
Combo2.Clear
Combo2.AddItem "ALAÇAM"
Combo2.AddItem "ASARCIK"
Combo2.AddItem "AYVACIK/SAMSUN"
Combo2.AddItem "BAFRA"
Combo2.AddItem "ÇARÞAMBA"
Combo2.AddItem "HAVZA"
Combo2.AddItem "KAVAK"
Combo2.AddItem "LADÝK"
Combo2.AddItem "ONDOKUZMAYIS"
Combo2.AddItem "SALIPAZARI"
Combo2.AddItem "SAMSUN MERKEZ"
Combo2.AddItem "TEKKEKÖY"
Combo2.AddItem "TERME"
Combo2.AddItem "VEZÝRKÖPRÜ"
Combo2.AddItem "YAKAKENT"
ElseIf Combo1.ListIndex = 67 Then
Combo2.Clear
Combo2.AddItem "AYDINLAR"
Combo2.AddItem "BAYKAN"
Combo2.AddItem "ERUH"
Combo2.AddItem "KURTALAN"
Combo2.AddItem "PERVARÝ"
Combo2.AddItem "SÝÝRT MERKEZ"
Combo2.AddItem "ÞÝRVAN"
ElseIf Combo1.ListIndex = 68 Then
Combo2.Clear
Combo2.AddItem "AYANCIK"
Combo2.AddItem "BOYABAT"
Combo2.AddItem "DÝKMEN"
Combo2.AddItem "DURAÐAN"
Combo2.AddItem "ERFELEK"
Combo2.AddItem "GERZE"
Combo2.AddItem "SARAYDÜZÜ"
Combo2.AddItem "SÝNOP MERKEZ"
Combo2.AddItem "TÜRKELÝ"
ElseIf Combo1.ListIndex = 69 Then
Combo2.Clear
Combo2.AddItem "AKINCILAR"
Combo2.AddItem "ALTINYAYLA/SÝVAS"
Combo2.AddItem "DÝVRÝÐÝ"
Combo2.AddItem "DOÐANÞAR"
Combo2.AddItem "GEMEREK"
Combo2.AddItem "GÜLOVA"
Combo2.AddItem "GÜRÜN"
Combo2.AddItem "HAFÝK"
Combo2.AddItem "ÝMRANLI"
Combo2.AddItem "KANGAL"
Combo2.AddItem "KOYULHÝSAR"
Combo2.AddItem "SÝVAS MERKEZ"
Combo2.AddItem "SUÞEHRÝ"
Combo2.AddItem "ÞARKIÞLA"
Combo2.AddItem "ULAÞ"
Combo2.AddItem "YILDIZELÝ"
Combo2.AddItem "ZARA"
ElseIf Combo1.ListIndex = 70 Then
Combo2.Clear
Combo2.AddItem "AKÇAKALE"
Combo2.AddItem "BÝRECÝK"
Combo2.AddItem "BOZOVA"
Combo2.AddItem "CEYLANPINAR"
Combo2.AddItem "HALFETÝ"
Combo2.AddItem "HARRAN"
Combo2.AddItem "HÝLVAN"
Combo2.AddItem "SÝVEREK"
Combo2.AddItem "SURUÇ"
Combo2.AddItem "ÞANLIURFA MERKEZ"
Combo2.AddItem "VÝRANÞEHÝR"
ElseIf Combo1.ListIndex = 71 Then
Combo2.Clear
Combo2.AddItem "BEYTÜÞÞEBAP"
Combo2.AddItem "CÝZRE"
Combo2.AddItem "GÜÇLÜKONAK"
Combo2.AddItem "ÝDÝL"
Combo2.AddItem "SÝLOPÝ"
Combo2.AddItem "ÞIRNAK MERKEZ"
Combo2.AddItem "ULUDERE"

ElseIf Combo1.ListIndex = 72 Then
Combo2.Clear
Combo2.AddItem "ÇERKEZKÖY"
Combo2.AddItem "ÇORLU"
Combo2.AddItem "HAYRABOLU"
Combo2.AddItem "MALKARA"
Combo2.AddItem "MARMARAEREÐLÝSÝ"
Combo2.AddItem "MURATLI"
Combo2.AddItem "SARAY/TEKÝRDAÐ"
Combo2.AddItem "ÞARKÖY"
Combo2.AddItem "TEKÝRDAÐ MERKEZ"
ElseIf Combo1.ListIndex = 73 Then
Combo2.Clear
Combo2.AddItem "ALMUS"
Combo2.AddItem "ARTOVA"
Combo2.AddItem "BAÞÇÝFTLÝK"
Combo2.AddItem "ERBAA"
Combo2.AddItem "NÝKSAR"
Combo2.AddItem "PAZAR/TOKAT"
Combo2.AddItem "REÞADÝYE"
Combo2.AddItem "SULUSARAY"
Combo2.AddItem "TOKAT MERKEZ"
Combo2.AddItem "TURHAL"
Combo2.AddItem "YEÞÝLYURT/TOKAT"
Combo2.AddItem "ZÝLE"
ElseIf Combo1.ListIndex = 74 Then
Combo2.Clear
Combo2.AddItem "AKÇAABAT"
Combo2.AddItem "ARAKLI"
Combo2.AddItem "ARSÝN"
Combo2.AddItem "BEÞÝKDÜZÜ"
Combo2.AddItem "ÇARÞIBAÞI"
Combo2.AddItem "ÇAYKARA"
Combo2.AddItem "DERNEPAZARI"
Combo2.AddItem "DÜZKÖY"
Combo2.AddItem "HAYRAT"
Combo2.AddItem "KÖPRÜBAÞI/TRABZON"
Combo2.AddItem "MAÇKA"
Combo2.AddItem "OF"
Combo2.AddItem "ÞALPAZARI"
Combo2.AddItem "SÜRMENE"
Combo2.AddItem "TONYA"
Combo2.AddItem "TRABZON MERKEZ"
Combo2.AddItem "VAKFIKEBÝR"
Combo2.AddItem "YOMRA"
ElseIf Combo1.ListIndex = 75 Then
Combo2.Clear
Combo2.AddItem "ÇEMÝÞGEZEK"
Combo2.AddItem "HOZAT"
Combo2.AddItem "MAZGÝRT"
Combo2.AddItem "NAZÝMÝYE"
Combo2.AddItem "OVACIK/TUNCELÝ"
Combo2.AddItem "PERTEK"
Combo2.AddItem "PÜLÜMBÜR"
Combo2.AddItem "TUNCELÝ MERKEZ"
ElseIf Combo1.ListIndex = 76 Then
Combo2.Clear
Combo2.AddItem "BANAZ"
Combo2.AddItem "EÞME"
Combo2.AddItem "KARAHALLI"
Combo2.AddItem "SÝVASLI"
Combo2.AddItem "ULUBEY/UÞAK"
Combo2.AddItem "UÞAK MERKEZ"
ElseIf Combo1.ListIndex = 77 Then
Combo2.Clear
Combo2.AddItem "BAHÇESARAY"
Combo2.AddItem "BAÞKALE"
Combo2.AddItem "ÇALDIRAN"
Combo2.AddItem "ÇATAK"
Combo2.AddItem "EDREMÝT/VAN"
Combo2.AddItem "ERCÝÞ"
Combo2.AddItem "GEVAÞ"
Combo2.AddItem "GÜRPINAR"
Combo2.AddItem "MURADÝYE"
Combo2.AddItem "ÖZALP"
Combo2.AddItem "SARAY/VAN"
Combo2.AddItem "VAN MERKEZ"
ElseIf Combo1.ListIndex = 78 Then
Combo2.Clear
Combo2.AddItem "ALTINOVA"
Combo2.AddItem "ARMUTLU"
Combo2.AddItem "ÇÝFTLÝKKÖY"
Combo2.AddItem "ÇINARCIK"
Combo2.AddItem "TERMAL"
Combo2.AddItem "YALOVA MERKEZ"
ElseIf Combo1.ListIndex = 79 Then
Combo2.Clear
Combo2.AddItem "AKDAÐMADENÝ"
Combo2.AddItem "AYDINCIK/YOZGAT"
Combo2.AddItem "BOÐAZLIYAN"
Combo2.AddItem "ÇANDIR"
Combo2.AddItem "ÇAYIRALAN"
Combo2.AddItem "ÇEKEREK"
Combo2.AddItem "KADIÞEHRÝ"
Combo2.AddItem "SARAYKENT"
Combo2.AddItem "SARIKAYA"
Combo2.AddItem "ÞEFAATLÝ"
Combo2.AddItem "SORGUN"
Combo2.AddItem "YENÝFAKILI"
Combo2.AddItem "YERKÖY"
Combo2.AddItem "YOZGAT MERKEZ"
ElseIf Combo1.ListIndex = 80 Then
Combo2.AddItem "ALAPLI"
Combo2.AddItem "ÇAYCUMA"
Combo2.AddItem "DEVREK"
Combo2.AddItem "EREÐLÝ/ZONGULDAK"
Combo2.AddItem "GÖKÇEBEY"
Combo2.AddItem "ZONGULDAK MERKEZ"
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Len(Text14) = 0 Then
ElseIf Len(Text14) < 10 Then
MsgBox "TELEFON NUMARASI 10 HANELÝ OLMALIDIR."
Text14.Text = ""
Text14.BackColor = vbYellow
End If

If Len(Text5) = 0 Then
ElseIf Len(Text5) < 11 Then
MsgBox "T.C KÝMLÝK NUMARASI NUMARASI 11 HANELÝ OLMALIDIR."
Text5.Text = ""
Text5.BackColor = vbYellow
End If


If Text1 = "" Then
Text1.BackColor = vbYellow
a = a + 1
Else
Text1.BackColor = vbWhite
End If

If Text2 = "" Then
Text2.BackColor = vbYellow
a = a + 1
Else
Text2.BackColor = vbWhite
End If

If Text3 = "" Then
Text3.BackColor = vbYellow
a = a + 1
Else
Text3.BackColor = vbWhite
End If

If Text4 = "" Then
Text4.BackColor = vbYellow
a = a + 1
Else
Text4.BackColor = vbWhite
End If

If Text5 = "" Then
Text5.BackColor = vbYellow
a = a + 1
Else
Text5.BackColor = vbWhite
End If

If Combo1.Text = "" Then
Combo1.BackColor = vbYellow
a = a + 1
Else
Combo1.BackColor = vbWhite
End If

If Combo2.Text = "" Then
Combo2.BackColor = vbYellow
a = a + 1
Else
Combo2.BackColor = vbWhite
End If

If Text8 = "" Then
Text8.BackColor = vbYellow
a = a + 1
Else
Text8.BackColor = vbWhite
End If

If Text9 = "" Then
Text9.BackColor = vbYellow
a = a + 1
Else
Text9.BackColor = vbWhite
End If

If Text10 = "" Then
Text10.BackColor = vbYellow
a = a + 1
Else
Text10.BackColor = vbWhite
End If

If Text11 = "" Then
Text11.BackColor = vbYellow
a = a + 1
Else
Text11.BackColor = vbWhite
End If

If Combo3.Text = "" Then
Combo3.BackColor = vbYellow
a = a + 1
Else
Combo3.BackColor = vbWhite
End If

If Combo4.Text = "" Then
Combo4.BackColor = vbYellow
a = a + 1
Else
Combo4.BackColor = vbWhite
End If

If Text14 = "" Then
Text14.BackColor = vbYellow
a = a + 1
Else
Text14.BackColor = vbWhite
End If

If Text15 = "" Then
Text15.BackColor = vbYellow
a = a + 1
Else
Text15.BackColor = vbWhite
End If

If Text16 = "" Then
Text16.BackColor = vbYellow
a = a + 1
Else
Text16.BackColor = vbWhite
End If

If Text17 = "" Then
Text17.BackColor = vbYellow
a = a + 1
Else
Text17.BackColor = vbWhite
End If

If Text6 = "" Then
Text6.BackColor = vbYellow
a = a + 1
Else
Text6.BackColor = vbWhite
End If

If Label23.Caption = "" Then
Label23.BackColor = vbYellow
a = a + 1
Else
Label23.BackColor = vbWhite
End If



Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
c = rs.RecordCount
For sayac = 0 To c
If Text17 <> rs!OdaNo Then
rs.MoveNext
Else
MsgBox "ODA DOLUDUR.LÜTFEN BAÞKA ODA NUMARASI GÝRÝNÝZ."
Text17 = ""
Text17.BackColor = vbYellow
a = a + 1
b = 1
End If
Next sayac
If a > 0 Then
MsgBox a & "tane boþ alan býrakýlmýþ"
Label18.Visible = True
ElseIf a <= 0 And b <> 1 Then
MsgBox "boþ alan býrakmadýnýz"
Label18.Visible = False
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
rs.AddNew
rs.Fields("Adý") = Text1.Text
rs.Fields("soyadý") = Text2.Text
rs.Fields("Baba_adý") = Text3.Text
rs.Fields("Anne_adý") = Text4.Text
rs.Fields("Tc") = Text5.Text
rs.Fields("il") = Combo1.Text
rs.Fields("ilçe") = Combo2.Text
rs.Fields("Mahalle_Köy") = Text8.Text
rs.Fields("Ýkametgah_Adresi") = Text9.Text
rs.Fields("Doðum_Yeri") = Text10.Text
rs.Fields("Doðum_Tarih") = Text11.Text
rs.Fields("Cinsiyet") = Combo3.Text
rs.Fields("Medeni_Hali") = Combo4.Text
rs.Fields("Telefon") = Text14.Text
rs.Fields("Mesleði") = Text15.Text
rs.Fields("E_Posta") = Text16.Text
rs.Fields("OdaNo") = Text17.Text
rs.Fields("Geliþ_Tarihi") = Label19.Caption
rs.Fields("Geliþ_Saati") = Label20.Caption
rs.Fields("Gün") = Text6.Text
rs.Fields("Fiyat") = Label23.Caption
rs.Update
rs.Close
MsgBox "müþteri kayýdý yapýldý."
Command3.Visible = True
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Combo1.Text = ""
Combo2.Text = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text11 = ""
Combo3.Text = ""
Combo4.Text = ""
Text14 = ""
Text15 = ""
Text16 = ""
Text17 = ""
Text6 = ""
Label23.Caption = ""
End If
End Sub
Private Sub Command2_Click()
Calendar1.Visible = False
Calendar1.Value = Date
Label18.Visible = False
Form5.Hide
Form3.Show
Command3.Visible = False
Text1.BackColor = vbWhite
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
Text4.BackColor = vbWhite
Text5.BackColor = vbWhite
Combo1.BackColor = vbWhite
Combo2.BackColor = vbWhite
Text8.BackColor = vbWhite
Text9.BackColor = vbWhite
Text10.BackColor = vbWhite
Text11.BackColor = vbWhite
Combo3.BackColor = vbWhite
Combo4.BackColor = vbWhite
Text14.BackColor = vbWhite
Text15.BackColor = vbWhite
Text16.BackColor = vbWhite
Text17.BackColor = vbWhite
Text6.BackColor = vbWhite
Label23.BackColor = vbWhite

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text6.Text = ""
Label23.Caption = ""

End Sub

Private Sub Command3_Click()
Calendar1.Value = Date
Label18.Visible = False
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
If a >= 1 Then
rs.MoveFirst
Form4.Text1.Text = rs!Adý
Form4.Text2.Text = rs!soyadý
Form4.Text3.Text = rs!Baba_adý
Form4.Text4.Text = rs!Anne_adý
Form4.Text5.Text = rs!Tc
Form4.Text6.Text = rs!il
Form4.Text7.Text = rs!ilçe
Form4.Text8.Text = rs!Doðum_Yeri
Form4.Text9.Text = rs!Doðum_Tarih
Form4.Text10.Text = rs!Cinsiyet
Form4.Text11.Text = rs!Medeni_Hali
Form4.Text12.Text = rs!Mahalle_Köy
Form4.Label18.Caption = rs!Ýkametgah_Adresi
Form4.Label20.Caption = rs!E_Posta
Form4.Text14.Text = rs!Telefon
Form4.Text15.Text = rs!Mesleði
Form4.Text13.Text = rs!OdaNo
Form4.Text16.Text = rs!Geliþ_Tarihi
Form4.Text17.Text = rs!Geliþ_Saati
Form5.Hide
Form4.Show
End If
Command3.Visible = False
Text1.BackColor = vbWhite
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
Text4.BackColor = vbWhite
Text5.BackColor = vbWhite
Combo1.BackColor = vbWhite
Combo2.BackColor = vbWhite
Text8.BackColor = vbWhite
Text9.BackColor = vbWhite
Text10.BackColor = vbWhite
Text11.BackColor = vbWhite
Combo3.BackColor = vbWhite
Combo4.BackColor = vbWhite
Text14.BackColor = vbWhite
Text15.BackColor = vbWhite
Text16.BackColor = vbWhite
Text17.BackColor = vbWhite
Text6.BackColor = vbWhite
Label23.BackColor = vbWhite

End Sub
Private Sub Form_Load()
Label19.Caption = Date
Label20.Caption = Time
End Sub

Private Sub Text11_Click()
Calendar1.Visible = True
End Sub

Private Sub Text14_Change()
If Len(Text14) > 10 Then
MsgBox "TELEFON NUMARASI 10 HANELÝ OLMALIDIR."
Text14.Text = ""
End If
End Sub

Private Sub Text17_Change()
If Len(Text17) > 1 Then
MsgBox "9 TANE ODA MEVCUTTUR."
Text17.Text = ""
End If
End Sub

Private Sub Text5_Change()
If Len(Text5) > 11 Then
MsgBox "T.C KÝMLÝK NUMARASI NUMARASI 11 HANELÝ OLMALIDIR."
Text5.Text = ""
End If
End Sub

Private Sub Text6_Change()
On Error Resume Next
If Text6 = "" Then
Label23.Caption = ""
End If
Label23.Caption = (Text6 * 100) & "  TL"
End Sub

Private Sub Timer1_Timer()
Label19.Caption = Date
Label20.Caption = Time
End Sub
