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
         ToolTipText     =   "M��teri giri�i yapmak i�in t�klay�n."
         Top             =   7920
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ana Sayfaya D�n"
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
         ToolTipText     =   "Ana sayfaya d�nmek i�in t�klay�n."
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
         Caption         =   "M��teri Bilgilerini G�ster"
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
         ToolTipText     =   "M��teri bilgilerini g�rmek istiyorsan�z t�klay�n."
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
         Caption         =   "Konaklama G�n�===>"
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
         Caption         =   "Mahalle/K�y===>"
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
         Caption         =   "Mesle�i===>"
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
         Caption         =   "Telefon Numaras�===>"
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
         Caption         =   "�kametgah Adresi===>"
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
         Caption         =   "Do�um Tarihi===>"
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
         Caption         =   "Do�um Yeri===>"
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
         Caption         =   "�l�e===>"
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
         Caption         =   "�l===>"
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
         Caption         =   "T.C Kimlik Numaras�===>"
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
         Caption         =   "Anne Ad�==>"
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
         Caption         =   "Baba Ad�===>"
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
         Caption         =   "Soyad�===>"
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
         Caption         =   "BO� ALANLAR SARI RENKTE  G�STER�LM��T�R."
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
MsgBox "18 ya��ndan k���kler al�nmaz"
Text11 = ""
Calendar1.Visible = False
End If

End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
Combo2.Clear
Combo2.AddItem "ALADA�"
Combo2.AddItem "CEYHAN"
Combo2.AddItem "FEKE"
Combo2.AddItem "�MAMO�LU"
Combo2.AddItem "KARA�SALI"
Combo2.AddItem "KARATA�"
Combo2.AddItem "KOZAN"
Combo2.AddItem "POZANTI"
Combo2.AddItem "SA�MBEYL�"
Combo2.AddItem "SEYHAN"
Combo2.AddItem "TUFANBEYL�"
Combo2.AddItem "YUMURTALIK"
Combo2.AddItem "Y�RE��R"
ElseIf Combo1.ListIndex = 1 Then
Combo2.Clear
Combo2.AddItem "ADIYAMAN MERKEZ"
Combo2.AddItem "BESN�"
Combo2.AddItem "�EL�KHAN"
Combo2.AddItem "GERGER"
Combo2.AddItem "G�LBA�I/ADIYAMAN"
Combo2.AddItem "KAHTA"
Combo2.AddItem "SAMSAT"
Combo2.AddItem "S�NC�K"
Combo2.AddItem "TUT"
ElseIf Combo1.ListIndex = 2 Then
Combo2.Clear
Combo2.AddItem "AFYONKARAH�SAR MERKEZ"
Combo2.AddItem "BA�MAK�I"
Combo2.AddItem "BAYAT/AFYON"
Combo2.AddItem "BOLVAD�N"
Combo2.AddItem "�AY"
Combo2.AddItem "�OBANLAR"
Combo2.AddItem "DAZKIRI"
Combo2.AddItem "D�NAR"
Combo2.AddItem "EM�RDA�"
Combo2.AddItem "EVC�LER"
Combo2.AddItem "HOCALAR"
Combo2.AddItem "�HSAN�YE"
Combo2.AddItem "�SCEH�SAR"
Combo2.AddItem "KIZIL�REN"
Combo2.AddItem "SANDIKLI"
Combo2.AddItem "S�NANPA�A"
Combo2.AddItem "�UHUT"
Combo2.AddItem "SULTANDA�I"
ElseIf Combo1.ListIndex = 3 Then
Combo2.Clear
Combo2.AddItem "A�RI MERKEZ"
Combo2.AddItem "D�YAD�N"
Combo2.AddItem "DO�UBAYAZIT"
Combo2.AddItem "ELE�K�RT"
Combo2.AddItem "HAMUR"
Combo2.AddItem "PATNOS"
Combo2.AddItem "TA�LI�AY"
Combo2.AddItem "TUTAK"
ElseIf Combo1.ListIndex = 4 Then
Combo2.Clear
Combo2.AddItem "A�A��REN"
Combo2.AddItem "AKSARAY MERKEZ"
Combo2.AddItem "ESK�L"
Combo2.AddItem "G�LA�A�"
Combo2.AddItem "G�ZELYURT"
Combo2.AddItem "ORTAK�Y"
Combo2.AddItem "SARIYAH��"
ElseIf Combo1.ListIndex = 5 Then
Combo2.Clear
Combo2.AddItem "AMASYA MERKEZ"
Combo2.AddItem "G�YN�CEK"
Combo2.AddItem "G�M��HACIK�Y"
Combo2.AddItem "HAMAM�Z�"
Combo2.AddItem "MERZ�FON"
Combo2.AddItem "SULUOVA"
Combo2.AddItem "TA�OVA"
ElseIf Combo1.ListIndex = 6 Then
Combo2.Clear
Combo2.AddItem "AKYURT"
Combo2.AddItem "ALTINDA�"
Combo2.AddItem "ANKARA MERKEZ"
Combo2.AddItem "AYA�"
Combo2.AddItem "BALA"
Combo2.AddItem "BEYPAZARI"
Combo2.AddItem "�AMLIDERE"
Combo2.AddItem "�ANKAYA"
Combo2.AddItem "�UBUK"
Combo2.AddItem "ELMADA�"
Combo2.AddItem "ET�MESGUT"
Combo2.AddItem "EVREN"
Combo2.AddItem "G�LBA�I/ANKARA"
Combo2.AddItem "G�D�L"
Combo2.AddItem "HAYMANA"
Combo2.AddItem "KALEC�K"
Combo2.AddItem "KAZAN"
Combo2.AddItem "KE���REN"
Combo2.AddItem "KIZILCAHAMAM"
Combo2.AddItem "MAMAK"
Combo2.AddItem "NALLIHAN"
Combo2.AddItem "POLATLI"
Combo2.AddItem "�EREFL�KO�H�SAR"
Combo2.AddItem "S�NCAN"
Combo2.AddItem "YEN�MAHALLE"
ElseIf Combo1.ListIndex = 7 Then
Combo2.Clear
Combo2.AddItem "AKSEK�"
Combo2.AddItem "ALANYA"
Combo2.AddItem "ANTALYA MERKEZ"
Combo2.AddItem "DEMRE"
Combo2.AddItem "ELMALI"
Combo2.AddItem "F�N�KE"
Combo2.AddItem "GAZ�PA�A"
Combo2.AddItem "G�NDO�MU�"
Combo2.AddItem "�BRADI"
Combo2.AddItem "KA�"
Combo2.AddItem "KEMER/ANTALYA"
Combo2.AddItem "KORKUTEL�"
Combo2.AddItem "KUMLUCA"
Combo2.AddItem "MANAVGAT"
Combo2.AddItem "SER�K"
ElseIf Combo1.ListIndex = 8 Then
Combo2.Clear
Combo2.AddItem "ARDAHAN MERKEZ"
Combo2.AddItem "�ILDIR"
Combo2.AddItem "DAMAL"
Combo2.AddItem "G�LE"
Combo2.AddItem "HANAK"
Combo2.AddItem "POSOF"
ElseIf Combo1.ListIndex = 9 Then
Combo2.Clear
Combo2.AddItem "ARDANU�"
Combo2.AddItem "ARHAV�"
Combo2.AddItem "ARTV�N MERKEZ"
Combo2.AddItem "BOR�KA"
Combo2.AddItem "HOPA"
Combo2.AddItem "MURGUL"
Combo2.AddItem "�AV�AT"
Combo2.AddItem "YUSUFEL�"
ElseIf Combo1.ListIndex = 10 Then
Combo2.Clear
Combo2.AddItem "AYDIN MERKEZ"
Combo2.AddItem "BOZDO�AN"
Combo2.AddItem "BUHARKENT"
Combo2.AddItem "��NE"
Combo2.AddItem "D�D�M(YEN�H�SAR)"
Combo2.AddItem "GERMENC�K"
Combo2.AddItem "�NC�RL�OVA"
Combo2.AddItem "KARACASU"
Combo2.AddItem "KARPUZLU"
Combo2.AddItem "KO�ARLI"
Combo2.AddItem "K��K"
Combo2.AddItem "KU�ADASI"
Combo2.AddItem "KUYUCAK"
Combo2.AddItem "NAZ�LL�"
Combo2.AddItem "S�KE"
Combo2.AddItem "SULTANH�SAR"
Combo2.AddItem "YEN�PAZAR/AYDIN"
ElseIf Combo1.ListIndex = 11 Then
Combo2.Clear
Combo2.AddItem "AYVALIK"
Combo2.AddItem "BALIKES�R MERKEZ"
Combo2.AddItem "BALYA"
Combo2.AddItem "BANDIRMA"
Combo2.AddItem "B�GAD��"
Combo2.AddItem "BURHAN�YE"
Combo2.AddItem "DURSUNBEY"
Combo2.AddItem "EDREM�T/BALIKES�R"
Combo2.AddItem "ERDEK"
Combo2.AddItem "G�ME�"
Combo2.AddItem "G�NEN/BALIKES�R"
Combo2.AddItem "HAVRAN"
Combo2.AddItem "�VR�ND�"
Combo2.AddItem "KEPSUT"
Combo2.AddItem "MANYAS"
Combo2.AddItem "MARMARA"
Combo2.AddItem "SAVA�TEPE"
Combo2.AddItem "SINDIRGI"
Combo2.AddItem "SUSURLUK"
ElseIf Combo1.ListIndex = 12 Then
Combo2.Clear
Combo2.AddItem "AMASRA"
Combo2.AddItem "BARTIN MERKEZ"
Combo2.AddItem "KURUCA��LE"
Combo2.AddItem "ULUS"
ElseIf Combo1.ListIndex = 13 Then
Combo2.Clear
Combo2.AddItem "BATMAN MERKEZ"
Combo2.AddItem "BE��R�"
Combo2.AddItem "GERC��"
Combo2.AddItem "HASANKEYF"
Combo2.AddItem "KOZLUK"
Combo2.AddItem "SASON"
ElseIf Combo1.ListIndex = 14 Then
Combo2.Clear
Combo2.AddItem "AYDINTEPE"
Combo2.AddItem "BAYBURT MERKEZ"
Combo2.AddItem "DEM�R�Z�"
ElseIf Combo1.ListIndex = 15 Then
Combo2.Clear
Combo2.AddItem "B�LEC�K MERKEZ"
Combo2.AddItem "BOZ�Y�K"
Combo2.AddItem "G�LPAZARI"
Combo2.AddItem "�NH�SAR"
Combo2.AddItem "OSMANEL�"
Combo2.AddItem "PAZARYER�"
Combo2.AddItem "S���T"
Combo2.AddItem "YEN�PAZAR/B�LEC�K"
ElseIf Combo1.ListIndex = 16 Then
Combo2.Clear
Combo2.AddItem "ADAKLI"
Combo2.AddItem "B�NG�L MERKEZ"
Combo2.AddItem "GEN�"
Combo2.AddItem "KARLIOVA"
Combo2.AddItem "K��I"
Combo2.AddItem "SOLHAN"
Combo2.AddItem "YAYLADERE"
Combo2.AddItem "YED�SU"
ElseIf Combo1.ListIndex = 17 Then
Combo2.Clear
Combo2.AddItem "AD�LCEVAZ"
Combo2.AddItem "AHLAT"
Combo2.AddItem "B�TL�S MERKEZ"
Combo2.AddItem "G�ROYMAK"
Combo2.AddItem "H�ZAN"
Combo2.AddItem "MUTK�"
Combo2.AddItem "TATVAN"
ElseIf Combo1.ListIndex = 18 Then
Combo2.Clear
Combo2.AddItem "BOLU MERKEZ"
Combo2.AddItem "D�RTD�VAN"
Combo2.AddItem "GEREDE"
Combo2.AddItem "G�YN�K"
Combo2.AddItem "KIBRISCIK"
Combo2.AddItem "MENGEN"
Combo2.AddItem "MUDURNU"
Combo2.AddItem "SEBEN"
Combo2.AddItem "YEN��A�A"
ElseIf Combo1.ListIndex = 19 Then
Combo2.Clear
Combo2.AddItem "A�LASUN"
Combo2.AddItem "ALTINYAYLA/BURDUR"
Combo2.AddItem "BUCAK"
Combo2.AddItem "BURDUR MERKEZ"
Combo2.AddItem "�AVDIR"
Combo2.AddItem "�ELT�K��"
Combo2.AddItem "G�LH�SAR"
Combo2.AddItem "KARAMANLI"
Combo2.AddItem "KEMER/BURDUR"
Combo2.AddItem "TEFENN�"
Combo2.AddItem "YE��LOVA"
ElseIf Combo1.ListIndex = 20 Then
Combo2.Clear
Combo2.AddItem "BURSA MERKEZ"
Combo2.AddItem "B�Y�KORHAN"
Combo2.AddItem "GEML�K"
Combo2.AddItem "G�RSU"
Combo2.AddItem "HARMANCIK"
Combo2.AddItem "�NEG�L"
Combo2.AddItem "�ZN�K"
Combo2.AddItem "KARACABEY"
Combo2.AddItem "KELES"
Combo2.AddItem "KESTEL"
Combo2.AddItem "MUDANYA"
Combo2.AddItem "MUSTAFA KEMAL PA�A"
Combo2.AddItem "N�L�FER"
Combo2.AddItem "ORHANEL�"
Combo2.AddItem "ORHANGAZ�"
Combo2.AddItem "OSMANGAZ�"
Combo2.AddItem "YEN��EH�R"
Combo2.AddItem "YILDIRIM"
ElseIf Combo1.ListIndex = 21 Then
Combo2.Clear
Combo2.AddItem "AYVACIK/�ANAKKALE"
Combo2.AddItem "BAYRAM��"
Combo2.AddItem "B�GA"
Combo2.AddItem "BOZCAADA"
Combo2.AddItem "�AN"
Combo2.AddItem "�ANAKKALE MERKEZ"
Combo2.AddItem "ECEABAT"
Combo2.AddItem "EZ�NE"
Combo2.AddItem "GEL�BOLU"
Combo2.AddItem "G�K�EADA"
Combo2.AddItem "LAPSEK�"
Combo2.AddItem "YEN�CE/�ANAKKALE"
ElseIf Combo1.ListIndex = 22 Then
Combo2.Clear
Combo2.AddItem "ATKARACALAR"
Combo2.AddItem "BAYRAM�REN"
Combo2.AddItem "�ANKIRI MERKEZ"
Combo2.AddItem "�ERKES"
Combo2.AddItem "ELD�VAN"
Combo2.AddItem "ILGAZ"
Combo2.AddItem "KIZILIRMAK"
Combo2.AddItem "KORGUN"
Combo2.AddItem "KUR�UNLU"
Combo2.AddItem "ORTA"
Combo2.AddItem "�ABAN�Z�"
Combo2.AddItem "YAPRAKLI"
ElseIf Combo1.ListIndex = 23 Then
Combo2.Clear
Combo2.AddItem "ALACA"
Combo2.AddItem "BAYAT/�ORUM"
Combo2.AddItem "BO�AZKALE"
Combo2.AddItem "�ORUM MERKEZ"
Combo2.AddItem "DODURGA"
Combo2.AddItem "�SK�L�P"
Combo2.AddItem "KARGI"
Combo2.AddItem "LA��N"
Combo2.AddItem "MEC�T�Z�"
Combo2.AddItem "O�UZLAR"
Combo2.AddItem "ORTAK�Y/�ORUM"
Combo2.AddItem "OSMANCIK"
Combo2.AddItem "SUNGURLU"
Combo2.AddItem "U�URLUDA�"
ElseIf Combo1.ListIndex = 24 Then
Combo2.Clear
Combo2.AddItem "ACIPAYAM"
Combo2.AddItem "AKK�Y"
Combo2.AddItem "BABADA�"
Combo2.AddItem "BAKLAN"
Combo2.AddItem "BEK�LL�"
Combo2.AddItem "BEYA�A�"
Combo2.AddItem "BOZKURT/DEN�ZL�"
Combo2.AddItem "BULDAN"
Combo2.AddItem "�AL"
Combo2.AddItem "�AMEL�"
Combo2.AddItem "�ARDAK"
Combo2.AddItem "��VR�L"
Combo2.AddItem "DEN�ZL� MERKEZ"
Combo2.AddItem "G�NEY"
Combo2.AddItem "HONAZ"
Combo2.AddItem "KALE/DEN�ZL�"
Combo2.AddItem "SARAYK�Y"
Combo2.AddItem "SER�NH�SAR"
Combo2.AddItem "TAVAS"
ElseIf Combo1.ListIndex = 25 Then
Combo2.Clear
Combo2.AddItem "B�SM�L"
Combo2.AddItem "�ERM�K"
Combo2.AddItem "�INAR"
Combo2.AddItem "��NG��"
Combo2.AddItem "D�CLE"
Combo2.AddItem "D�YARBAKIR MERKEZ"
Combo2.AddItem "E��L"
Combo2.AddItem "ERGAN�"
Combo2.AddItem "HAN�"
Combo2.AddItem "HAZRO"
Combo2.AddItem "KOCAK�Y"
Combo2.AddItem "KULP"
Combo2.AddItem "L�CE"
Combo2.AddItem "S�LVAN"
ElseIf Combo1.ListIndex = 26 Then
Combo2.Clear
Combo2.AddItem "AK�AKOCA"
Combo2.AddItem "��L�ML�"
Combo2.AddItem "CUMAYER�"
Combo2.AddItem "D�ZCE MERKEZ"
Combo2.AddItem "G�LYAKA"
Combo2.AddItem "G�M��OVA"
Combo2.AddItem "KAYNA�LI"
Combo2.AddItem "YI�ILCA"
ElseIf Combo1.ListIndex = 27 Then
Combo2.Clear
Combo2.AddItem "ED�RNE MERKEZ"
Combo2.AddItem "ENEZ"
Combo2.AddItem "HAVSA"
Combo2.AddItem "�PSALA"
Combo2.AddItem "KE�AN"
Combo2.AddItem "LALAPA�A"
Combo2.AddItem "MER��"
Combo2.AddItem "S�LO�LU"
Combo2.AddItem "UZUNK�PR�"
ElseIf Combo1.ListIndex = 28 Then
Combo2.Clear
Combo2.AddItem "A�IN"
Combo2.AddItem "ALACAKAYA"
Combo2.AddItem "ARICAK"
Combo2.AddItem "BASK�L"
Combo2.AddItem "ELAZI� MERKEZ"
Combo2.AddItem "KARAKO�AN"
Combo2.AddItem "KEBAN"
Combo2.AddItem "KOVANCILAR"
Combo2.AddItem "MADEN"
Combo2.AddItem "PALU"
Combo2.AddItem "S�VR�CE"
ElseIf Combo1.ListIndex = 29 Then
Combo2.Clear
Combo2.AddItem "�AYIRLI"
Combo2.AddItem "ERZ�NCAN MERKEZ"
Combo2.AddItem "�L��"
Combo2.AddItem "KEMAH"
Combo2.AddItem "KEMAL�YE"
Combo2.AddItem "OTLUKBEL�"
Combo2.AddItem "REFAH�YE"
Combo2.AddItem "TERCAN"
Combo2.AddItem "�Z�ML�"
ElseIf Combo1.ListIndex = 30 Then
Combo2.Clear
Combo2.AddItem "A�KALE"
Combo2.AddItem "�AT"
Combo2.AddItem "ERZURUM MERKEZ"
Combo2.AddItem "HINIS"
Combo2.AddItem "HORASAN"
Combo2.AddItem "ILICA"
Combo2.AddItem "�SP�R"
Combo2.AddItem "KARA�OBAN"
Combo2.AddItem "KARAYAZI"
Combo2.AddItem "K�PR�K�Y"
Combo2.AddItem "NARMAN"
Combo2.AddItem "OLTU"
Combo2.AddItem "OLUR"
Combo2.AddItem "PAS�NLER"
Combo2.AddItem "PAZARYOLU"
Combo2.AddItem "�ENKAYA"
Combo2.AddItem "TEKMAN"
Combo2.AddItem "TORTUM"
Combo2.AddItem "UZUNDERE"
ElseIf Combo1.ListIndex = 31 Then
Combo2.Clear
Combo2.AddItem "ALPU"
Combo2.AddItem "BEYL�KOVA"
Combo2.AddItem "��FTELER"
Combo2.AddItem "ESK��EH�R MERKEZ"
Combo2.AddItem "G�NY�Z�"
Combo2.AddItem "HAN"
Combo2.AddItem "�N�N�"
Combo2.AddItem "MAHMUD�YE"
Combo2.AddItem "M�HALGAZ�"
Combo2.AddItem "M�HALI��IK"
Combo2.AddItem "SARICAKAYA"
Combo2.AddItem "SEY�TGAZ�"
Combo2.AddItem "S�VR�H�SAR"
ElseIf Combo1.ListIndex = 32 Then
Combo2.Clear
Combo2.AddItem "ARABAN"
Combo2.AddItem "GAZ�ANTEP MERKEZ"
Combo2.AddItem "�SLAH�YE"
Combo2.AddItem "KARKAMI�"
Combo2.AddItem "N�Z�P"
Combo2.AddItem "NURDA�I"
Combo2.AddItem "O�UZEL�"
Combo2.AddItem "�AH�NBEY"
Combo2.AddItem "�EH�TKAM�L"
Combo2.AddItem "YAVUZEL�"
ElseIf Combo1.ListIndex = 33 Then
Combo2.Clear
Combo2.AddItem "ALUCRA"
Combo2.AddItem "BULANCAK"
Combo2.AddItem "�AMOLUK"
Combo2.AddItem "�ANAK�I"
Combo2.AddItem "DEREL�"
Combo2.AddItem "DO�ANKENT"
Combo2.AddItem "ESP�YE"
Combo2.AddItem "EYNES�L"
Combo2.AddItem "G�RESUN MERKEZ"
Combo2.AddItem "G�RELE"
Combo2.AddItem "G�CE"
Combo2.AddItem "KE�AP"
Combo2.AddItem "P�RAZ�Z"
Combo2.AddItem "�EB�NKARAH�SAR"
Combo2.AddItem "T�REBOLU"
Combo2.AddItem "YA�LIDERE"
ElseIf Combo1.ListIndex = 34 Then
Combo2.Clear
Combo2.AddItem "G�M��HANE MERKEZ"
Combo2.AddItem "KELK�T"
Combo2.AddItem "K�SE"
Combo2.AddItem "K�RT�N"
Combo2.AddItem "��RAN"
Combo2.AddItem "TORUL"
ElseIf Combo1.ListIndex = 35 Then
Combo2.Clear
Combo2.AddItem "�UKURCA"
Combo2.AddItem "HAKKAR� MERKEZ"
Combo2.AddItem "�EMD�NL�"
Combo2.AddItem "Y�KSEKOVA"
ElseIf Combo1.ListIndex = 36 Then
Combo2.Clear
Combo2.AddItem "ALTIN�Z�"
Combo2.AddItem "BELEN"
Combo2.AddItem "D�RTYOL"
Combo2.AddItem "ERZ�N"
Combo2.AddItem "HASSA"
Combo2.AddItem "HATAY MERKEZ"
Combo2.AddItem "�SKENDERUN"
Combo2.AddItem "KIRIKHAN"
Combo2.AddItem "KUMLU"
Combo2.AddItem "REYHANLI"
Combo2.AddItem "SAMANDA�"
Combo2.AddItem "YAYLADA�"
ElseIf Combo1.ListIndex = 37 Then
Combo2.Clear
Combo2.AddItem "ARALIK"
Combo2.AddItem "I�DIR MERKEZ"
Combo2.AddItem "KARAKOYUNLU"
Combo2.AddItem "TUZLUCA"
ElseIf Combo1.ListIndex = 38 Then
Combo2.Clear
Combo2.AddItem "AKSU"
Combo2.AddItem "ATABEY"
Combo2.AddItem "E��RD�R"
Combo2.AddItem "GELENDOST"
Combo2.AddItem "G�NEN/ISPARTA"
Combo2.AddItem "ISPARTA MERKEZ"
Combo2.AddItem "KE��BORLU"
Combo2.AddItem "�ARK�KARAA�A�"
Combo2.AddItem "SEN�RKENT"
Combo2.AddItem "S�T��LER"
Combo2.AddItem "ULUBORLU"
Combo2.AddItem "YALVA�"
Combo2.AddItem "YEN��ARBADEML�"
ElseIf Combo1.ListIndex = 39 Then
Combo2.Clear
Combo2.AddItem "ADALAR"
Combo2.AddItem "AVCILAR"
Combo2.AddItem "BA�CILAR"
Combo2.AddItem "BAH�EL�EVLER"
Combo2.AddItem "BAKIRK�Y"
Combo2.AddItem "BAYRAMPA�A"
Combo2.AddItem "BE��KTA�"
Combo2.AddItem "BEYKOZ"
Combo2.AddItem "BEYO�LU"
Combo2.AddItem "B�Y�K�EKMECE"
Combo2.AddItem "�ATALCA"
Combo2.AddItem "EM�N�N�"
Combo2.AddItem "ESENLER"
Combo2.AddItem "EY�P"
Combo2.AddItem "FAT�H"
Combo2.AddItem "GAZ�OSMANPA�A"
Combo2.AddItem "G�NG�REN"
Combo2.AddItem "�STANBUL MERKEZ"
Combo2.AddItem "KADIK�Y"
Combo2.AddItem "KA�ITHANE"
Combo2.AddItem "KARTAL"
Combo2.AddItem "K���K�EKMECE"
Combo2.AddItem "MALTEPE"
Combo2.AddItem "PEND�K"
Combo2.AddItem "SARIYER"
Combo2.AddItem "S�L�VR�"
Combo2.AddItem "SULTANBEYL�"
Combo2.AddItem "��LE"
Combo2.AddItem "���L�"
Combo2.AddItem "TUZLA"
Combo2.AddItem "�MRAN�YE"
Combo2.AddItem "�SK�DAR"
Combo2.AddItem "ZEYT�NBURNU"
ElseIf Combo1.ListIndex = 40 Then
Combo2.Clear
Combo2.AddItem "AL�A�A"
Combo2.AddItem "BAL�OVA"
Combo2.AddItem "BAYINDIR"
Combo2.AddItem "BERGAMA"
Combo2.AddItem "BEYDA�"
Combo2.AddItem "BORNOVA"
Combo2.AddItem "BUCA"
Combo2.AddItem "�E�ME"
Combo2.AddItem "���L�"
Combo2.AddItem "D�K�L�"
Combo2.AddItem "FO�A"
Combo2.AddItem "GAZ�EM�R"
Combo2.AddItem "G�ZELBAH�E"
Combo2.AddItem "�ZM�R MERKEZ"
Combo2.AddItem "KARABURUN"
Combo2.AddItem "KAR�IYAKA"
Combo2.AddItem "KEMALPA�A"
Combo2.AddItem "KINIK"
Combo2.AddItem "K�RAZ"
Combo2.AddItem "KONAK"
Combo2.AddItem "MENDERES"
Combo2.AddItem "MENEMEN"
Combo2.AddItem "NARLIDERE"
Combo2.AddItem "�DEM��"
Combo2.AddItem "SEFER�H�SAR"
Combo2.AddItem "SEL�UK"
Combo2.AddItem "T�RE"
Combo2.AddItem "TORBALI"
Combo2.AddItem "URLA"
ElseIf Combo1.ListIndex = 41 Then
Combo2.Clear
Combo2.AddItem "AF��N"
Combo2.AddItem "ANDIRIN"
Combo2.AddItem "�A�LIYANCER�T"
Combo2.AddItem "EK�N�Z�"
Combo2.AddItem "ELB�STAN"
Combo2.AddItem "G�KSUN"
Combo2.AddItem "KAHRAMANMARA� MERKEZ"
Combo2.AddItem "NURHAK"
Combo2.AddItem "PAZARCIK"
Combo2.AddItem "T�RKO�LU"
ElseIf Combo1.ListIndex = 42 Then
Combo2.Clear
Combo2.AddItem "EFLAN�"
Combo2.AddItem "ESK�PAZAR"
Combo2.AddItem "KARAB�K MERKEZ"
Combo2.AddItem "OVACIK/KARAB�K"
Combo2.AddItem "SAFRANBOLU"
Combo2.AddItem "YEN�CE/KARAB�K"
ElseIf Combo1.ListIndex = 43 Then
Combo2.Clear
Combo2.AddItem "AYRANCI"
Combo2.AddItem "BA�YAYLA"
Combo2.AddItem "ERMENEK"
Combo2.AddItem "KARAMAN MERKEZ"
Combo2.AddItem "KAZIMKARABEK�R"
Combo2.AddItem "SARIVEL�LER"
ElseIf Combo1.ListIndex = 44 Then
Combo2.Clear
Combo2.AddItem "AKYAKA"
Combo2.AddItem "ARPA�AY"
Combo2.AddItem "D�GOR"
Combo2.AddItem "KA�IZMAN"
Combo2.AddItem "KARS MERKEZ"
Combo2.AddItem "SARIKAMI�"
Combo2.AddItem "SEL�M"
Combo2.AddItem "SUSUZ"
ElseIf Combo1.ListIndex = 45 Then
Combo2.Clear
Combo2.AddItem "ABANA"
Combo2.AddItem "A�LI"
Combo2.AddItem "ARA�"
Combo2.AddItem "AZDAVAY"
Combo2.AddItem "BOZKURT/KASTAMONU"
Combo2.AddItem "�ATALZEYT�N"
Combo2.AddItem "C�DE"
Combo2.AddItem "DADAY"
Combo2.AddItem "DEVREKAN�"
Combo2.AddItem "DO�ANYURT"
Combo2.AddItem "HAN�N�"
Combo2.AddItem "�HSANGAZ�"
Combo2.AddItem "�NEBOLU"
Combo2.AddItem "KASTAMONU MERKEZ"
Combo2.AddItem "K�RE"
Combo2.AddItem "PINARBA�I/KASTAMONU"
Combo2.AddItem "SEYD�LER"
Combo2.AddItem "�ENPAZAR"
Combo2.AddItem "TA�K�PR�"
Combo2.AddItem "TOSYA"
ElseIf Combo1.ListIndex = 46 Then
Combo2.Clear
Combo2.AddItem "AKKI�LA"
Combo2.AddItem "B�NYAN"
Combo2.AddItem "DEVEL�"
Combo2.AddItem "FELAH�YE"
Combo2.AddItem "HACILAR"
Combo2.AddItem "�NCESU"
Combo2.AddItem "KAYSER� MERKEZ"
Combo2.AddItem "KOCAS�NAN"
Combo2.AddItem "MEL�KGAZ�"
Combo2.AddItem "�ZVATAN"
Combo2.AddItem "PINARBA�I/KAYSER�"
Combo2.AddItem "SARIO�LAN"
Combo2.AddItem "SARIZ"
Combo2.AddItem "TALAS"
Combo2.AddItem "TOMARZA"
Combo2.AddItem "YAHYALI"
Combo2.AddItem "YE��LH�SAR"
ElseIf Combo1.ListIndex = 47 Then
Combo2.Clear
Combo2.AddItem "BAH��L�"
Combo2.AddItem "BALI�EYH"
Combo2.AddItem "�ELEB�"
Combo2.AddItem "DEL�CE"
Combo2.AddItem "KARAKE��L�"
Combo2.AddItem "KESK�N"
Combo2.AddItem "KIRIKKALE MERKEZ"
Combo2.AddItem "SULAKYURT"
Combo2.AddItem "YAH��HAN"
ElseIf Combo1.ListIndex = 48 Then
Combo2.Clear
Combo2.AddItem "BABAESK�"
Combo2.AddItem "DEM�RK�Y"
Combo2.AddItem "KIRKLAREL� MERKEZ"
Combo2.AddItem "KOF�AZ"
Combo2.AddItem "L�LEBURGAZ"
Combo2.AddItem "PEHL�VANK�Y"
Combo2.AddItem "PINARH�SAR"
Combo2.AddItem "V�ZE"
ElseIf Combo1.ListIndex = 49 Then
Combo2.Clear
Combo2.AddItem "AK�AKENT"
Combo2.AddItem "AKPINAR"
Combo2.AddItem "BOZTEPE"
Combo2.AddItem "���EKDA�I"
Combo2.AddItem "KAMAN"
Combo2.AddItem "KIR�EH�R MERKEZ"
Combo2.AddItem "MUCUR"
ElseIf Combo1.ListIndex = 50 Then
Combo2.Clear
Combo2.AddItem "ELBEYL�"
Combo2.AddItem "K�L�S MERKEZ"
Combo2.AddItem "MUSABEYL�"
Combo2.AddItem "POLATEL�"
ElseIf Combo1.ListIndex = 51 Then
Combo2.Clear
Combo2.AddItem "DER�NCE"
Combo2.AddItem "GEBZE"
Combo2.AddItem "G�LC�K"
Combo2.AddItem "KANDIRA"
Combo2.AddItem "KARAM�RSEL"
Combo2.AddItem "KOCAEL� MERKEZ"
Combo2.AddItem "K�RFEZ"
ElseIf Combo1.ListIndex = 52 Then
Combo2.Clear
Combo2.AddItem "AHIRLI"
Combo2.AddItem "AK�REN"
Combo2.AddItem "AK�EH�R"
Combo2.AddItem "ALTINEK�N"
Combo2.AddItem "BEY�EH�R"
Combo2.AddItem "BOZKIR"
Combo2.AddItem "�ELT�K"
Combo2.AddItem "C�HANBEYL�"
Combo2.AddItem "�UMRA"
Combo2.AddItem "DERBENT"
Combo2.AddItem "DEREBUCAK"
Combo2.AddItem "DO�ANH�SAR"
Combo2.AddItem "EM�RGAZ�"
Combo2.AddItem "ERE�L�/KONYA"
Combo2.AddItem "G�NEYSINIR"
Combo2.AddItem "HAD�M"
Combo2.AddItem "HALKAPINAR"
Combo2.AddItem "H�Y�K"
Combo2.AddItem "ILGIN"
Combo2.AddItem "KADINHANI"
Combo2.AddItem "KARAPINAR"
Combo2.AddItem "KARATAY"
Combo2.AddItem "KONYA MERKEZ"
Combo2.AddItem "KULU"
Combo2.AddItem "MERAM"
Combo2.AddItem "SARAY�N�"
Combo2.AddItem "SEL�UKLU"
Combo2.AddItem "SEYD��EH�R"
Combo2.AddItem "TA�KENT"
Combo2.AddItem "TUZLUK�U"
Combo2.AddItem "YALIH�Y�K"
Combo2.AddItem "YUNAK"
ElseIf Combo1.ListIndex = 53 Then
Combo2.Clear
Combo2.AddItem "ALTINTA�"
Combo2.AddItem "ASLANAPA"
Combo2.AddItem "�AVDARH�SAR"
Combo2.AddItem "DOMAN��"
Combo2.AddItem "DUMLUPINAR"
Combo2.AddItem "EMET"
Combo2.AddItem "GED�Z"
Combo2.AddItem "H�SARCIK"
Combo2.AddItem "K�TAHYA MERKEZ"
Combo2.AddItem "PAZARLAR"
Combo2.AddItem "S�MAV"
Combo2.AddItem "�APHANE"
Combo2.AddItem "TAV�ANLI"
ElseIf Combo1.ListIndex = 54 Then
Combo2.Clear
Combo2.AddItem "AK�ADA�"
Combo2.AddItem "ARAPG�R"
Combo2.AddItem "ARGUVAN"
Combo2.AddItem "BATTALGAZ�"
Combo2.AddItem "DARENDE"
Combo2.AddItem "DO�AN�EH�R"
Combo2.AddItem "DO�ANYOL"
Combo2.AddItem "HEK�MHAN"
Combo2.AddItem "KALE/MALATYA"
Combo2.AddItem "KULUNCAK"
Combo2.AddItem "MALATYA MERKEZ"
Combo2.AddItem "P�T�RGE"
Combo2.AddItem "YAZIHAN"
Combo2.AddItem "YE��LYURT/MALATYA"
ElseIf Combo1.ListIndex = 55 Then
Combo2.Clear
Combo2.AddItem "AHMETL�"
Combo2.AddItem "AKH�SAR"
Combo2.AddItem "ALA�EH�R"
Combo2.AddItem "DEM�RC�"
Combo2.AddItem "G�LMARMARA"
Combo2.AddItem "G�RDES"
Combo2.AddItem "KIRKA�A�"
Combo2.AddItem "K�PR�BA�I/MAN�SA"
Combo2.AddItem "KULA"
Combo2.AddItem "MAN�SA MERKEZ"
Combo2.AddItem "SAL�HL�"
Combo2.AddItem "SARIG�L"
Combo2.AddItem "SARUHANLI"
Combo2.AddItem "SELEND�"
Combo2.AddItem "SOMA"
Combo2.AddItem "TURGUTLU"
ElseIf Combo1.ListIndex = 56 Then
Combo2.Clear
Combo2.AddItem "DARGE��T"
Combo2.AddItem "DER�K"
Combo2.AddItem "KIZILTEPE"
Combo2.AddItem "MARD�N MERKEZ"
Combo2.AddItem "MAZIDA�I"
Combo2.AddItem "M�DYAT"
Combo2.AddItem "NUSAYB�N"
Combo2.AddItem "�MERL�"
Combo2.AddItem "SAVUR"
Combo2.AddItem "YE��LL�"
ElseIf Combo1.ListIndex = 57 Then
Combo2.Clear
Combo2.AddItem "ANAMUR"
Combo2.AddItem "AYDINCIK/MERS�N"
Combo2.AddItem "BOZYAZI"
Combo2.AddItem "�AMLIYAYLA"
Combo2.AddItem "ERDEML�"
Combo2.AddItem "G�LNAR"
Combo2.AddItem "MERS�N MERKEZ"
Combo2.AddItem "MUT"
Combo2.AddItem "S�L�FKE"
Combo2.AddItem "TARSUS"
ElseIf Combo1.ListIndex = 58 Then
Combo2.Clear
Combo2.AddItem "BODRUM"
Combo2.AddItem "DALAMAN"
Combo2.AddItem "DAT�A"
Combo2.AddItem "FETH�YE"
Combo2.AddItem "KAVAKLIDERE"
Combo2.AddItem "K�YCE��Z"
Combo2.AddItem "MARMAR�S"
Combo2.AddItem "M�LAS"
Combo2.AddItem "MU�LA MERKEZ"
Combo2.AddItem "ORTACA"
Combo2.AddItem "ULA"
Combo2.AddItem "YATA�AN"
ElseIf Combo1.ListIndex = 59 Then
Combo2.Clear
Combo2.AddItem "BULANIK"
Combo2.AddItem "HASK�Y"
Combo2.AddItem "KORKUT"
Combo2.AddItem "MALAZG�RT"
Combo2.AddItem "MU� MERKEZ"
Combo2.AddItem "VARTO"
ElseIf Combo1.ListIndex = 60 Then
Combo2.Clear
Combo2.AddItem "ACIG�L"
Combo2.AddItem "AVONOS"
Combo2.AddItem "DER�NKUYU"
Combo2.AddItem "G�L�EH�R"
Combo2.AddItem "HACIBEKTA�"
Combo2.AddItem "KOZAKLI"
Combo2.AddItem "NEV�EH�R MERKEZ"
Combo2.AddItem "�RG�P"
ElseIf Combo1.ListIndex = 61 Then
Combo2.Clear
Combo2.AddItem "ALTUNH�SAR"
Combo2.AddItem "BOR"
Combo2.AddItem "�AMARDI"
Combo2.AddItem "��FTL�K"
Combo2.AddItem "N��DE MERKEZ"
Combo2.AddItem "ULUKI�LA"
ElseIf Combo1.ListIndex = 62 Then
Combo2.Clear
Combo2.AddItem "AKKU�"
Combo2.AddItem "AYBASTI"
Combo2.AddItem "�AMA�"
Combo2.AddItem "�ATALPINAR"
Combo2.AddItem "�AYBA�I"
Combo2.AddItem "FATSA"
Combo2.AddItem "G�LK�Y"
Combo2.AddItem "G�LYALI"
Combo2.AddItem "G�RGENTEPE"
Combo2.AddItem "�K�ZCE"
Combo2.AddItem "KABAD�Z"
Combo2.AddItem "KABATA�"
Combo2.AddItem "KORGAN"
Combo2.AddItem "KUMRU"
Combo2.AddItem "MESUD�YE"
Combo2.AddItem "ORDU MERKEZ"
Combo2.AddItem "PER�EMBE"
Combo2.AddItem "ULUBEY/ORDU"
Combo2.AddItem "�NYE"
ElseIf Combo1.ListIndex = 63 Then
Combo2.Clear
Combo2.AddItem "BAH�E"
Combo2.AddItem "D�Z���"
Combo2.AddItem "HASANBEYL�"
Combo2.AddItem "KAD�RL�"
Combo2.AddItem "OSMAN�YE MERKEZ"
Combo2.AddItem "SUMBAS"
Combo2.AddItem "TOPRAKKALE"
ElseIf Combo1.ListIndex = 64 Then
Combo2.Clear
Combo2.AddItem "ARDE�EN"
Combo2.AddItem "�AMLIHEM��N"
Combo2.AddItem "�AYEL�"
Combo2.AddItem "DEREPAZARI"
Combo2.AddItem "FINDIKLI"
Combo2.AddItem "G�NEYSU"
Combo2.AddItem "HEM��N"
Combo2.AddItem "�K�ZDERE"
Combo2.AddItem "�Y�DERE"
Combo2.AddItem "KALKANDERE"
Combo2.AddItem "PAZAR/R�ZE"
Combo2.AddItem "R�ZE MERKEZ"
ElseIf Combo1.ListIndex = 65 Then
Combo2.Clear
Combo2.AddItem "AKYAZI"
Combo2.AddItem "FER�ZL�"
Combo2.AddItem "GEYVE"
Combo2.AddItem "HENDEK"
Combo2.AddItem "KARAP�R�EK"
Combo2.AddItem "KARASU"
Combo2.AddItem "KAYNARCA"
Combo2.AddItem "KOCAAL�"
Combo2.AddItem "PAMUKOVA"
Combo2.AddItem "SAKARYA MERKEZ"
Combo2.AddItem "SAPANCA"
Combo2.AddItem "S���TL�"
Combo2.AddItem "TARAKLI"
ElseIf Combo1.ListIndex = 66 Then
Combo2.Clear
Combo2.AddItem "ALA�AM"
Combo2.AddItem "ASARCIK"
Combo2.AddItem "AYVACIK/SAMSUN"
Combo2.AddItem "BAFRA"
Combo2.AddItem "�AR�AMBA"
Combo2.AddItem "HAVZA"
Combo2.AddItem "KAVAK"
Combo2.AddItem "LAD�K"
Combo2.AddItem "ONDOKUZMAYIS"
Combo2.AddItem "SALIPAZARI"
Combo2.AddItem "SAMSUN MERKEZ"
Combo2.AddItem "TEKKEK�Y"
Combo2.AddItem "TERME"
Combo2.AddItem "VEZ�RK�PR�"
Combo2.AddItem "YAKAKENT"
ElseIf Combo1.ListIndex = 67 Then
Combo2.Clear
Combo2.AddItem "AYDINLAR"
Combo2.AddItem "BAYKAN"
Combo2.AddItem "ERUH"
Combo2.AddItem "KURTALAN"
Combo2.AddItem "PERVAR�"
Combo2.AddItem "S��RT MERKEZ"
Combo2.AddItem "��RVAN"
ElseIf Combo1.ListIndex = 68 Then
Combo2.Clear
Combo2.AddItem "AYANCIK"
Combo2.AddItem "BOYABAT"
Combo2.AddItem "D�KMEN"
Combo2.AddItem "DURA�AN"
Combo2.AddItem "ERFELEK"
Combo2.AddItem "GERZE"
Combo2.AddItem "SARAYD�Z�"
Combo2.AddItem "S�NOP MERKEZ"
Combo2.AddItem "T�RKEL�"
ElseIf Combo1.ListIndex = 69 Then
Combo2.Clear
Combo2.AddItem "AKINCILAR"
Combo2.AddItem "ALTINYAYLA/S�VAS"
Combo2.AddItem "D�VR���"
Combo2.AddItem "DO�AN�AR"
Combo2.AddItem "GEMEREK"
Combo2.AddItem "G�LOVA"
Combo2.AddItem "G�R�N"
Combo2.AddItem "HAF�K"
Combo2.AddItem "�MRANLI"
Combo2.AddItem "KANGAL"
Combo2.AddItem "KOYULH�SAR"
Combo2.AddItem "S�VAS MERKEZ"
Combo2.AddItem "SU�EHR�"
Combo2.AddItem "�ARKI�LA"
Combo2.AddItem "ULA�"
Combo2.AddItem "YILDIZEL�"
Combo2.AddItem "ZARA"
ElseIf Combo1.ListIndex = 70 Then
Combo2.Clear
Combo2.AddItem "AK�AKALE"
Combo2.AddItem "B�REC�K"
Combo2.AddItem "BOZOVA"
Combo2.AddItem "CEYLANPINAR"
Combo2.AddItem "HALFET�"
Combo2.AddItem "HARRAN"
Combo2.AddItem "H�LVAN"
Combo2.AddItem "S�VEREK"
Combo2.AddItem "SURU�"
Combo2.AddItem "�ANLIURFA MERKEZ"
Combo2.AddItem "V�RAN�EH�R"
ElseIf Combo1.ListIndex = 71 Then
Combo2.Clear
Combo2.AddItem "BEYT���EBAP"
Combo2.AddItem "C�ZRE"
Combo2.AddItem "G��L�KONAK"
Combo2.AddItem "�D�L"
Combo2.AddItem "S�LOP�"
Combo2.AddItem "�IRNAK MERKEZ"
Combo2.AddItem "ULUDERE"

ElseIf Combo1.ListIndex = 72 Then
Combo2.Clear
Combo2.AddItem "�ERKEZK�Y"
Combo2.AddItem "�ORLU"
Combo2.AddItem "HAYRABOLU"
Combo2.AddItem "MALKARA"
Combo2.AddItem "MARMARAERE�L�S�"
Combo2.AddItem "MURATLI"
Combo2.AddItem "SARAY/TEK�RDA�"
Combo2.AddItem "�ARK�Y"
Combo2.AddItem "TEK�RDA� MERKEZ"
ElseIf Combo1.ListIndex = 73 Then
Combo2.Clear
Combo2.AddItem "ALMUS"
Combo2.AddItem "ARTOVA"
Combo2.AddItem "BA���FTL�K"
Combo2.AddItem "ERBAA"
Combo2.AddItem "N�KSAR"
Combo2.AddItem "PAZAR/TOKAT"
Combo2.AddItem "RE�AD�YE"
Combo2.AddItem "SULUSARAY"
Combo2.AddItem "TOKAT MERKEZ"
Combo2.AddItem "TURHAL"
Combo2.AddItem "YE��LYURT/TOKAT"
Combo2.AddItem "Z�LE"
ElseIf Combo1.ListIndex = 74 Then
Combo2.Clear
Combo2.AddItem "AK�AABAT"
Combo2.AddItem "ARAKLI"
Combo2.AddItem "ARS�N"
Combo2.AddItem "BE��KD�Z�"
Combo2.AddItem "�AR�IBA�I"
Combo2.AddItem "�AYKARA"
Combo2.AddItem "DERNEPAZARI"
Combo2.AddItem "D�ZK�Y"
Combo2.AddItem "HAYRAT"
Combo2.AddItem "K�PR�BA�I/TRABZON"
Combo2.AddItem "MA�KA"
Combo2.AddItem "OF"
Combo2.AddItem "�ALPAZARI"
Combo2.AddItem "S�RMENE"
Combo2.AddItem "TONYA"
Combo2.AddItem "TRABZON MERKEZ"
Combo2.AddItem "VAKFIKEB�R"
Combo2.AddItem "YOMRA"
ElseIf Combo1.ListIndex = 75 Then
Combo2.Clear
Combo2.AddItem "�EM��GEZEK"
Combo2.AddItem "HOZAT"
Combo2.AddItem "MAZG�RT"
Combo2.AddItem "NAZ�M�YE"
Combo2.AddItem "OVACIK/TUNCEL�"
Combo2.AddItem "PERTEK"
Combo2.AddItem "P�L�MB�R"
Combo2.AddItem "TUNCEL� MERKEZ"
ElseIf Combo1.ListIndex = 76 Then
Combo2.Clear
Combo2.AddItem "BANAZ"
Combo2.AddItem "E�ME"
Combo2.AddItem "KARAHALLI"
Combo2.AddItem "S�VASLI"
Combo2.AddItem "ULUBEY/U�AK"
Combo2.AddItem "U�AK MERKEZ"
ElseIf Combo1.ListIndex = 77 Then
Combo2.Clear
Combo2.AddItem "BAH�ESARAY"
Combo2.AddItem "BA�KALE"
Combo2.AddItem "�ALDIRAN"
Combo2.AddItem "�ATAK"
Combo2.AddItem "EDREM�T/VAN"
Combo2.AddItem "ERC��"
Combo2.AddItem "GEVA�"
Combo2.AddItem "G�RPINAR"
Combo2.AddItem "MURAD�YE"
Combo2.AddItem "�ZALP"
Combo2.AddItem "SARAY/VAN"
Combo2.AddItem "VAN MERKEZ"
ElseIf Combo1.ListIndex = 78 Then
Combo2.Clear
Combo2.AddItem "ALTINOVA"
Combo2.AddItem "ARMUTLU"
Combo2.AddItem "��FTL�KK�Y"
Combo2.AddItem "�INARCIK"
Combo2.AddItem "TERMAL"
Combo2.AddItem "YALOVA MERKEZ"
ElseIf Combo1.ListIndex = 79 Then
Combo2.Clear
Combo2.AddItem "AKDA�MADEN�"
Combo2.AddItem "AYDINCIK/YOZGAT"
Combo2.AddItem "BO�AZLIYAN"
Combo2.AddItem "�ANDIR"
Combo2.AddItem "�AYIRALAN"
Combo2.AddItem "�EKEREK"
Combo2.AddItem "KADI�EHR�"
Combo2.AddItem "SARAYKENT"
Combo2.AddItem "SARIKAYA"
Combo2.AddItem "�EFAATL�"
Combo2.AddItem "SORGUN"
Combo2.AddItem "YEN�FAKILI"
Combo2.AddItem "YERK�Y"
Combo2.AddItem "YOZGAT MERKEZ"
ElseIf Combo1.ListIndex = 80 Then
Combo2.AddItem "ALAPLI"
Combo2.AddItem "�AYCUMA"
Combo2.AddItem "DEVREK"
Combo2.AddItem "ERE�L�/ZONGULDAK"
Combo2.AddItem "G�K�EBEY"
Combo2.AddItem "ZONGULDAK MERKEZ"
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Len(Text14) = 0 Then
ElseIf Len(Text14) < 10 Then
MsgBox "TELEFON NUMARASI 10 HANEL� OLMALIDIR."
Text14.Text = ""
Text14.BackColor = vbYellow
End If

If Len(Text5) = 0 Then
ElseIf Len(Text5) < 11 Then
MsgBox "T.C K�ML�K NUMARASI NUMARASI 11 HANEL� OLMALIDIR."
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



Set db = OpenDatabase(App.Path & "\�ifre.mdb")
Set rs = db.OpenRecordset("tablo1")
c = rs.RecordCount
For sayac = 0 To c
If Text17 <> rs!OdaNo Then
rs.MoveNext
Else
MsgBox "ODA DOLUDUR.L�TFEN BA�KA ODA NUMARASI G�R�N�Z."
Text17 = ""
Text17.BackColor = vbYellow
a = a + 1
b = 1
End If
Next sayac
If a > 0 Then
MsgBox a & "tane bo� alan b�rak�lm��"
Label18.Visible = True
ElseIf a <= 0 And b <> 1 Then
MsgBox "bo� alan b�rakmad�n�z"
Label18.Visible = False
Set db = OpenDatabase(App.Path & "\�ifre.mdb")
Set rs = db.OpenRecordset("tablo1")
rs.AddNew
rs.Fields("Ad�") = Text1.Text
rs.Fields("soyad�") = Text2.Text
rs.Fields("Baba_ad�") = Text3.Text
rs.Fields("Anne_ad�") = Text4.Text
rs.Fields("Tc") = Text5.Text
rs.Fields("il") = Combo1.Text
rs.Fields("il�e") = Combo2.Text
rs.Fields("Mahalle_K�y") = Text8.Text
rs.Fields("�kametgah_Adresi") = Text9.Text
rs.Fields("Do�um_Yeri") = Text10.Text
rs.Fields("Do�um_Tarih") = Text11.Text
rs.Fields("Cinsiyet") = Combo3.Text
rs.Fields("Medeni_Hali") = Combo4.Text
rs.Fields("Telefon") = Text14.Text
rs.Fields("Mesle�i") = Text15.Text
rs.Fields("E_Posta") = Text16.Text
rs.Fields("OdaNo") = Text17.Text
rs.Fields("Geli�_Tarihi") = Label19.Caption
rs.Fields("Geli�_Saati") = Label20.Caption
rs.Fields("G�n") = Text6.Text
rs.Fields("Fiyat") = Label23.Caption
rs.Update
rs.Close
MsgBox "m��teri kay�d� yap�ld�."
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
Set db = OpenDatabase(App.Path & "\�ifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
If a >= 1 Then
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
MsgBox "TELEFON NUMARASI 10 HANEL� OLMALIDIR."
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
MsgBox "T.C K�ML�K NUMARASI NUMARASI 11 HANEL� OLMALIDIR."
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
