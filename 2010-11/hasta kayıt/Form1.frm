VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   11100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   11100
   ScaleWidth      =   15585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Sil"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   54
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000002&
      Caption         =   "Son Kayýt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   53
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000002&
      Caption         =   "Sonraki Kayýt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   52
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000002&
      Caption         =   "Önceki Kayýt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   51
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000002&
      Caption         =   "Ýlk Kayýt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   50
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000002&
      Caption         =   "Kayit Bul"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      Picture         =   "Form1.frx":172D0
      TabIndex        =   49
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000002&
      Caption         =   "Düzenle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      TabIndex        =   48
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000002&
      Caption         =   "Yeni Kayýt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   47
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000002&
      Caption         =   "Ekle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   46
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text22 
      DataField       =   "Soyad"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   3240
      TabIndex        =   45
      Text            =   "Text22"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text21 
      DataField       =   "Adres"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   11880
      TabIndex        =   44
      Text            =   "Text21"
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox Text20 
      DataField       =   "E-mail adresi"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   11880
      TabIndex        =   43
      Text            =   "Text20"
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text19 
      DataField       =   "telefon(cep)"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   11880
      TabIndex        =   42
      Text            =   "Text19"
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text18 
      DataField       =   "Telefon(sabit)"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   11880
      TabIndex        =   41
      Text            =   "Text18"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox Text17 
      DataField       =   "Kan Grubu"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   6720
      TabIndex        =   40
      Text            =   "Text17"
      Top             =   6480
      Width           =   2895
   End
   Begin VB.TextBox Text16 
      DataField       =   "Diyabet"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   6720
      TabIndex        =   39
      Text            =   "Text16"
      Top             =   5760
      Width           =   2895
   End
   Begin VB.TextBox Text15 
      DataField       =   "HIV"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   6720
      TabIndex        =   38
      Text            =   "Text15"
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox Text14 
      DataField       =   "Hepatit"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   6720
      TabIndex        =   37
      Text            =   "Text14"
      Top             =   4320
      Width           =   2895
   End
   Begin VB.TextBox Text13 
      DataField       =   "Alkol"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   6720
      TabIndex        =   36
      Text            =   "Text13"
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox Text12 
      DataField       =   "Sigara"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   6720
      TabIndex        =   35
      Text            =   "Text12"
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text11 
      DataField       =   "Kilo"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   6720
      TabIndex        =   34
      Text            =   "Text11"
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text10 
      DataField       =   "Boy"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   6720
      TabIndex        =   33
      Text            =   "Text10"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox Text9 
      DataField       =   "Baðlý Old Kurum"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   1800
      TabIndex        =   32
      Text            =   "Text9"
      Top             =   7200
      Width           =   2895
   End
   Begin VB.TextBox Text8 
      DataField       =   "Meslek"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   1800
      TabIndex        =   31
      Text            =   "Text8"
      Top             =   6480
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      DataField       =   "Baba Adý"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   1800
      TabIndex        =   30
      Text            =   "Text7"
      Top             =   5760
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      DataField       =   "Ana Adý"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   1800
      TabIndex        =   29
      Text            =   "Text6"
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      DataField       =   "Doðum Tarihi"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   1800
      TabIndex        =   28
      Text            =   "Text5"
      Top             =   4320
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      DataField       =   "Doðum Yeri"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   1800
      TabIndex        =   27
      Text            =   "Text4"
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      DataField       =   "Cinsiyet"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   1800
      TabIndex        =   23
      Text            =   "Text3"
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      DataField       =   "Ad"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   1800
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "TC Kimlik No"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   1800
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\hasta kayýt\hasta kayýt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "hstkyt"
      Top             =   9120
      Visible         =   0   'False
      Width           =   10575
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "MEDÝKAL BÝLGÝLER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   26
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label23 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "HASTA BÝLGÝLERÝ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   25
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ÝLETÝÞÝM BÝLGÝLERÝ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   24
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Diyabet:"
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
      Left            =   5400
      TabIndex        =   20
      Top             =   5880
      Width           =   3000
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "HIV:"
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
      Left            =   5400
      TabIndex        =   19
      Top             =   5160
      Width           =   3000
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hepatit:"
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
      Left            =   5400
      TabIndex        =   18
      Top             =   4440
      Width           =   3000
   End
   Begin VB.Label Label18 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Alkol:"
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
      Left            =   5400
      TabIndex        =   17
      Top             =   3720
      Width           =   3000
   End
   Begin VB.Label Label17 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sigara:"
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
      Left            =   5400
      TabIndex        =   16
      Top             =   3000
      Width           =   3000
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Kilo:"
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
      Left            =   5400
      TabIndex        =   15
      Top             =   2280
      Width           =   3000
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Boy:"
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
      Left            =   5400
      TabIndex        =   14
      Top             =   1560
      Width           =   3000
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Adres:"
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
      Left            =   10320
      TabIndex        =   13
      Top             =   3720
      Width           =   3000
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Adresi:"
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
      Left            =   10320
      TabIndex        =   12
      Top             =   3000
      Width           =   3000
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Telefon(Cep):"
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
      Left            =   10320
      TabIndex        =   11
      Top             =   2280
      Width           =   3000
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Telefon(Sabit):"
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
      Left            =   10320
      TabIndex        =   10
      Top             =   1560
      Width           =   3000
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Kurum:"
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
      Left            =   240
      TabIndex        =   9
      Top             =   7320
      Width           =   3000
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Meslek:"
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
      TabIndex        =   8
      Top             =   6480
      Width           =   3000
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Kan Grubu:"
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
      Left            =   5400
      TabIndex        =   7
      Top             =   6600
      Width           =   3000
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Baba Adý:"
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
      TabIndex        =   6
      Top             =   5880
      Width           =   3000
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ana Adý:"
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
      TabIndex        =   5
      Top             =   5160
      Width           =   3000
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Doðum Tarihi:"
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
      Top             =   4440
      Width           =   3000
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Doðum Yeri:"
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
      TabIndex        =   3
      Top             =   3720
      Width           =   3000
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cinsiyet:"
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
      TabIndex        =   2
      Top             =   3000
      Width           =   3000
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ad - Soyad :"
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
      TabIndex        =   1
      Top             =   2280
      Width           =   2880
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "T.C Kimlik No:"
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
      TabIndex        =   0
      Top             =   1560
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.UpdateRecord
Data1.Recordset.MoveLast
Command2.Visible = True
End Sub
Private Sub Command2_Click()
Data1.Recordset.AddNew
Command2.Visible = False
End Sub
Private Sub Command3_Click()
Data1.Recordset.Edit
Data1.UpdateRecord
End Sub
Private Sub Command4_Click()
Data1.Recordset.Delete
Data1.Recordset.MoveLast
End Sub

Private Sub Command5_Click()
ad = InputBox("Aramak istediðiniz hastanýn adýný giriniz", "Ada göre arama")
aranan = "ad = ' " & ad & "'"
Data1.Recordset.FindFirst aranan
End Sub

Private Sub Command6_Click()
Data1.Recordset.MoveFirst
End Sub
Private Sub Command7_Click()
If Data1.Recordset.BOF Then Exit Sub
Data1.Recordset.MovePrevious
End Sub
Private Sub Command8_Click()
If Data1.Recordset.EOF Then Exit Sub
Data1.Recordset.MoveNext
End Sub
Private Sub Command9_Click()
Data1.Recordset.MoveLast
End Sub

