VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   9210
   ClientLeft      =   2565
   ClientTop       =   645
   ClientWidth     =   9705
   LinkTopic       =   "Form7"
   ScaleHeight     =   9210
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   9855
      Left            =   0
      Picture         =   "Form7.frx":0000
      ScaleHeight     =   9795
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5040
         TabIndex        =   37
         Top             =   840
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5040
         TabIndex        =   36
         Top             =   840
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "T.C NUMARASINA GÖRE ARA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "T.C NUMARASINA GÖRE ARAMA YAPMAK ÝÇÝN TIKLAYIN."
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "ODA NUMARASINA GÖRE ARA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "ODA NUMARASINA GÖRE ARAMA YAPMAK ÝÇÝN TIKLAYIN."
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "TELEFON NUMARASINA GÖRE ARA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "TELEFON NUMARASINA GÖRE ARAMA YAPMAK ÝÇÝN TIKLAYIN."
         Top             =   4680
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080C0FF&
         Caption         =   "ANA SAYFAYA DÖN"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "ANA SAYFAYA DÖNMEK ÝÇÝN TIKLAYIN."
         Top             =   5400
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5040
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0080C0FF&
         Caption         =   "ARA"
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "MÜÞTERÝ ARAMAK ÝÇÝN TIKLAYIN."
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0080C0FF&
         Caption         =   "ARA"
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "MÜÞTERÝ ARAMAK ÝÇÝN TIKLAYIN."
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080C0FF&
         Caption         =   "ARA"
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "MÜÞTERÝ ARAMAK ÝÇÝN TIKLAYIN."
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0080C0FF&
         Caption         =   "ARAMA SAYFASINA DÖN <========"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "ARAMA SAYFASINA DÖNMEK ÝÇÝN TIKLAYIN."
         Top             =   8040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
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
         Height          =   615
         Left            =   5160
         TabIndex        =   35
         Top             =   2280
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label25 
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
         Height          =   615
         Left            =   2280
         TabIndex        =   34
         Top             =   2280
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "ARAMA SAYFASI"
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
         Height          =   615
         Left            =   2280
         TabIndex        =   33
         Top             =   120
         Width           =   6375
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
         Height          =   615
         Left            =   2280
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label3 
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
         Left            =   2280
         TabIndex        =   31
         Top             =   3000
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label4 
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
         Left            =   5160
         TabIndex        =   30
         Top             =   3000
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "SOYADI===>"
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
         Left            =   2280
         TabIndex        =   29
         Top             =   3480
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label6 
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
         Left            =   5160
         TabIndex        =   28
         Top             =   3480
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   "BABA ADI===>"
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
         Left            =   2280
         TabIndex        =   27
         Top             =   3960
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label8 
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
         Left            =   5160
         TabIndex        =   26
         Top             =   3960
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         Caption         =   "ANNE ADI===>"
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
         Left            =   2280
         TabIndex        =   25
         Top             =   4440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label10 
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
         Left            =   5160
         TabIndex        =   24
         Top             =   4440
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         Caption         =   "DOÐUM TARÝHÝ===>"
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
         Left            =   2280
         TabIndex        =   23
         Top             =   4920
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label12 
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
         Left            =   5160
         TabIndex        =   22
         Top             =   4920
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
         Caption         =   "CÝNSÝYETÝ===>"
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
         Left            =   2280
         TabIndex        =   21
         Top             =   5400
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label14 
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
         Left            =   5160
         TabIndex        =   20
         Top             =   5400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080C0FF&
         Caption         =   "MEDENÝ HALÝ===>"
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
         Left            =   2280
         TabIndex        =   19
         Top             =   5880
         Visible         =   0   'False
         Width           =   2655
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
         Height          =   375
         Left            =   5160
         TabIndex        =   18
         Top             =   5880
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label17 
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
         Height          =   975
         Left            =   2280
         TabIndex        =   17
         Top             =   6360
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5160
         TabIndex        =   16
         Top             =   6360
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label19 
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
         Height          =   735
         Left            =   2280
         TabIndex        =   15
         Top             =   7440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label20 
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
         Height          =   735
         Left            =   5160
         TabIndex        =   14
         Top             =   7440
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label21 
         BackColor       =   &H0080C0FF&
         Caption         =   "GELÝÞ TARÝHÝ==>"
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
         Left            =   2280
         TabIndex        =   13
         Top             =   8280
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label22 
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
         Left            =   5160
         TabIndex        =   12
         Top             =   8280
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label23 
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
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   8760
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label24 
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
         Left            =   5160
         TabIndex        =   10
         Top             =   8760
         Visible         =   0   'False
         Width           =   3495
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command5.Visible = False
Label1.Caption = "T.C NUMARASINA GÖRE ARAMA"
Label2.Visible = True
Label2.Caption = "T.C NUMARASI GÝRÝNÝZ===>"
Text1.Visible = True
Command6.Visible = True
Command10.Visible = True
Label25.Caption = "ODA NUMARASI===>"
Label19.Caption = "TELEFON NUMARASI===>"
End Sub

Private Sub Command10_Click()
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command5.Visible = True
Label1.Caption = "ARAMA SAYFASI"
Label2.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command10.Visible = False
Text1 = ""
Text2 = ""
Text3 = ""
Label4.Caption = ""
Label6.Caption = ""
Label8.Caption = ""
Label10.Caption = ""
Label12.Caption = ""
Label14.Caption = ""
Label16.Caption = ""
Label18.Caption = ""
Label20.Caption = ""
Label22.Caption = ""
Label24.Caption = ""
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
End Sub

Private Sub Command2_Click()
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command5.Visible = False
Label1.Caption = "ODA NUMARASINA GÖRE ARAMA"
Label2.Visible = True
Label2.Caption = "ODA NUMARASI GÝRÝNÝZ===>"
Text2.Visible = True
Command7.Visible = True
Command10.Visible = True
Label25.Caption = "T.C KÝMLÝK NUMARASI==>"
Label19.Caption = "TELEFON NUMARASI===>"
End Sub

Private Sub Command3_Click()
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command5.Visible = False
Label1.Caption = "TELEFON NUMARASINA GÖRE ARAMA"
Label2.Visible = True
Label2.Caption = "TELEFON NUMARASI GÝRÝNÝZ===>"
Text3.Visible = True
Command8.Visible = True
Command10.Visible = True
Label25.Caption = "T.C KÝMLÝK NUMARASI==>"
Label19.Caption = "ODA NUMARASI===>"
End Sub
Private Sub Command5_Click()
Form7.Hide
Form3.Show
End Sub
Private Sub Command6_Click()
On Error Resume Next
If Len(Text1) < 11 Then
MsgBox "T.C KÝMLÝK NUMARASI 11 HANELÝ OLMALIDIR."
Text1.Text = ""
End If
If Text1 = "" Then
MsgBox "T.C KÝMLÝK NUMARASI GÝRÝNÝZ."
Else
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
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
Label18.Visible = True
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True
Label22.Visible = True
Label23.Visible = True
Label24.Visible = True
Label25.Visible = True
Label26.Visible = True
End If
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If Text1 = rs!Tc Then
Label4.Caption = rs!Adý
Label6.Caption = rs!soyadý
Label8.Caption = rs!Baba_adý
Label10.Caption = rs!Anne_adý
Label12.Caption = rs!Doðum_Tarih
Label14.Caption = rs!Cinsiyet
Label16.Caption = rs!Medeni_Hali
Label18.Caption = rs!ikametgah_Adresi
Label20.Caption = rs!Telefon
Label22.Caption = rs!Geliþ_Tarihi
Label24.Caption = rs!Geliþ_Saati
Label26.Caption = rs!OdaNo

Exit For
Else
rs.MoveNext
End If
Next sayac
If Label4.Caption = "" And Label4.Visible = True Then
MsgBox "MÜÞTERÝ BULUNAMADI."
Text1.Text = ""
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
ElseIf Label4.Caption <> "" And Label4.Visible = True Then
MsgBox "MÜÞTERÝ BULUNDU."
End If
End Sub

Private Sub Command7_Click()
On Error Resume Next
If Text2 = "" Then
MsgBox "ODA  NUMARASI GÝRÝNÝZ."
Else
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
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
Label18.Visible = True
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True
Label22.Visible = True
Label23.Visible = True
Label24.Visible = True
Label25.Visible = True
Label26.Visible = True
End If
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If Text2 = rs!OdaNo Then
Label4.Caption = rs!Adý
Label6.Caption = rs!soyadý
Label8.Caption = rs!Baba_adý
Label10.Caption = rs!Anne_adý
Label12.Caption = rs!Doðum_Tarih
Label14.Caption = rs!Cinsiyet
Label16.Caption = rs!Medeni_Hali
Label18.Caption = rs!ikametgah_Adresi
Label20.Caption = rs!Telefon
Label22.Caption = rs!Geliþ_Tarihi
Label24.Caption = rs!Geliþ_Saati
Label26.Caption = rs!Tc

Exit For
Else
rs.MoveNext
End If
Next sayac
If Label4.Caption = "" And Label4.Visible = True Then
MsgBox "MÜÞTERÝ BULUNAMADI."
Text2.Text = ""
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
ElseIf Label4.Caption <> "" And Label4.Visible = True Then
MsgBox "MÜÞTERÝ BULUNDU."
End If
End Sub
Private Sub Command8_Click()
On Error Resume Next
If Text3 = "" Then
MsgBox "TELEFON NUMARASI GÝRÝNÝZ."
Else
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
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
Label18.Visible = True
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True
Label22.Visible = True
Label23.Visible = True
Label24.Visible = True
Label25.Visible = True
Label26.Visible = True
End If
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If Text3 = rs!Telefon Then
Label4.Caption = rs!Adý
Label6.Caption = rs!soyadý
Label8.Caption = rs!Baba_adý
Label10.Caption = rs!Anne_adý
Label12.Caption = rs!Doðum_Tarih
Label14.Caption = rs!Cinsiyet
Label16.Caption = rs!Medeni_Hali
Label18.Caption = rs!ikametgah_Adresi
Label20.Caption = rs!OdaNo
Label22.Caption = rs!Geliþ_Tarihi
Label24.Caption = rs!Geliþ_Saati
Label26.Caption = rs!Tc
Exit For
Else
rs.MoveNext
End If
Next sayac
If Label4.Caption = "" And Label4.Visible = True Then
MsgBox "MÜÞTERÝ BULUNAMADI."
Text3.Text = ""
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
ElseIf Label4.Caption <> "" And Label4.Visible = True Then
MsgBox "MÜÞTERÝ BULUNDU."
End If
End Sub





Private Sub Text1_Change()
If Len(Text1) > 11 Then
MsgBox "T.C KÝMLÝK NUMARASI 11 HANELÝ OLMALIDIR."
Text1.Text = ""
End If
Label4.Caption = ""
Label6.Caption = ""
Label8.Caption = ""
Label10.Caption = ""
Label12.Caption = ""
Label14.Caption = ""
Label16.Caption = ""
Label18.Caption = ""
Label20.Caption = ""
Label22.Caption = ""
Label24.Caption = ""
Label26.Caption = ""
End Sub
Private Sub Text2_Change()
If Len(Text2) > 1 Then
MsgBox "EN FAZLA 9 TANE ODA VAR."
Text2.Text = ""
End If
Label4.Caption = ""
Label6.Caption = ""
Label8.Caption = ""
Label10.Caption = ""
Label12.Caption = ""
Label14.Caption = ""
Label16.Caption = ""
Label18.Caption = ""
Label20.Caption = ""
Label22.Caption = ""
Label24.Caption = ""
Label26.Caption = ""
End Sub
Private Sub Text3_Change()
If Len(Text3) > 10 Then
MsgBox "TELEFON NUMARASI 10 HANELÝ OLMALIDIR."
Text3.Text = ""
End If
Label4.Caption = ""
Label6.Caption = ""
Label8.Caption = ""
Label10.Caption = ""
Label12.Caption = ""
Label14.Caption = ""
Label16.Caption = ""
Label18.Caption = ""
Label20.Caption = ""
Label22.Caption = ""
Label24.Caption = ""
Label26.Caption = ""
End Sub
