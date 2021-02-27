VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3900
      ItemData        =   "ornek.frx":0000
      Left            =   7440
      List            =   "ornek.frx":00A6
      TabIndex        =   15
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Left            =   5160
      Top             =   9840
   End
   Begin VB.Timer Timer2 
      Left            =   9720
      Top             =   8520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "KELÝMEYÝ VE ANLAMINI SÖZLÜÐE EKLEMEK ÝÇÝN TIKLAYIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   8160
      Width           =   5175
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   4200
      TabIndex        =   6
      Text            =   "ANLAMINI GÝRÝN"
      Top             =   7320
      Width           =   6015
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Text            =   "KELÝMEYÝ GÝRÝN"
      Top             =   7320
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Left            =   8880
      Top             =   1440
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "ara"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FF00&
      Caption         =   "                   ARADIÐINIZ KELÝMENÝN ANLAMI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4560
      Width           =   7215
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FFFF&
      Caption         =   "                                         "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   7215
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
      Caption         =   "                            ARANILAN KELÝME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   7215
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   7215
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Caption         =   "                  PROGRAMIN YAPIMINDA EMEÐÝ GEÇEN HERKEZE TEÞEKKÜRLER SAYGILARIMLA !!BY ÖZGÜR!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   9600
      Width           =   10335
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      Caption         =   "     PROGRAMI HAZIRLAYAN=ÖZGÜR KAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   9000
      Width           =   9855
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "SÖZLÜÐE YENÝ KELÝME EKLEMEK ÝSTÝYORSANIZ KELÝMEYÝ VE ANLAMINI AÞAÐIYA YAZIN."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   6600
      Width           =   9615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "SÖZLÜÐÜMÜZDE BULUNAN                 KELÝMELER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "    COÐRAFÝ    TERÝMLER                       SÖZLÜÐÜ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "!!LÜTFEN KELÝMENÝZÝN HARFLERÝNÝ KÜÇÜK GÝRÝNÝZ!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set db = OpenDatabase("c:\sözlük\ornek.mdb")
Set Rs = db.OpenRecordset("tblornek", dbOpenSnapshot)
Do While Not Rs.EOF
If Text1.Text = Rs!kelime Then
Text1.Text = "aradýðýnýz kelime sözlüðümüzde bulundu"
Label9.Caption = Rs!kelime
Else:
Rs.MoveNext
End If
Loop
If Text1.Text <> "aradýðýnýz kelime sözlüðümüzde bulundu" Then Text1.Text = "aradýðýnýz kelime sözlüðümüzde bulunamadi"

If Text1.Text = "aradýðýnýz kelime sözlüðümüzde bulundu" Then

Rs.MoveFirst
Do Until Rs.EOF
If Label9.Caption = "açýk havza" Then
a = Rs!anlam
ElseIf Label9.Caption = "açýsal hýz" Then
Rs.MoveNext
a = Rs!anlam
ElseIf Label9.Caption = "aðýl" Then
For sayac = 1 To 2
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "akarsu" Then
For sayac = 1 To 3
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "akarsu akýmý" Then
For sayac = 1 To 4
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "akarsu rejimi" Then
For sayac = 1 To 5
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "alizeler" Then
For sayac = 1 To 6
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "altimetre" Then
For sayac = 1 To 7
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "ana yön" Then
For sayac = 1 To 8
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "andezit" Then
For sayac = 1 To 9
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "anemometre" Then
For sayac = 1 To 10
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "aneroid barometre" Then
For sayac = 1 To 11
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "artezyen" Then
For sayac = 1 To 12
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "atmosfer" Then
For sayac = 1 To 13
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "atmosfer basýncý" Then
For sayac = 1 To 14
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "aysberg" Then
For sayac = 1 To 15
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "ay tutulmasý" Then
For sayac = 1 To 16
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "baðýl nem" Then
For sayac = 1 To 17
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "bankiz" Then
For sayac = 1 To 18
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "baraj gölü" Then
For sayac = 1 To 19
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "barograf" Then
For sayac = 1 To 20
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "bazalt" Then
For sayac = 1 To 21
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "birinci zaman" Then
For sayac = 1 To 22
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "bora" Then
For sayac = 1 To 23
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "boylam" Then
For sayac = 1 To 24
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "boyun" Then
For sayac = 1 To 25
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "bozkýr" Then
For sayac = 1 To 26
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "buzul gölleri" Then
For sayac = 1 To 27
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "coðrafi bölge" Then
For sayac = 1 To 28
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "coðrafi konum" Then
For sayac = 1 To 29
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "çakýltaþý" Then
For sayac = 1 To 30
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "çakmaktaþý" Then
For sayac = 1 To 31
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "çay" Then
For sayac = 1 To 32
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "çekirdek" Then
For sayac = 1 To 33
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "çýð" Then
For sayac = 1 To 34
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "çiy" Then
For sayac = 1 To 35
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "çizgi ölçek" Then
For sayac = 1 To 36
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "çizgisel hýz" Then
For sayac = 1 To 37
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "çökme dolini" Then
For sayac = 1 To 38
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "dað" Then
For sayac = 1 To 39
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "dalgalar" Then
For sayac = 1 To 40
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "dam" Then
For sayac = 1 To 41
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "delta" Then
For sayac = 1 To 42
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "deniz" Then
For sayac = 1 To 43
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "deprem" Then
For sayac = 1 To 44
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "dere" Then
For sayac = 1 To 45
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "diyorit" Then
For sayac = 1 To 46
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "doðal bitki örtüsü" Then
For sayac = 1 To 47
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "dolin" Then
For sayac = 1 To 48
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "don olayý" Then
For sayac = 1 To 49
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "doruk" Then
For sayac = 1 To 50
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "dördüncü zaman" Then
For sayac = 1 To 51
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "düden" Then
For sayac = 1 To 52
Rs.MoveNext
Next sayac
a = Rs!anlam
ElseIf Label9.Caption = "dünya" Then
For sayac = 1 To 53
Rs.MoveNext
Next sayac
a = Rs!anlam
End If
GoTo dvm1
Rs.MoveNext
Loop
dvm1:
Label7.Caption = a
Else
Label7.Caption = "sözlüðümüzde bulunmayan bir kelime girdiniz"
Label9.Caption = "sözlüðümüzde bulunmayan bir kelime girdiniz"
End If
End Sub

Private Sub Command2_Click()
cevap = MsgBox("bu kelimeyi sözlüðe eklemek istiyormusunuz?", 36, "kayýt kutusu")
If cevap = 6 Then
rstblornek.AddNew
rstblornek!kelime = Text3.Text
rstblornek!anlam = Text4.Text
rstblornek.Update

ElseIf cevap = 7 Then
MsgBox ("kelime sözlüðe eklenemedi!")
End If
End Sub

Private Sub Form_Load()
Dim db As Database
Dim Rs As Recordset
Label1.Caption = "                                                    !!!   Lütfen kelimeyi küçük harfle yazýn!!!!   "
Label1.Left = ScaleLeft
Label1.Width = Me.ScaleWidth
Timer1.Interval = 100


Label5.Caption = "            PROGRAMI HAZIRLAYAN=ÖZGÜR KAN "
Label5.Left = ScaleLeft
Label5.Width = Me.ScaleWidth
Timer2.Interval = 100


Label6.Caption = "               PROGRAMIN YAPIMINDA EMEÐÝ GEÇEN HERKEZE TEÞEKÜRLER !!SAYGILARIMLA BY ÖZGÜR!!"
Label6.Left = ScaleLeft
Label6.Width = Me.ScaleWidth
Timer3.Interval = 100


Form1.Height = 0
    Form1.Width = 0
    For i = 1 To 141
      Form1.Width = Form1.Width + i
      Form1.Height = Form1.Height + i
      Form1.Show
      Form1.Refresh
    Next i







End Sub

Private Sub Form_Unload(Cancel As Integer)
cevap = MsgBox("çýkmak istiyormusunuz?", 36, "onay kutusu")
If cevap = 6 Then
MsgBox ("Çýkýþ iþlemi yapýlýyor.Biz terhic ettiðiniz için teþekkürler.")
Cancel = False
ElseIf cevap = 7 Then
MsgBox ("Çýkýþ iþlemi gerçekleþtirilemedi.")
Cancel = True
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label1 = Mid(Label1, 2, Len(Label1) - 1) + Left(Label1, 1)
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Label5 = Mid(Label5, 2, Len(Label5) - 1) + Left(Label5, 1)
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Label6 = Mid(Label6, 2, Len(Label6) - 1) + Left(Label6, 1)
End Sub
