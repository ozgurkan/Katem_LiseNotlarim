VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   0  'None
   Caption         =   "Form9"
   ClientHeight    =   7425
   ClientLeft      =   2055
   ClientTop       =   1275
   ClientWidth     =   10440
   LinkTopic       =   "Form9"
   ScaleHeight     =   7425
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture10 
      Height          =   7455
      Left            =   0
      Picture         =   "Form9.frx":0000
      ScaleHeight     =   7395
      ScaleWidth      =   10395
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.PictureBox Picture1 
         Height          =   3375
         Left            =   4800
         Picture         =   "Form9.frx":114582
         ScaleHeight     =   3315
         ScaleWidth      =   5115
         TabIndex        =   111
         Top             =   1080
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.PictureBox Picture2 
         Height          =   4215
         Left            =   4800
         Picture         =   "Form9.frx":14D8DC
         ScaleHeight     =   4155
         ScaleWidth      =   5115
         TabIndex        =   110
         Top             =   1080
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.PictureBox Picture3 
         Height          =   3495
         Left            =   4680
         Picture         =   "Form9.frx":19427E
         ScaleHeight     =   3435
         ScaleWidth      =   5235
         TabIndex        =   109
         Top             =   1080
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   240
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   720
         Top             =   240
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "SLAYTI DURDUR"
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
         TabIndex        =   108
         ToolTipText     =   "SLAYTI DURDURMAK ÝÇÝN TIKLAYIN."
         Top             =   6120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "SLAYTI BAÞLAT"
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
         TabIndex        =   107
         ToolTipText     =   "SLAYTI BAÞLATMAK ÝÇÝN TIKLAYIN."
         Top             =   6120
         Width           =   2055
      End
      Begin VB.PictureBox Picture4 
         Height          =   3135
         Left            =   4440
         Picture         =   "Form9.frx":1CFC98
         ScaleHeight     =   3075
         ScaleWidth      =   5475
         TabIndex        =   106
         Top             =   1080
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.PictureBox Picture5 
         Height          =   3735
         Left            =   4440
         Picture         =   "Form9.frx":20764E
         ScaleHeight     =   3675
         ScaleWidth      =   5595
         TabIndex        =   105
         Top             =   1080
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.PictureBox Picture6 
         Height          =   3615
         Left            =   5280
         Picture         =   "Form9.frx":24C39C
         ScaleHeight     =   3555
         ScaleWidth      =   4755
         TabIndex        =   104
         Top             =   1080
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.PictureBox Picture7 
         Height          =   3375
         Left            =   4920
         Picture         =   "Form9.frx":284F62
         ScaleHeight     =   3315
         ScaleWidth      =   5115
         TabIndex        =   103
         Top             =   1080
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.PictureBox Picture8 
         Height          =   3975
         Left            =   4200
         Picture         =   "Form9.frx":2BDFA4
         ScaleHeight     =   3915
         ScaleWidth      =   5835
         TabIndex        =   102
         Top             =   1080
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.PictureBox Picture9 
         Height          =   4455
         Left            =   4200
         Picture         =   "Form9.frx":308EDE
         ScaleHeight     =   4395
         ScaleWidth      =   5835
         TabIndex        =   101
         Top             =   1080
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0000FFFF&
         Caption         =   "<===1.ODA===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4575
         Left            =   0
         TabIndex        =   90
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            Caption         =   "GENEL ÖZELLÝKLERÝ"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   720
            TabIndex        =   100
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000FFFF&
            Caption         =   "Banyo ve Tuvalet=VAR"
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
            TabIndex        =   99
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label4 
            BackColor       =   &H0000FFFF&
            Caption         =   "Klima ve Havalandýrma=VAR"
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
            TabIndex        =   98
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label5 
            BackColor       =   &H0000FFFF&
            Caption         =   "Mini Bar=VAR"
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
            TabIndex        =   97
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label6 
            BackColor       =   &H0000FFFF&
            Caption         =   "Telefon=VAR"
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
            TabIndex        =   96
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label7 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda Servisi=VAR"
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
            TabIndex        =   95
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label8 
            BackColor       =   &H0000FFFF&
            Caption         =   "Jakuzi=YOK"
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
            TabIndex        =   94
            Top             =   2640
            Width           =   3255
         End
         Begin VB.Label Label9 
            BackColor       =   &H0000FFFF&
            Caption         =   "Yatak=1 ADET ÇÝFT KÝÞÝLÝK"
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
            TabIndex        =   93
            Top             =   3000
            Width           =   3855
         End
         Begin VB.Label Label10 
            BackColor       =   &H0000FFFF&
            Caption         =   "Televizyon=VAR"
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
            TabIndex        =   92
            Top             =   3360
            Width           =   3255
         End
         Begin VB.Label Label11 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda durumu="
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
            TabIndex        =   91
            Top             =   3720
            Width           =   3255
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0000FFFF&
         Caption         =   "<===2.ODA===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4575
         Left            =   0
         TabIndex        =   79
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Label Label12 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda durumu="
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
            TabIndex        =   89
            Top             =   3720
            Width           =   3255
         End
         Begin VB.Label Label13 
            BackColor       =   &H0000FFFF&
            Caption         =   "Televizyon=YOK"
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
            TabIndex        =   88
            Top             =   3360
            Width           =   3255
         End
         Begin VB.Label Label14 
            BackColor       =   &H0000FFFF&
            Caption         =   "Yatak=1 ADET ÇÝFT KÝÞÝLÝK"
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
            TabIndex        =   87
            Top             =   3000
            Width           =   3735
         End
         Begin VB.Label Label15 
            BackColor       =   &H0000FFFF&
            Caption         =   "Jakuzi=YOK"
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
            TabIndex        =   86
            Top             =   2640
            Width           =   3255
         End
         Begin VB.Label Label16 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda Servisi=VAR"
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
            TabIndex        =   85
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label17 
            BackColor       =   &H0000FFFF&
            Caption         =   "Telefon=VAR"
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
            TabIndex        =   84
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label18 
            BackColor       =   &H0000FFFF&
            Caption         =   "Mini Bar=YOK"
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
            TabIndex        =   83
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label19 
            BackColor       =   &H0000FFFF&
            Caption         =   "Klima ve Havalandýrma=VAR"
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
            TabIndex        =   82
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label20 
            BackColor       =   &H0000FFFF&
            Caption         =   "Banyo ve Tuvalet=VAR"
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
            TabIndex        =   81
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            Caption         =   "GENEL ÖZELLÝKLERÝ"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   720
            TabIndex        =   80
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0000FFFF&
         Caption         =   "<===3.ODA===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4575
         Left            =   0
         TabIndex        =   68
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            Caption         =   "GENEL ÖZELLÝKLERÝ"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   720
            TabIndex        =   78
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label23 
            BackColor       =   &H0000FFFF&
            Caption         =   "Banyo ve Tuvalet=VAR"
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
            TabIndex        =   77
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label24 
            BackColor       =   &H0000FFFF&
            Caption         =   "Klima ve Havalandýrma=VAR"
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
            TabIndex        =   76
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label25 
            BackColor       =   &H0000FFFF&
            Caption         =   "Mini Bar=VAR"
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
            TabIndex        =   75
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label26 
            BackColor       =   &H0000FFFF&
            Caption         =   "Telefon=VAR"
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
            TabIndex        =   74
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label27 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda Servisi=VAR"
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
            TabIndex        =   73
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label28 
            BackColor       =   &H0000FFFF&
            Caption         =   "Jakuzi=VAR"
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
            TabIndex        =   72
            Top             =   2640
            Width           =   3255
         End
         Begin VB.Label Label29 
            BackColor       =   &H0000FFFF&
            Caption         =   "Yatak=1 ADET ÇÝFT KÝÞÝLÝK"
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
            TabIndex        =   71
            Top             =   3000
            Width           =   3735
         End
         Begin VB.Label Label30 
            BackColor       =   &H0000FFFF&
            Caption         =   "Televizyon=VAR"
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
            TabIndex        =   70
            Top             =   3360
            Width           =   3255
         End
         Begin VB.Label Label31 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda durumu="
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
            TabIndex        =   69
            Top             =   3720
            Width           =   3255
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0000FFFF&
         Caption         =   "<===4.ODA===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4575
         Left            =   0
         TabIndex        =   57
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Label Label32 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda durumu="
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
            TabIndex        =   67
            Top             =   3720
            Width           =   3255
         End
         Begin VB.Label Label33 
            BackColor       =   &H0000FFFF&
            Caption         =   "Televizyon=VAR"
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
            TabIndex        =   66
            Top             =   3360
            Width           =   3255
         End
         Begin VB.Label Label34 
            BackColor       =   &H0000FFFF&
            Caption         =   "Yatak=2 ADET ÇÝFT KÝÞÝLÝK "
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
            TabIndex        =   65
            Top             =   3000
            Width           =   3855
         End
         Begin VB.Label Label35 
            BackColor       =   &H0000FFFF&
            Caption         =   "Jakuzi=VAR"
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
            TabIndex        =   64
            Top             =   2640
            Width           =   3255
         End
         Begin VB.Label Label36 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda Servisi=VAR"
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
            TabIndex        =   63
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label37 
            BackColor       =   &H0000FFFF&
            Caption         =   "Telefon=VAR"
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
            TabIndex        =   62
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label38 
            BackColor       =   &H0000FFFF&
            Caption         =   "Mini Bar=VAR"
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
            TabIndex        =   61
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label39 
            BackColor       =   &H0000FFFF&
            Caption         =   "Klima ve Havalandýrma=VAR"
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
            TabIndex        =   60
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label40 
            BackColor       =   &H0000FFFF&
            Caption         =   "Banyo ve Tuvalet=VAR"
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
            TabIndex        =   59
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label41 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            Caption         =   "GENEL ÖZELLÝKLERÝ"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   720
            TabIndex        =   58
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H0000FFFF&
         Caption         =   "<===5.ODA===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4575
         Left            =   0
         TabIndex        =   46
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            Caption         =   "GENEL ÖZELLÝKLERÝ"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   720
            TabIndex        =   56
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label43 
            BackColor       =   &H0000FFFF&
            Caption         =   "Banyo ve Tuvalet=VAR"
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
            TabIndex        =   55
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label44 
            BackColor       =   &H0000FFFF&
            Caption         =   "Klima ve Havalandýrma=VAR"
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
            TabIndex        =   54
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label45 
            BackColor       =   &H0000FFFF&
            Caption         =   "Mini Bar=YOK"
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
            TabIndex        =   53
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label46 
            BackColor       =   &H0000FFFF&
            Caption         =   "Telefon=YOK"
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
            TabIndex        =   52
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label47 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda Servisi=VAR"
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
            TabIndex        =   51
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label48 
            BackColor       =   &H0000FFFF&
            Caption         =   "Jakuzi=VAR"
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
            TabIndex        =   50
            Top             =   2640
            Width           =   3255
         End
         Begin VB.Label Label49 
            BackColor       =   &H0000FFFF&
            Caption         =   "Yatak=1 ADET ÇÝFT KÝÞÝLÝK "
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
            TabIndex        =   49
            Top             =   3000
            Width           =   3855
         End
         Begin VB.Label Label50 
            BackColor       =   &H0000FFFF&
            Caption         =   "Televizyon=VAR"
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
            TabIndex        =   48
            Top             =   3360
            Width           =   3255
         End
         Begin VB.Label Label51 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda durumu="
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
            TabIndex        =   47
            Top             =   3720
            Width           =   3255
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H0000FFFF&
         Caption         =   "<===6.ODA===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4575
         Left            =   0
         TabIndex        =   35
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Label Label52 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda durumu="
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
            TabIndex        =   45
            Top             =   3720
            Width           =   3255
         End
         Begin VB.Label Label53 
            BackColor       =   &H0000FFFF&
            Caption         =   "Televizyon=VAR"
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
            TabIndex        =   44
            Top             =   3360
            Width           =   3255
         End
         Begin VB.Label Label54 
            BackColor       =   &H0000FFFF&
            Caption         =   "Yatak=1 ADET TEK KÝÞÝLÝK "
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
            TabIndex        =   43
            Top             =   3000
            Width           =   3855
         End
         Begin VB.Label Label55 
            BackColor       =   &H0000FFFF&
            Caption         =   "Jakuzi=VAR"
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
            TabIndex        =   42
            Top             =   2640
            Width           =   3255
         End
         Begin VB.Label Label56 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda Servisi=VAR"
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
            TabIndex        =   41
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label57 
            BackColor       =   &H0000FFFF&
            Caption         =   "Telefon=VAR"
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
            TabIndex        =   40
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label58 
            BackColor       =   &H0000FFFF&
            Caption         =   "Mini Bar=YOK"
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
            TabIndex        =   39
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label59 
            BackColor       =   &H0000FFFF&
            Caption         =   "Klima ve Havalandýrma=VAR"
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
            TabIndex        =   38
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label60 
            BackColor       =   &H0000FFFF&
            Caption         =   "Banyo ve Tuvalet=VAR"
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
            TabIndex        =   37
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label61 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            Caption         =   "GENEL ÖZELLÝKLERÝ"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   720
            TabIndex        =   36
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H0000FFFF&
         Caption         =   "<===7.ODA===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4575
         Left            =   0
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Label Label62 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            Caption         =   "GENEL ÖZELLÝKLERÝ"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   720
            TabIndex        =   34
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label63 
            BackColor       =   &H0000FFFF&
            Caption         =   "Banyo ve Tuvalet=VAR"
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
            TabIndex        =   33
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label64 
            BackColor       =   &H0000FFFF&
            Caption         =   "Klima ve Havalandýrma=VAR"
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
            TabIndex        =   32
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label65 
            BackColor       =   &H0000FFFF&
            Caption         =   "Mini Bar=YOK"
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
            TabIndex        =   31
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label66 
            BackColor       =   &H0000FFFF&
            Caption         =   "Telefon=VAR"
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
            TabIndex        =   30
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label67 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda Servisi=VAR"
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
            TabIndex        =   29
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label68 
            BackColor       =   &H0000FFFF&
            Caption         =   "Jakuzi=VAR"
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
            TabIndex        =   28
            Top             =   2640
            Width           =   3255
         End
         Begin VB.Label Label69 
            BackColor       =   &H0000FFFF&
            Caption         =   "Yatak=1 ADET ÇÝFT KÝÞÝLÝK "
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
            TabIndex        =   27
            Top             =   3000
            Width           =   3855
         End
         Begin VB.Label Label70 
            BackColor       =   &H0000FFFF&
            Caption         =   "Televizyon=VAR"
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
            TabIndex        =   26
            Top             =   3360
            Width           =   3255
         End
         Begin VB.Label Label71 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda durumu="
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
            Top             =   3720
            Width           =   3255
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H0000FFFF&
         Caption         =   "<===8.ODA===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4575
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Label Label72 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda durumu="
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
            TabIndex        =   23
            Top             =   3720
            Width           =   3255
         End
         Begin VB.Label Label73 
            BackColor       =   &H0000FFFF&
            Caption         =   "Televizyon=VAR"
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
            Top             =   3360
            Width           =   3255
         End
         Begin VB.Label Label74 
            BackColor       =   &H0000FFFF&
            Caption         =   "Yatak=2 ADET TEK KÝÞÝLÝK "
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
            TabIndex        =   21
            Top             =   3000
            Width           =   3855
         End
         Begin VB.Label Label75 
            BackColor       =   &H0000FFFF&
            Caption         =   "Jakuzi=YOK"
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
            TabIndex        =   20
            Top             =   2640
            Width           =   3255
         End
         Begin VB.Label Label76 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda Servisi=VAR"
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
            TabIndex        =   19
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label77 
            BackColor       =   &H0000FFFF&
            Caption         =   "Telefon=VAR"
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
            TabIndex        =   18
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label78 
            BackColor       =   &H0000FFFF&
            Caption         =   "Mini Bar=YOK"
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
            TabIndex        =   17
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label79 
            BackColor       =   &H0000FFFF&
            Caption         =   "Klima ve Havalandýrma=YOK"
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
            TabIndex        =   16
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label80 
            BackColor       =   &H0000FFFF&
            Caption         =   "Banyo ve Tuvalet=VAR"
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
            TabIndex        =   15
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label81 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            Caption         =   "GENEL ÖZELLÝKLERÝ"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   720
            TabIndex        =   14
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H0000FFFF&
         Caption         =   "<===9.ODA===>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4575
         Left            =   0
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Label Label82 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            Caption         =   "GENEL ÖZELLÝKLERÝ"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   720
            TabIndex        =   12
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label83 
            BackColor       =   &H0000FFFF&
            Caption         =   "Banyo ve Tuvalet=VAR"
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
            TabIndex        =   11
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label84 
            BackColor       =   &H0000FFFF&
            Caption         =   "Klima ve Havalandýrma=YOK"
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
            TabIndex        =   10
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label85 
            BackColor       =   &H0000FFFF&
            Caption         =   "Mini Bar=VAR"
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
            TabIndex        =   9
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label86 
            BackColor       =   &H0000FFFF&
            Caption         =   "Telefon=VAR"
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
            TabIndex        =   8
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label87 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda Servisi=VAR"
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
            TabIndex        =   7
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label88 
            BackColor       =   &H0000FFFF&
            Caption         =   "Jakuzi=YOK"
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
            TabIndex        =   6
            Top             =   2640
            Width           =   3255
         End
         Begin VB.Label Label89 
            BackColor       =   &H0000FFFF&
            Caption         =   "Yatak=1 ADET ÇÝFT KÝÞÝLÝK "
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
            TabIndex        =   5
            Top             =   3000
            Width           =   3855
         End
         Begin VB.Label Label90 
            BackColor       =   &H0000FFFF&
            Caption         =   "Televizyon=YOK"
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
            TabIndex        =   4
            Top             =   3360
            Width           =   3255
         End
         Begin VB.Label Label91 
            BackColor       =   &H0000FFFF&
            Caption         =   "Oda durumu="
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
            TabIndex        =   3
            Top             =   3720
            Width           =   3255
         End
      End
      Begin VB.CommandButton Command3 
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
         TabIndex        =   1
         ToolTipText     =   "ANA SAYFAYA DÖNMEK ÝÇÝN TIKLAYIN"
         Top             =   6120
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "SLAYT OYNATILIYOR."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   114
         Top             =   720
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label92 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "NOT=SLAYTIN AÇILMASI 10-15 SANÝYE SÜREBÝLÝR."
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
         Left            =   720
         TabIndex        =   113
         Top             =   5640
         Visible         =   0   'False
         Width           =   7215
      End
      Begin VB.Label Label93 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "ODA DURUMU VE ÖZELLÝKLERÝ"
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
         Height          =   495
         Left            =   1800
         TabIndex        =   112
         Top             =   120
         Width           =   7215
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Caption = "SLAYT DURDU."
Label1.BackColor = vbRed
Command2.Visible = True
Command1.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
End Sub
Private Sub Command2_Click()
Label92.Visible = True
Timer1.Enabled = True
If Picture1.Visible = True Or Picture3.Visible = True Or Picture5.Visible = True Or Picture7.Visible = True Or Picture9.Visible = True Then
Timer2.Enabled = True
ElseIf Picture2.Visible = True Or Picture4.Visible = True Or Picture6.Visible = True Or Picture8.Visible = True Then
Timer1.Enabled = True
Else
Timer1.Enabled = True
End If
Command2.Visible = False
Command1.Visible = True
Label1.Visible = True
Label1.Caption = "SLAYT OYNATILIYOR."
Label1.BackColor = vbGreen
End Sub

Private Sub Command3_Click()
Label92.Visible = False
Form9.Hide
Form3.Show
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame8.Visible = False
Frame9.Visible = False
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = False
Command1.Visible = False
Command2.Visible = True
Timer1.Enabled = False
Timer2.Enabled = False
Label1.Visible = False
End Sub



Private Sub Form_Load()
On Error Resume Next
'1.kayýt
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If rs!OdaNo <> 1 Then
Label11.Caption = "Oda durumu=BOÞ"
rs.MoveNext
Else
Label11.Caption = "Oda durumu=DOLU"
End If
Next sayac
rs.Close
'2.kayýt
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If rs!OdaNo <> 2 Then
Label12.Caption = "Oda durumu=BOÞ"
rs.MoveNext
Else
Label12.Caption = "Oda durumu=DOLU"
End If
Next sayac
rs.Close
'3.kayýt
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If rs!OdaNo <> 3 Then
Label31.Caption = "Oda durumu=BOÞ"
rs.MoveNext
Else
Label31.Caption = "Oda durumu=DOLU"
End If
Next sayac
rs.Close

'4.kayýt
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If rs!OdaNo <> 4 Then
Label32.Caption = "Oda durumu=BOÞ"
rs.MoveNext
Else
Label32.Caption = "Oda durumu=DOLU"
End If
Next sayac
rs.Close

'5.kayýt
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If rs!OdaNo <> 5 Then
Label51.Caption = "Oda durumu=BOÞ"
rs.MoveNext
Else
Label51.Caption = "Oda durumu=DOLU"
End If
Next sayac
rs.Close

'6.kayýt
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If rs!OdaNo <> 6 Then
Label52.Caption = "Oda durumu=BOÞ"
rs.MoveNext
Else
Label52.Caption = "Oda durumu=DOLU"
End If
Next sayac
rs.Close

'7.kayýt
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If rs!OdaNo <> 7 Then
Label71.Caption = "Oda durumu=BOÞ"
rs.MoveNext
Else
Label71.Caption = "Oda durumu=DOLU"
End If
Next sayac
rs.Close

'8.kayýt
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If rs!OdaNo <> 8 Then
Label72.Caption = "Oda durumu=BOÞ"
rs.MoveNext
Else
Label72.Caption = "Oda durumu=DOLU"
End If
Next sayac
rs.Close

'9.kayýt
Set db = OpenDatabase(App.Path & "\þifre.mdb")
Set rs = db.OpenRecordset("tablo1")
a = rs.RecordCount
For sayac = 0 To a
If rs!OdaNo <> 9 Then
Label91.Caption = "Oda durumu=BOÞ"
rs.MoveNext
Else
Label91.Caption = "Oda durumu=DOLU"
End If
Next sayac
rs.Close


End Sub

Private Sub Timer1_Timer()
Picture1.Visible = True
Frame1.Visible = True
Timer1.Enabled = False
Timer2.Enabled = True
If Picture2.Visible = True Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = True
Timer1.Interval = 10000
Timer2.Interval = 10000
Timer1.Enabled = False
Timer2.Enabled = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
End If
If Picture4.Visible = True Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = True
Timer1.Interval = 10000
Timer2.Interval = 10000
Timer1.Enabled = False
Timer2.Enabled = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = True
End If
If Picture6.Visible = True Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = True
Timer1.Interval = 10000
Timer2.Interval = 10000
Timer1.Enabled = False
Timer2.Enabled = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = True
End If
If Picture8.Visible = True Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = True
Timer1.Interval = 10000
Timer2.Interval = 10000
Timer1.Enabled = False
Timer2.Enabled = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame8.Visible = False
Frame9.Visible = True
End If
End Sub
Private Sub Timer2_Timer()
Timer2.Interval = 10000
If Picture1.Visible = True Then
Picture1.Visible = False
Picture2.Visible = True
Timer1.Interval = 10000
Timer2.Interval = 10000
Timer1.Enabled = True
Timer2.Enabled = False
Frame1.Visible = False
Frame2.Visible = True
End If
If Picture3.Visible = True Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = True
Timer1.Interval = 10000
Timer1.Enabled = True
Timer2.Enabled = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = True
End If
If Picture5.Visible = True Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = True
Timer1.Interval = 10000
Timer1.Enabled = True
Timer2.Enabled = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = True
End If
If Picture7.Visible = True Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = True
Timer1.Interval = 10000
Timer1.Enabled = True
Timer2.Enabled = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame8.Visible = True
End If
If Picture9.Visible = True Then
Picture1.Visible = True
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = False
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame8.Visible = False
Frame9.Visible = False
End If
End Sub


