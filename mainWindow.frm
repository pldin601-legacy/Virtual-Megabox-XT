VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form wndMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00004040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5790
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "mainWindow.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5790
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picClock 
      BackColor       =   &H00000000&
      Height          =   675
      Left            =   8100
      ScaleHeight     =   615
      ScaleWidth      =   1875
      TabIndex        =   52
      Top             =   4680
      Width           =   1935
      Begin PicClip.PictureClip TSR 
         Left            =   60
         Top             =   720
         _ExtentX        =   4233
         _ExtentY        =   2064
         _Version        =   393216
         Rows            =   2
         Cols            =   5
         Picture         =   "mainWindow.frx":030A
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   555
         Left            =   840
         TabIndex        =   53
         Top             =   0
         Width           =   135
      End
      Begin VB.Image clckItem 
         Height          =   465
         Index           =   3
         Left            =   1500
         Picture         =   "mainWindow.frx":959C
         Stretch         =   -1  'True
         Top             =   60
         Width           =   285
      End
      Begin VB.Image clckItem 
         Height          =   465
         Index           =   2
         Left            =   1080
         Picture         =   "mainWindow.frx":A47E
         Stretch         =   -1  'True
         Top             =   60
         Width           =   300
      End
      Begin VB.Image clckItem 
         Height          =   465
         Index           =   1
         Left            =   480
         Picture         =   "mainWindow.frx":B360
         Stretch         =   -1  'True
         Top             =   60
         Width           =   285
      End
      Begin VB.Image clckItem 
         Height          =   465
         Index           =   0
         Left            =   60
         Picture         =   "mainWindow.frx":C242
         Stretch         =   -1  'True
         Top             =   60
         Width           =   285
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   540
      ScaleHeight     =   795
      ScaleWidth      =   3675
      TabIndex        =   16
      Top             =   5940
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Timer gStart 
         Interval        =   1000
         Left            =   2280
         Top             =   0
      End
      Begin VB.Timer gStop 
         Interval        =   1000
         Left            =   2700
         Top             =   0
      End
      Begin VB.Timer RealUpdate 
         Interval        =   250
         Left            =   1680
         Top             =   0
      End
      Begin MSComDlg.CommonDialog Desc2000 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "All Files (*.*) |*.*"
      End
      Begin vmbxt.cSysTray cSysTray 
         Left            =   1080
         Top             =   0
         _ExtentX        =   900
         _ExtentY        =   900
         InTray          =   0   'False
         TrayIcon        =   "mainWindow.frx":D124
         TrayTip         =   ""
      End
      Begin MSComDlg.CommonDialog dlgOpenSave 
         Left            =   540
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   "MXT"
         Filter          =   "Virtual MegaBox XT 4.01  (*.mxt) |*.mxt"
      End
      Begin MCI.MMControl MMHeader 
         Height          =   330
         Left            =   0
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   582
         _Version        =   393216
         UpdateInterval  =   250
         DeviceType      =   ""
         FileName        =   ""
      End
   End
   Begin VB.PictureBox SSPanel1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004040&
      Height          =   2355
      Left            =   60
      ScaleHeight     =   2295
      ScaleWidth      =   10155
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   10215
      Begin VB.DirListBox Dir1 
         Height          =   315
         Left            =   120
         TabIndex        =   54
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin PicClip.PictureClip picRotate 
         Left            =   720
         Top             =   720
         _ExtentX        =   1667
         _ExtentY        =   1667
         _Version        =   393216
         Rows            =   3
         Cols            =   3
         Picture         =   "mainWindow.frx":D27E
      End
      Begin vmbxt.QSImgButton btnSC 
         Height          =   315
         Left            =   4980
         Top             =   1920
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         Picture         =   "mainWindow.frx":10210
         BackColor       =   0
      End
      Begin PicClip.PictureClip PCDiR 
         Left            =   9420
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   635
         _Version        =   393216
         Rows            =   3
         Picture         =   "mainWindow.frx":10502
      End
      Begin VB.PictureBox Scroller 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   1800
         ScaleHeight     =   0.692
         ScaleMode       =   0  'User
         ScaleWidth      =   150
         TabIndex        =   37
         Top             =   1680
         Width           =   3735
      End
      Begin VB.PictureBox scrTotal 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1800
         ScaleHeight     =   195
         ScaleWidth      =   3735
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1380
         Width           =   3735
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   1235
         Left            =   5640
         ScaleHeight     =   1170
         ScaleWidth      =   4335
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   60
         Width           =   4395
         Begin VB.PictureBox pDirect 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   120
            Left            =   3660
            Picture         =   "mainWindow.frx":10E54
            ScaleHeight     =   120
            ScaleWidth      =   480
            TabIndex        =   48
            Top             =   840
            Width           =   480
         End
         Begin vmbxt.MTimer tmLen 
            Height          =   330
            Left            =   1260
            Top             =   240
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin vmbxt.MTimer tmPos 
            Height          =   330
            Left            =   120
            Top             =   240
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin vmbxt.MTrack tmtLen 
            Height          =   315
            Left            =   1260
            TabIndex        =   21
            Top             =   780
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
         End
         Begin vmbxt.MTrack tmtPos 
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   780
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
         End
         Begin VB.Label AST 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AUTOSTOP ON"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   165
            Left            =   2400
            TabIndex        =   46
            Top             =   300
            Width           =   1020
         End
         Begin VB.Label lblLedDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DIRECTION"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   120
            Index           =   5
            Left            =   3600
            TabIndex        =   30
            Top             =   660
            Width           =   555
         End
         Begin VB.Label lblLedDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TRACK"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   120
            Index           =   4
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lblLedDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL TRACKS"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   120
            Index           =   3
            Left            =   1260
            TabIndex        =   28
            Top             =   600
            Width           =   780
         End
         Begin VB.Label lblLedDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LENGTH"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   120
            Index           =   2
            Left            =   1260
            TabIndex        =   27
            Top             =   60
            Width           =   405
         End
         Begin VB.Label lblLedDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "POSITION"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   120
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   60
            Width           =   495
         End
         Begin VB.Image Model 
            Height          =   405
            Left            =   3660
            Picture         =   "mainWindow.frx":11196
            Top             =   120
            Width           =   465
         End
         Begin VB.Label ioREPEAT 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REPEAT"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   165
            Left            =   2580
            TabIndex        =   25
            Top             =   660
            Width           =   585
         End
         Begin VB.Line REPLine 
            BorderColor     =   &H0080FFFF&
            Index           =   0
            X1              =   3360
            X2              =   2400
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line REPLine 
            BorderColor     =   &H0080FFFF&
            Index           =   1
            X1              =   2400
            X2              =   2400
            Y1              =   600
            Y2              =   960
         End
         Begin VB.Line REPLine 
            BorderColor     =   &H0080FFFF&
            Index           =   4
            X1              =   3360
            X2              =   3360
            Y1              =   600
            Y2              =   960
         End
         Begin VB.Line REPLine 
            BorderColor     =   &H0080FFFF&
            Index           =   2
            X1              =   2400
            X2              =   2520
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line REPLine 
            BorderColor     =   &H0080FFFF&
            Index           =   5
            X1              =   3360
            X2              =   3240
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label ioALL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ALL"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   165
            Left            =   2580
            TabIndex        =   24
            Top             =   900
            Width           =   255
         End
         Begin VB.Label ioTRK 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TRK"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   165
            Left            =   2880
            TabIndex        =   23
            Top             =   900
            Width           =   285
         End
      End
      Begin VB.PictureBox VUPANEL 
         BackColor       =   &H00000000&
         Height          =   1235
         Left            =   1800
         ScaleHeight     =   1170
         ScaleWidth      =   3675
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   60
         Width           =   3735
         Begin VB.PictureBox picRots 
            AutoRedraw      =   -1  'True
            BackColor       =   &H0000FFFF&
            DrawWidth       =   5
            Height          =   255
            Left            =   1260
            ScaleHeight     =   195
            ScaleWidth      =   675
            TabIndex        =   51
            Top             =   420
            Width           =   735
         End
         Begin VB.PictureBox ledCont 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   75
            Left            =   3420
            ScaleHeight     =   75
            ScaleWidth      =   135
            TabIndex        =   49
            ToolTipText     =   "PLAYLIST ACTIVITY"
            Top             =   360
            Width           =   135
         End
         Begin VB.PictureBox picEnd 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   75
            Left            =   3420
            ScaleHeight     =   75
            ScaleWidth      =   135
            TabIndex        =   45
            Top             =   240
            Width           =   135
         End
         Begin VB.PictureBox F_LED 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   75
            Left            =   3420
            ScaleHeight     =   75
            ScaleWidth      =   135
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   120
            Width           =   135
         End
         Begin VB.Image imgCas2 
            Height          =   315
            Left            =   2030
            Picture         =   "mainWindow.frx":11BF8
            Top             =   390
            Width           =   315
         End
         Begin VB.Image imgCas1 
            Height          =   315
            Left            =   900
            Picture         =   "mainWindow.frx":1217A
            Top             =   390
            Width           =   315
         End
         Begin VB.Image Image1 
            Height          =   1095
            Left            =   480
            Picture         =   "mainWindow.frx":126FC
            Top             =   25
            Width           =   2280
         End
      End
      Begin vmbxt.QSButton btnExplore 
         Height          =   255
         Left            =   9540
         Top             =   1980
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
         BackColor       =   0
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65535
      End
      Begin vmbxt.QSButton btnCopy 
         Height          =   255
         Left            =   7260
         Top             =   1980
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         BackColor       =   0
         Caption         =   "COPY FILES..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   12632064
      End
      Begin vmbxt.QSButton btnKill 
         Height          =   255
         Left            =   5640
         ToolTipText     =   "Kill"
         Top             =   1980
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   450
         BackColor       =   0
         Caption         =   "KILL FILE..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   255
      End
      Begin vmbxt.QSButton btnVolume 
         Height          =   315
         Left            =   8700
         Top             =   1680
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BackColor       =   0
         Caption         =   "MIXER"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65535
      End
      Begin vmbxt.QSButton btnMode 
         Height          =   315
         Left            =   6900
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BackColor       =   0
         Caption         =   "DIRECTION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65535
      End
      Begin vmbxt.QSButton btnRepeat 
         Height          =   315
         Left            =   5640
         Top             =   1680
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BackColor       =   0
         Caption         =   "REPEAT"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65535
      End
      Begin vmbxt.QSImgButton btnStop 
         Height          =   315
         Left            =   4080
         Top             =   1920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Picture         =   "mainWindow.frx":1A946
         BackColor       =   0
      End
      Begin vmbxt.QSImgButton btnPause 
         Height          =   315
         Left            =   3540
         Top             =   1920
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Picture         =   "mainWindow.frx":1AC38
         BackColor       =   0
      End
      Begin vmbxt.QSImgButton btnNext 
         Height          =   315
         Left            =   2940
         Top             =   1920
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         Picture         =   "mainWindow.frx":1AE82
         BackColor       =   0
      End
      Begin vmbxt.QSImgButton btnPlay 
         Height          =   315
         Left            =   2400
         Top             =   1920
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Picture         =   "mainWindow.frx":1B1E0
         BackColor       =   0
      End
      Begin vmbxt.QSImgButton btnPrev 
         Height          =   315
         Left            =   1800
         ToolTipText     =   "Prev"
         Top             =   1920
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         Picture         =   "mainWindow.frx":1B43A
         BackColor       =   0
      End
      Begin vmbxt.QSButton btnWindow 
         Height          =   255
         Left            =   120
         Top             =   1680
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   450
         BackColor       =   0
         Caption         =   "BIG WINDOW"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65535
      End
      Begin vmbxt.QSButton Command1 
         Height          =   255
         Left            =   900
         Top             =   1980
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BackColor       =   0
         Caption         =   "SETUP"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65280
      End
      Begin vmbxt.QSButton btnExit 
         Height          =   255
         Left            =   120
         Top             =   1980
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BackColor       =   192
         Caption         =   "OFF"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   12632319
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MegaBOX"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Index           =   1
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Virtual "
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RENNSoft Virtual MegaBox XT 4.04"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   5640
         TabIndex        =   33
         Top             =   1380
         Width           =   4335
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   15
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label lblTitleBack 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RENNSoft Virtual MegaBox XT 4.04"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5670
         TabIndex        =   34
         Top             =   1410
         Width           =   4335
      End
   End
   Begin vmbxt.QSButton btnHide 
      Height          =   255
      Left            =   9720
      Top             =   5460
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   450
      BackColor       =   4210752
      Caption         =   "--->"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8,25
   End
   Begin VB.PictureBox picCPU 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7800
      ScaleHeight     =   255
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   5460
      Width           =   1875
   End
   Begin VB.PictureBox tblTimer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004040&
      Height          =   1155
      Left            =   60
      ScaleHeight     =   1095
      ScaleWidth      =   7635
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4560
      Width           =   7695
      Begin VB.PictureBox picCover 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   0
         Picture         =   "mainWindow.frx":1B80C
         ScaleHeight     =   1095
         ScaleWidth      =   7635
         TabIndex        =   50
         Top             =   0
         Width           =   7635
      End
      Begin VB.OptionButton opDone 
         BackColor       =   &H00000000&
         Caption         =   "DONE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   5940
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   660
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton opStop 
         BackColor       =   &H00000000&
         Caption         =   "STOP TIME"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   5940
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   420
         Width           =   1335
      End
      Begin VB.OptionButton opStart 
         BackColor       =   &H00000000&
         Caption         =   "START TIME"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   5940
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   180
         Width           =   1335
      End
      Begin vmbxt.QSButton SetTimeBut 
         Height          =   195
         Index           =   0
         Left            =   5100
         Top             =   720
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   0
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65280
      End
      Begin VB.CheckBox btnTimer 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   660
         Width           =   915
      End
      Begin VB.PictureBox tblTmrIndicator 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   855
         Left            =   1320
         ScaleHeight     =   795
         ScaleWidth      =   3615
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   3675
         Begin VB.PictureBox ledStop 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   75
            Left            =   3180
            ScaleHeight     =   75
            ScaleWidth      =   255
            TabIndex        =   42
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox ledStart 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   75
            Left            =   120
            ScaleHeight     =   75
            ScaleWidth      =   255
            TabIndex        =   41
            Top             =   240
            Width           =   255
         End
         Begin vmbxt.MTimer tmrStop 
            Height          =   330
            Left            =   2580
            Top             =   360
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin vmbxt.MTimer tmrStart 
            Height          =   330
            Left            =   120
            Top             =   360
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin VB.Label lblPro1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Timer programming:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   1155
            TabIndex        =   7
            Top             =   60
            Width           =   1275
         End
      End
      Begin vmbxt.QSButton SetTimeBut 
         Height          =   195
         Index           =   1
         Left            =   5100
         Top             =   540
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   0
         Caption         =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65280
      End
      Begin vmbxt.QSButton SetTimeBut 
         Height          =   195
         Index           =   2
         Left            =   5280
         Top             =   540
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   0
         Caption         =   "2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65280
      End
      Begin vmbxt.QSButton SetTimeBut 
         Height          =   195
         Index           =   3
         Left            =   5460
         Top             =   540
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   0
         Caption         =   "3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65280
      End
      Begin vmbxt.QSButton SetTimeBut 
         Height          =   195
         Index           =   4
         Left            =   5100
         Top             =   360
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   0
         Caption         =   "4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65280
      End
      Begin vmbxt.QSButton SetTimeBut 
         Height          =   195
         Index           =   5
         Left            =   5280
         Top             =   360
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   0
         Caption         =   "5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65280
      End
      Begin vmbxt.QSButton SetTimeBut 
         Height          =   195
         Index           =   6
         Left            =   5460
         Top             =   360
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   0
         Caption         =   "6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65280
      End
      Begin vmbxt.QSButton SetTimeBut 
         Height          =   195
         Index           =   7
         Left            =   5100
         Top             =   180
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   0
         Caption         =   "7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65280
      End
      Begin vmbxt.QSButton SetTimeBut 
         Height          =   195
         Index           =   8
         Left            =   5280
         Top             =   180
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   0
         Caption         =   "8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65280
      End
      Begin vmbxt.QSButton SetTimeBut 
         Height          =   195
         Index           =   9
         Left            =   5460
         Top             =   180
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   0
         Caption         =   "9"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65280
      End
      Begin vmbxt.QSButton SetTimeBut 
         Height          =   195
         Index           =   11
         Left            =   5280
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         BackColor       =   0
         Caption         =   "-"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65280
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Timer On/Off"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   180
         TabIndex        =   11
         Top             =   120
         Width           =   930
      End
   End
   Begin VB.PictureBox tblListCtrl 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004040&
      Height          =   2055
      Left            =   60
      ScaleHeight     =   1995
      ScaleWidth      =   10170
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2460
      Width           =   10230
      Begin VB.PictureBox picPos 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1875
         Left            =   1860
         ScaleHeight     =   1815
         ScaleWidth      =   195
         TabIndex        =   57
         Top             =   60
         Width           =   255
         Begin VB.Label lbPlayable 
            BackStyle       =   0  'Transparent
            Caption         =   ">>"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   195
         End
      End
      Begin VB.ListBox lstTimes 
         Height          =   1620
         Left            =   5640
         TabIndex        =   56
         Top             =   180
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MCI.MMControl TimeX 
         Height          =   330
         Left            =   840
         TabIndex        =   55
         Top             =   1260
         Visible         =   0   'False
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   582
         _Version        =   393216
         PrevVisible     =   0   'False
         NextVisible     =   0   'False
         PlayVisible     =   0   'False
         PauseVisible    =   0   'False
         StepVisible     =   0   'False
         RecordVisible   =   0   'False
         EjectVisible    =   0   'False
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.Timer Timer1 
         Interval        =   125
         Left            =   120
         Top             =   1500
      End
      Begin VB.TextBox inTitle 
         Height          =   285
         Left            =   8040
         TabIndex        =   47
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin vmbxt.QSButton btnPlF 
         Height          =   255
         Left            =   8040
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         BackColor       =   65280
         Caption         =   "Open files..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   0
      End
      Begin vmbxt.QSButton btnPLMnu 
         Height          =   255
         Left            =   8040
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         BackColor       =   8388608
         Caption         =   "PLAYLIST MENU"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65535
      End
      Begin VB.ListBox lstSec 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1875
         IntegralHeight  =   0   'False
         Left            =   2100
         MultiSelect     =   2  'Extended
         OLEDropMode     =   1  'Manual
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   60
         Width           =   5475
      End
      Begin VB.CommandButton gDown 
         BackColor       =   &H00008080&
         Height          =   975
         Left            =   7560
         MaskColor       =   &H00008080&
         Picture         =   "mainWindow.frx":36C06
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Width           =   435
      End
      Begin VB.CommandButton gUp 
         BackColor       =   &H00008080&
         Height          =   915
         Left            =   7560
         MaskColor       =   &H00008080&
         Picture         =   "mainWindow.frx":37048
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   60
         Width           =   435
      End
      Begin VB.CommandButton btnLN 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "LN"
         Top             =   1020
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ListBox lstMain 
         BackColor       =   &H00808080&
         Columns         =   1
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         ItemData        =   "mainWindow.frx":3748A
         Left            =   120
         List            =   "mainWindow.frx":3748C
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         Style           =   1  'Checkbox
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1020
         Visible         =   0   'False
         Width           =   1575
      End
      Begin vmbxt.QSButton btnRecPL 
         Height          =   255
         Left            =   8040
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         BackColor       =   16384
         Caption         =   "RECENT PLAYLIST"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65535
      End
      Begin vmbxt.QSButton btnFav 
         Height          =   255
         Left            =   8040
         Top             =   780
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         BackColor       =   16384
         Caption         =   "FAVOURITES"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6,75
         ForeColor       =   65535
      End
      Begin VB.Label lblZag 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   1005
         TabIndex        =   44
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblMask 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Loading..."
         ForeColor       =   &H0000FF00&
         Height          =   1875
         Left            =   2100
         TabIndex        =   43
         Top             =   60
         Width           =   5535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MegaLIST"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Virtual"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   60
         Width           =   1140
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
   End
   Begin PicClip.PictureClip Modes 
      Left            =   7380
      Top             =   180
      _ExtentX        =   820
      _ExtentY        =   7144
      _Version        =   393216
      Rows            =   10
      Picture         =   "mainWindow.frx":3748E
   End
   Begin VB.Menu mSetx 
      Caption         =   "SETx"
      Visible         =   0   'False
      Begin VB.Menu mSLA 
         Caption         =   "Plugins"
         Begin VB.Menu stngr 
            Caption         =   "1. Stinger"
         End
         Begin VB.Menu idtaged 
            Caption         =   "2. Launch ""idTAG Editor"""
         End
         Begin VB.Menu renameFile 
            Caption         =   "3. Rename File"
         End
         Begin VB.Menu htmlply 
            Caption         =   "4. Generate HTML Playlist"
         End
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResetSet 
         Caption         =   "Reset Settings"
      End
      Begin VB.Menu mnuSaveSet 
         Caption         =   "Save Settings"
      End
      Begin VB.Menu ss1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRep 
         Caption         =   "Repeat Mode"
         Begin VB.Menu mnuRepS 
            Caption         =   "No Repeat"
            Index           =   0
         End
         Begin VB.Menu mnuRepS 
            Caption         =   "List Repeat"
            Index           =   1
         End
         Begin VB.Menu mnuRepS 
            Caption         =   "Track Repeat"
            Index           =   2
         End
      End
      Begin VB.Menu mnuPD 
         Caption         =   "Play Direction"
         Begin VB.Menu mnuDirF 
            Caption         =   "Forward"
            Index           =   0
         End
         Begin VB.Menu mnuDirF 
            Caption         =   "Backward"
            Index           =   1
         End
         Begin VB.Menu mnuDirF 
            Caption         =   "Playlist Off"
            Index           =   2
         End
      End
      Begin VB.Menu mnuPSS 
         Caption         =   "Playlist Saving Settings"
         Begin VB.Menu mnuFS 
            Caption         =   "File Selection"
         End
         Begin VB.Menu mnuPP 
            Caption         =   "Play Position"
         End
         Begin VB.Menu mnuPlM 
            Caption         =   "Play Mode"
         End
         Begin VB.Menu mnuTS 
            Caption         =   "Timer Settings"
         End
         Begin VB.Menu mnuCP 
            Caption         =   "Cursor Position"
         End
      End
      Begin VB.Menu mnuLF 
         Caption         =   "TAG formatting"
         Begin VB.Menu mnuTAG 
            Caption         =   "TAG format :"
         End
         Begin VB.Menu mnuTAGLESS 
            Caption         =   "TAGless format :"
         End
         Begin VB.Menu ss2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuprotime 
            Caption         =   "Proportional time [0]"
         End
         Begin VB.Menu mnuNoTags 
            Caption         =   "No TAGs"
         End
         Begin VB.Menu mnuTupd 
            Caption         =   "Update Tags"
         End
      End
      Begin VB.Menu mnuPCM 
         Caption         =   "Playing completed mode"
         Begin VB.Menu mnuSTO 
            Caption         =   "Stop only"
            Index           =   0
         End
         Begin VB.Menu mnuSTO 
            Caption         =   "Exit player"
            Index           =   1
         End
         Begin VB.Menu mnuSTO 
            Caption         =   "Exit && stanby computer"
            Index           =   2
         End
      End
      Begin VB.Menu ss4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPau 
         Caption         =   "Pause between tracks: "
      End
      Begin VB.Menu mnuFLF 
         Caption         =   "Find lost files"
      End
   End
   Begin VB.Menu mnuPM 
      Caption         =   "Playlist Menu"
      Visible         =   0   'False
      Begin VB.Menu btnAdd 
         Caption         =   "Add file..."
      End
      Begin VB.Menu btnDelete 
         Caption         =   "Remove file"
      End
      Begin VB.Menu mnusepA0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChPt 
         Caption         =   "Change path..."
      End
      Begin VB.Menu sepd 
         Caption         =   "-"
      End
      Begin VB.Menu btnload 
         Caption         =   "Load playlist..."
      End
      Begin VB.Menu btnSave 
         Caption         =   "Save playlist..."
      End
      Begin VB.Menu sp8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRc 
         Caption         =   "Recent"
         Begin VB.Menu mnuLoadRec 
            Caption         =   "Load recent playlist..."
         End
         Begin VB.Menu mnuSaveRec 
            Caption         =   "Save recent playlist..."
         End
         Begin VB.Menu lstrecsep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecLoc 
            Caption         =   "Recent playlists location..."
         End
         Begin VB.Menu lstrecsep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecQ 
            Caption         =   "Quick Load"
            Begin VB.Menu mnuRecItm 
               Caption         =   "<no items>"
               Enabled         =   0   'False
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnusep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFavour 
         Caption         =   "Favourites"
         Begin VB.Menu mnuAddFav 
            Caption         =   "Add item to favourites"
         End
         Begin VB.Menu mnuAddAllFav 
            Caption         =   "Add all list entries to favour."
         End
         Begin VB.Menu mnuSeo7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuItm 
            Caption         =   "< no items >"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuSep8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAddAll 
            Caption         =   "Add all into playlist"
         End
         Begin VB.Menu mnuSep6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDelFav 
            Caption         =   "Open editor..."
         End
      End
      Begin VB.Menu mnusepA1 
         Caption         =   "-"
      End
      Begin VB.Menu btnFind 
         Caption         =   "Find item..."
      End
      Begin VB.Menu mnusepA2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSP 
         Caption         =   "Sort playlist"
         Begin VB.Menu btnRndSort 
            Caption         =   "Shuffle"
         End
      End
      Begin VB.Menu mnusepA4 
         Caption         =   "-"
      End
      Begin VB.Menu btnSelect 
         Caption         =   "Select all"
      End
      Begin VB.Menu btnUnSelect 
         Caption         =   "Unselect all"
      End
      Begin VB.Menu btnClear 
         Caption         =   "Clear all"
      End
   End
End
Attribute VB_Name = "wndMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TimeCount As Integer
Dim OldX, OldY
Dim TimerState As String
Dim Tik
Dim Fucka As String, LN As Integer
Dim Ench As Integer
Dim Mozhna As Boolean
Dim UFlag As Integer
Dim DefFile As String
Dim Quanta As Integer

Dim aSecs, aMins, aSecls, bSecs, bMins, bSecls

Dim StartTime As String
Dim StopTime As String

Dim LockCont As Boolean

Dim CanMove As Boolean
Dim Missed As Boolean

Dim PD As Long
Dim CT, TT, PT As Currency
Dim Rot1, Rot2 As Currency

Dim BackIndex As Integer

Function FindPlayDeal() As Integer

  On Error Resume Next
  FindPlayDeal = -1
  
  For K = 0 To lstSec.ListCount - 1
    If lstMain.List(K) = MMHeader.Filename Then FindPlayDeal = K: Exit For
  Next K

End Function

Sub PlayablePos()
 lbPlayable.Top = (BackIndex - lstSec.TopIndex) * 210
End Sub

Sub UpdateSetup()
On Error Resume Next
Dim Y, MY

' Play direction
mnuDirF(0).Checked = False
mnuDirF(1).Checked = False
mnuDirF(2).Checked = False
mnuDirF(GlbPlayMode).Checked = True

' Repeat mode
mnuRepS(0).Checked = False
mnuRepS(1).Checked = False
mnuRepS(2).Checked = False
mnuRepS(GlbRepeat).Checked = True

' Stop mode
mnuSTO(0).Checked = False
mnuSTO(1).Checked = False
mnuSTO(2).Checked = False
mnuSTO(GlbSTM).Checked = True


' Title
' cmdScroll.Value = GlbWndScroll
' cmdFluent.Value = GlbWndFluent
' lbTtlClr.ForeColor = GlbWndColor

mnuFS.Checked = CBool(GlbCHFS)
mnuPP.Checked = CBool(GlbCHPP)
mnuPlM.Checked = CBool(GlbCHPM)
mnuTS.Checked = CBool(GlbCHTS)
mnuCP.Checked = CBool(GlbCHCP)

mnuTAG.Caption = "TAG format: " + GlbTitle
mnuTAGLESS.Caption = "TAGless format: " + GlbTagLess
mnuNoTags.Checked = GlbNoTags
mnuprotime.Caption = "Proportional time: [" + Format(GlbPropTime, "0") + " bytes ]"


mnuPau.Caption = "Pause between tracks: " + Format(GlbPause, "0 ms")
mnuFLF.Checked = GlbSeeker

End Sub

Sub ApplySettings()
Setting.GlbLoadSettings
SelectPlay GlbPlayMode
SelectMode GlbRepeat
UpdateSetup
End Sub

Function ScanFilePath(FldPath As String, FlFind As String, dirList As DirListBox)

Dim Z As String, M As String, N As String
On Error Resume Next
dirList.Path = LowPath(FldPath)
If Err Then Exit Function
M = Dir(LowPath(FldPath) + FlFind)
If M > "" Then
   ScanFilePath = LowPath(FldPath) + M
   Exit Function
End If

For IDN = 0 To Dir1.ListCount - 1
  
 Z = dirList.List(IDN)

 N = ScanFilePath(Z, FlFind, dirList)
 If N > "" Then ScanFilePath = N: Exit Function
 dirList.Path = LowPath(FldPath)

Next IDN



End Function

Function CCPU() As Currency

CCPU = 0

End Function

Sub CenterList()

If BackIndex >= 4 Then lstSec.TopIndex = BackIndex - 4
If BackIndex < 4 Then lstSec.TopIndex = 0

End Sub
Sub CenterList2()

If lstSec.ListIndex >= 4 Then lstSec.TopIndex = lstSec.ListIndex - 4
If lstSec.ListIndex < 4 Then lstSec.TopIndex = 0

End Sub


Sub Command_Next()
On Error Resume Next

If BackIndex < lstMain.ListCount - 1 Then
 ResetStatus
 BackIndex = BackIndex + 1
 ' lstSec.ListIndex = lstSec.ListIndex + 1
 Command_Open
 Command_Play
End If


End Sub

Sub Command_Pause()
MMHeader.Command = "Pause"
End Sub

Sub Command_Play()
MMHeader.Command = "Play"

End Sub

Sub Command_Open()
On Error Resume Next

Dim T, LFILE As String, LFPATH As String, LFNAME As String, NN As Integer

 If BackIndex >= 0 Then
  
  If BackIndex + 1 > lstSec.ListCount Then BackIndex = 0
  
  Scroller.Visible = True

  F_LED.BackColor = RGB(255, 255, 0)
  
  LFILE = lstMain.List(BackIndex)
  LFPATH = PathHead(lstMain.List(BackIndex))
  LFNAME = FileHead(lstMain.List(BackIndex))
  CenterList
  Dim ZN
  
  If FileExists(LFILE) = False And GlbSeeker = 1 Then
   lblTitle.Caption = "File not found. Scanning disk..."
   lblTitleBack.Caption = lblTitle.Caption
   DoEvents
   Do
    ZN = ScanFilePath(LFPATH, LFNAME, Dir1)
    If ZN > "" Then
      LFILE = ZN
      lstMain.List(lstSec.ListIndex) = LFILE
      lstSec.List(lstSec.ListIndex) = GetMp3Song(LFILE, NN)
      lstTimes.List(lstSec.ListIndex) = Str(NN)
      Exit Do
    End If
    LFPATH = PathHead(LFPATH)
   Loop While LFPATH <> ""
  End If

  
  MMHeader.Command = "Close"
  MMHeader.Filename = LFILE
  
  CT = lstSec.ListIndex + 1
  TT = lstSec.ListCount
  lblTitle.Caption = lstSec.List(BackIndex)
  lblTitleBack.Caption = lblTitle.Caption
  
  wndTitle.lab(0).Caption = lblTitle.Caption + "  "

  T = Timer * 1000
  MMHeader.Command = "Open"
  
  PD = (Timer * 1000) - T
  PicCPUSub
  PlayablePos

  tmtLen.Track = lstMain.ListCount
  tmtPos.Track = BackIndex + 1
  
  
  If BackIndex = lstMain.ListCount - 1 Then
    tmtPos.LastRec
  End If
  
  If Err = 0 Then MMHeader_StatusUpdate
  
  If MMHeader.Error Then
    tmtPos.ErrorFound
    lblTitle.Caption = MMHeader.ErrorMessage
    F_LED.BackColor = RGB(255, 0, 0)
  End If
  
  F_LED.BackColor = RGB(0, 0, 0)
    
 End If

End Sub


Sub Command_Prev()
On Error Resume Next

If BackIndex > 0 Then
 ResetStatus
 BackIndex = BackIndex - 1
 ' lstSec.ListIndex = lstSec.ListIndex - 1
 Command_Open
 Command_Play
End If


End Sub

Sub Command_Stop()
 MMHeader.Command = "Stop"
 MMHeader.To = 0
 MMHeader.Command = "Seek"
End Sub

Sub DrawText(TextFile As String)
On Error Resume Next

I = FreeFile

Open TextFile For Input As #I
 txtFileData.TextRTF = ""
 txtFileData.TextRTF = Input$(LOF(I), I)
 If Err = 0 Then lblPadName.Caption = "Virtual MegaBox media info reader - " + FileHead(TextFile)
 If Err Then lblPadName.Caption = "Virtual MegaBox media info reader - text not found"
Close #I

Err.Clear

End Sub

Sub Facing()
Rem 'Lang32 Loading'
lblVersion(0).Caption = "Version XT " + GetVersion
lblVersion(1).Caption = "Version XT " + GetVersion
End Sub

Sub LC()
If Setting.GlbPlayMode = 2 Then
  ledCont.BackColor = RGB(0, 0, 0)
Else
  ledCont.BackColor = RGB(0, 0, 255)
End If
End Sub

Sub LoadList(ByRef ListFileName As String, Destination As ListBox, MainForm As Form)
On Error Resume Next

pinfo.Show
pinfo.lblSesc.Caption = "Loading playlist, wait..."
DoEvents
' Definitions
Dim FN As String, Mired As Boolean, SP As String, Lines As Long, Chent As String
Dim ComaData As String, ComaValue As String, Misses As Boolean, AutoPlay As Boolean, SeekPlay As Currency
Dim LastCommand As String

' Load File
DoEvents
Open ListFileName For Input As #1
If Err Then MsgBox "Error! ''" + ListFileName + "'' has no access.", vbCritical: Exit Sub
F_LED.BackColor = RGB(0, 128, 255)

Destination.Clear
lstSec.Clear

' Read lines
Do


Line Input #1, FN
Lines = Lines + 1


' Skiping remark lines
If Left(FN, 1) = ";" Then Mired = True
If Left(FN, 1) = " " Then Mired = True
If FN = "" Then Mired = True

' Reading for commands
If Left(FN, 1) = "#" Then
  ComaData = ReadCommand(FN, False)
  ComaValue = ReadCommand(FN, True)
  If Err Then MsgBox "Error found!" + Chr$(13) + Chr$(13) + "Line " + Format$(Lines, "0"), vbInformation: Err = 0: Misses = True
  
  ' Command found, then
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#FILE     :" Then
     
     SP = ComaValue
     
     Destination.AddItem LowPath(PathHead(SP)) + FileHead(SP)
     Me.lstSec.AddItem LowPath(PathHead(SP)) + FileHead(SP): Err = 0
     
     Err.Clear
     
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#SELECT   :" Then
   Destination.Selected(Val(ComaValue)) = True
   Me.lstSec.Selected(Val(ComaValue)) = True
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''SELECT'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#PATH     :" Then
   Chent = ComaValue
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''PATH'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#FOCUS    :" Then
   If ComaValue >= 0 Then Me.lstSec.ListIndex = Val(ComaValue)
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''FOCUS'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#TIMER    :" Then
   btnTimer.Value = Val(ComaValue)
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''TIMER'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#START    :" Then
   StartTime = ComaValue
   tmrStart.TimeSet2 = Mid(ComaValue, 1, 2) + ":" + Mid(ComaValue, 3, 2)
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''START'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#STOP     :" Then
   StopTime = ComaValue
   tmrStop.TimeSet2 = Mid(ComaValue, 1, 2) + ":" + Mid(ComaValue, 3, 2)
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''STOP'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#MESSAGE  :" Then Call MsgBox(ComaValue, vbInformation)
  If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''MESSAGE'' is used incorrect!", vbInformation: Err = 0: Misses = True
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#DIRECTION:" Then
   SelectPlay Val(ComaValue)
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''DIRECTION'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#AUTOPLAY :" Then
   AutoPlay = True
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''AUTOPLAY'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#SEEK     :" Then
   SeekPlay = Val(ComaValue)
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''SEEK'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#DEFAULT  :" Then
   DefFile = ComaValue
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''DEFAULT'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  ' Reset Ifs
  ComaData = ""
  ComaValue = ""
  Mired = True
End If

Dim NN As Integer

If Mired = False Then
   Destination.AddItem FN
   Me.lstSec.AddItem GetMp3Song(FN, NN)
   Me.lstTimes.AddItem Str(NN)
   Destination.Selected(Destination.ListCount - 1) = True
   Me.lstSec.Selected(Destination.ListCount - 1) = True
End If

Mired = False
Loop While Not EOF(1)

F_LED.BackColor = RGB(0, 0, 0)
pinfo.Hide

If Misses Then MsgBox "Errors found during load playlist!", vbInformation

If Err = 0 Then Beep

Close #1

If Chent > "" Then On Error Resume Next: ChDir (Chent): ChDrive (Mid(Chent, 1, 2))

UpdateTags -1
lstSec_Click

If AutoPlay = True Then
  btnPlay_Click
  MMHeader.To = SeekPlay
  MMHeader.Command = "Seek"
  MMHeader.Command = "Play"
End If



End Sub


Sub LoadListAdd(ByRef ListFileName As String, Destination As ListBox, MainForm As Form)
On Error Resume Next

pinfo.Show
pinfo.lblSesc.Caption = "Loading playlist '" + FileHead(ListFileName) + "', wait..."
DoEvents
' Definitions
Dim FN As String, Mired As Boolean, SP As String, Lines As Long, Chent As String
Dim ComaData As String, ComaValue As String, Misses As Boolean, AutoPlay As Boolean, SeekPlay As Currency
Dim LastCommand As String
Dim xCount As Integer

' Load File
DoEvents

Open ListFileName For Input As #1
If Err Then MsgBox "Error! ''" + ListFileName + "'' has no access.", vbCritical: Exit Sub
F_LED.BackColor = RGB(0, 128, 255)

xCount = lstSec.ListCount

' Read lines
Do


Line Input #1, FN
Lines = Lines + 1


' Skiping remark lines
If Left(FN, 1) = ";" Then Mired = True
If Left(FN, 1) = " " Then Mired = True
If FN = "" Then Mired = True

' Reading for commands
If Left(FN, 1) = "#" Then
  ComaData = ReadCommand(FN, False)
  ComaValue = ReadCommand(FN, True)
  If Err Then MsgBox "Error found!" + Chr$(13) + Chr$(13) + "Line " + Format$(Lines, "0"), vbInformation: Err = 0: Misses = True
  
  ' Command found, then
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#FILE     :" Then
     
     SP = ComaValue
     
     Destination.AddItem LowPath(PathHead(SP)) + FileHead(SP)
     Me.lstSec.AddItem LowPath(PathHead(SP)) + FileHead(SP): Err = 0
     
     Err.Clear
     
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#SELECT   :" Then
   Destination.Selected(xCount + Val(ComaValue)) = True
   Me.lstSec.Selected(xCount + Val(ComaValue)) = True
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''SELECT'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#PATH     :" Then
   Chent = ComaValue
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''PATH'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#FOCUS    :" Then
   If ComaValue >= 0 Then Me.lstSec.ListIndex = xCount + Val(ComaValue)
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''FOCUS'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  
  ' Reset Ifs
  ComaData = ""
  ComaValue = ""
  Mired = True
End If

Dim NN As Integer

If Mired = False Then
  Destination.AddItem FN
  Me.lstSec.AddItem GetMp3Song(FN, NN)
  Me.lstTimes.AddItem Str(NN)
  Destination.Selected(Destination.ListCount - 1) = True
  Me.lstSec.Selected(Destination.ListCount - 1) = True
End If

Mired = False
Loop While Not EOF(1)

F_LED.BackColor = RGB(0, 0, 0)
pinfo.Hide

If Misses Then MsgBox "Errors found during load playlist!", vbInformation

If Err = 0 Then Beep

Close #1

lstSec_Click

End Sub



Sub PicCPUSub()
 picCPU.Cls
 
 picCPU.Scale (0, 0)-(1, 2)

 picCPU.FontName = "Small Fonts"
 picCPU.FontSize = "5"
 picCPU.Line (0, 0)-(1, 0.9), RGB(0, 0, 255), BF
 picCPU.Line (0, 1)-(1 / 1000 * PD, 1.9), RGB(255, 255, 0), BF
 
 picCPU.ForeColor = RGB(255, 255, 0)
 picCPU.CurrentX = 0: picCPU.CurrentY = 0
 picCPU.Print "VIRTUAL  MEGABOX  XT  " + GetVersion + "  " + Format(App.Revision) + "E"
 
 picCPU.ForeColor = RGB(0, 0, 255)
 picCPU.CurrentX = 0: picCPU.CurrentY = 1
 picCPU.Print "HTTP://WWW.RENNSoft.COM.UA"

End Sub

Sub ResetStatus()

 Me.tmtLen.Waiting
 Me.tmtPos.Waiting
 Me.tmLen.Off
 Me.tmPos.Off

 lblTitle.Caption = ""
 lblTitleBack.Caption = ""


 ScrChange 1, 0
 Scroller.Visible = False
 
 
 MMHeader.Command = "Close"
 MMHeader.Filename = ""

 TotalChange 1, 0
 F_LED.BackColor = RGB(0, 0, 0)
 picEnd.BackColor = RGB(0, 0, 0)

 If Month(Now) = 12 And 19 < Day(Now) <= 31 Then
   wndTitle.lab(0).Caption = "Release 27.04.2002-24.01.2003 *** RENNSoft Multimedia Studio *** Happy New Year! *** "
 Else
   wndTitle.lab(0).Caption = "Release 27.04.2002-24.01.2003 *** RENNSoft Multimedia Studio *** "
 End If
 
End Sub

Sub Rots(Value As Currency)
 
 picRots.Cls
 
 Dim T As Integer
 
 For T = 0 To 200 + (350 / Sqr(100) * Sqr(100 - Value)) Step 5
   picRots.Circle (-300, 90), T, RGB(0, 0, 0)
 Next T
 
 For T = 0 To 200 + (350 / Sqr(100) * Sqr(Value)) Step 5
   picRots.Circle (picRots.Width + 200, 90), T, RGB(0, 0, 0)
 Next T
 
End Sub

Function SaveList(ListFile As String, Source As ListBox) As Boolean
' Standart
On Error Resume Next
Dim I, J, Current As String

Current = CurDir

pinfo.Show
pinfo.lblSesc.Caption = "Saving playlist..."
DoEvents

I = FreeFile
Open ListFile For Output As I
If Err Then Close I: SaveList = False: Exit Function

F_LED.BackColor = RGB(255, 128, 0)
 ' Some Comments
 Print #I, " RENNSoft (R) Virtual MegaBox eXTended Version playlist file"
 Print #I, " Copyright (C) 2000-" + Format(Year(Now), "0000") + " RENNSoft SM Studio. All Rights reserved."
 Print #I, " File format: .MXT, Created " + Format$(Now, "dd.mm.yyyy") + " at " + Format$(Now, "hh:mm")
 Print #I, ""
 Print #I, "#PATH     : " + Current
 Print #I, "#DEFAULT  : " + ListFile
 
 
' Files for add
For J = 0 To Source.ListCount - 1

 Print #I, "#FILE     : " + Source.List(J)
 Print #I, "#TITLE    : " + Me.lstSec.List(J)
Next
  
  Print #I, ""
' Files for select
If GlbCHFS = 1 Then
  For J = 0 To Source.ListCount - 1
   If Source.Selected(J) = True Then
     Print #I, "#SELECT   : " + Str(J)
   End If
  Next
End If

 ' Miscs
 Print #I, ""
If GlbCHPP = 1 Then
 If MMHeader.Mode = 526 Then
  Print #I, "#AUTOPLAY : PRESENT"
  Print #I, "#SEEK     : " + Format(MMHeader.Position, "000 000 000 000") + " MSeconds"
 End If
End If

If GlbCHTS = 1 Then
 Print #I, "#TIMER    : " + Str$(CInt(btnTimer.Value))
 Print #I, "#START    : " + StartTime
 Print #I, "#STOP     : " + StopTime
End If

 Print #I, ""
If GlbCHPM = 1 Then Print #I, "#DIRECTION: " + Str$(PlaySelected)
If GlbCHCP = 1 Then If lstSec.ListCount >= BackIndex + 1 Then Print #I, "#FOCUS    : " + Str(BackIndex)
 Print #I, ""
 Print #I, " End of playlist file --^--"
 Print #1, " Creater  : Rennie Fotten"
 Print #1, " Tester   : Tester is needed. Call: 8 (067) 321-18-36"
 Print #1, " PrgValue : " & Format(App.Revision, "0000") & " edition, 99%"
 Print #1, ""
 Print #1, " PLAYER INFORMATION:"
 Print #1, ""
 Print #1, " PLAYER NAME:       RENNSOFT VIRTUAL MEGABOX"
 Print #1, " MODEL (VER):       VMBXT" + GetVersion + "E" + Format(App.Revision, "0")


Close I

pinfo.Hide

F_LED.BackColor = RGB(0, 0, 0)

If Err = 0 Then SaveList = True

End Function

Sub Scr_Down()
MMHeader.UpdateInterval = 0
RealUpdate.Enabled = False
End Sub

Sub Scr_Scroll(X As Single)
Dim aMins, aSecs

Dim Melisa As Single
Melisa = CCur(MMHeader.Length) / Scroller.ScaleWidth * X

   ScrChange Scroller.ScaleWidth, CInt(X)
   
   aMins = Fix((Melisa / 1000) / 60)
   aSecs = Melisa / 1000 Mod 60
   tmPos.TimeSet = Format$(aMins, "00") + ":" + Format$(aSecs, "00")
   
   If Melisa < MMHeader.Position Then
    Moding 617
   Else
    Moding 618
   End If


End Sub

Sub Scr_Up(X As Single)
Dim FUCK As Currency
Dim Melisa As Single
Dim RetState As Integer

RetState = MMHeader.Mode

Melisa = CCur(MMHeader.Length) / Scroller.ScaleWidth * X

ScrChange Scroller.ScaleWidth, CInt(X)

MMHeader.To = Melisa
MMHeader.Command = "Seek"
If RetState = 526 Then MMHeader.Command = "Play"
If RetState = 529 Then MMHeader.Command = "Play": MMHeader.Command = "Pause"
MMHeader.UpdateInterval = 250
RealUpdate.Enabled = True


End Sub

Sub SelectMode(Mode As Integer)

Select Case Mode
Case 0
 ioREPEAT.ForeColor = RGB(50, 50, 0)
 ioTRK.ForeColor = RGB(50, 50, 0)
 ioALL.ForeColor = RGB(50, 50, 0)
 Setting.GlbRepeat = Mode
 
Case 1
 ioREPEAT.ForeColor = RGB(250, 250, 0)
 ioTRK.ForeColor = RGB(50, 50, 0)
 ioALL.ForeColor = RGB(250, 250, 0)
 Setting.GlbRepeat = Mode
 
Case 2
 ioREPEAT.ForeColor = RGB(250, 250, 0)
 ioTRK.ForeColor = RGB(250, 250, 0)
 ioALL.ForeColor = RGB(50, 50, 0)
 Setting.GlbRepeat = Mode
End Select

UpdateSetup

End Sub

Sub SelectPlay(Mode As Integer)

pDirect.Picture = PCDiR.GraphicCell(Mode)
Setting.GlbPlayMode = Mode
UpdateSetup

End Sub

Sub StopCode()

ResetStatus

Select Case GlbSTM
 Case 0: Exit Sub
 Case 1: Unload Me: End
 Case 2: SendMessage Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0: Unload Me: End
End Select

End Sub

Sub TotalChange(Max As Currency, Min As Currency)
Dim MOCbKA As Currency
Dim Valve As Currency
Dim Vise As Currency

scrTotal.Scale (0, 0)-(150, 1)
Vise = 150 / Max * Min
scrTotal.Cls

scrTotal.Line (Vise, 0)-(Vise + 1, 1), RGB(255, 255, 0), BF

For MOCbKA = Vise To Vise + 10
 Valve = MOCbKA - Vise
 scrTotal.Line (MOCbKA, 0)-(MOCbKA + 1, 1), RGB(255 - (255 / 10 * Valve), 255 - (255 / 10 * Valve), 0), BF
Next

For MOCbKA = Vise - 10 To Vise
 Valve = MOCbKA - (Vise - 10)
 scrTotal.Line (MOCbKA, 0)-(MOCbKA + 1, 1), RGB(255 / 10 * Valve, 255 / 10 * Valve, 0), BF
Next

End Sub

Sub ScrChange(Max As Integer, Min As Integer)
Dim MOCbKA As Integer
Dim Valve As Integer
Dim Vise As Integer

Vise = 150 / Max * Min
Scroller.Cls

 For MOCbKA = Vise To Vise + 10
 Valve = MOCbKA - Vise
 Scroller.Line (MOCbKA, 0)-(MOCbKA + 1, 1), RGB(255 - (255 / 10 * Valve), 200 - (200 / 10 * Valve), 0), BF
 Next

 For MOCbKA = Vise - 10 To Vise
 Valve = MOCbKA - (Vise - 10)
 Scroller.Line (MOCbKA, 0)-(MOCbKA + 1, 1), RGB(100 / 10 * Valve, 255 / 10 * Valve, 0), BF
 Next


End Sub



Sub UpdateRecentMenu()
On Error Resume Next
 
 Dim G, H, J
 
 For G = 1 To mnuRecItm.UBound
   Unload mnuRecItm(G)
 Next G
 
 mnuRecItm(0).Caption = "< no items >"
 mnuRecItm(0).Enabled = False
 
 Dim xLit As String
 H = 0
 xLit = Dir(LowPath(GlbRecentDir) + "r_*.mxt", vbNormal)
 If xLit = "" Then Exit Sub
  
 mnuRecItm(0).Caption = Mid(xLit, 3, Len(xLit) - 4 - 2)
 mnuRecItm(0).Enabled = True
  
 Do While xLit > ""
    xLit = Dir
    If xLit = "" Then Exit Do
    H = H + 1
    Load mnuRecItm(H)
    mnuRecItm(H).Visible = True
    mnuRecItm(H).Enabled = True
    mnuRecItm(H).Caption = Mid(xLit, 3, Len(xLit) - 4 - 2)
 Loop

 
End Sub

Sub UpdateFavourites()
On Error Resume Next
 
 Dim G, H, J, L
 
 For G = 1 To mnuItm.UBound
   Unload mnuItm(G)
 Next G
 
 mnuItm(0).Caption = "< no items >"
 mnuItm(0).Enabled = False
 
 Dim xLit As String
 H = 0: L = 0
 J = FreeFile
 
 
On Error Resume Next

Dim I, K

I = FreeFile

Open (LowPath(App.Path) + "favourites.vmb") For Random As #I Len = Len(RecFileHead)

J = LOF(I) / Len(RecFileHead)

For K = 1 To J
  Get #I, K, RecFileHead
  If RecFileHead.rfUsed = True Then
   L = L + 1
   xLit = Trim(RecFileHead.rfDescription)
   If L = 1 Then
    mnuItm(0).Caption = xLit
    mnuItm(0).Enabled = True
   Else
    If Trim(xLit) = "" Then Exit For
    H = H + 1
    Load mnuItm(H)
    mnuItm(H).Visible = True
    mnuItm(H).Enabled = True
    mnuItm(H).Caption = xLit
   End If
  End If
Next

Close I
 
End Sub


Function FileNameU(fTitle As String) As String

Dim G, H, J, L, I, K
 
On Error Resume Next

I = FreeFile

Open (LowPath(App.Path) + "favourites.vmb") For Random As #I Len = Len(RecFileHead)

J = LOF(I) / Len(RecFileHead)

For K = 1 To J
  Get #I, K, RecFileHead
  If RecFileHead.rfUsed = True Then
    If Trim(RecFileHead.rfDescription) = fTitle Then
       FileNameU = Trim(RecFileHead.rfName)
       Exit For
    End If
  End If
Next K

Close I
 
End Function



Sub UpdateTags(Index As Integer)

On Error Resume Next

Dim X As Integer, Y As Integer, NN As Integer

Y = lstSec.ListIndex

lblMask.Caption = "Updating information..."
lblMask.Visible = True
lstSec.Visible = False
pinfo.Show 0, Me
pinfo.lblSesc.Caption = "Updating..."
DoEvents

If Index > lstMain.ListCount - 1 Then Index = -1

If Index > -1 Then
  lstSec.List(Index) = GetMp3Song(lstMain.List(Index), NN)
  lstTimes.List(Index) = Str(NN)
Else
 For X = 0 To lstMain.ListCount - 1
  lstSec.List(X) = GetMp3Song(lstMain.List(X), NN)
  lstTimes.List(X) = Str(NN)
 Next X
End If

Unload pinfo
lstSec.ListIndex = Y

lblMask.Visible = False
lstSec.Visible = True


End Sub


Private Sub btnAdd_Click()
wndAddFiles.Tag = "save"
wndAddFiles.Show 0, Me
End Sub

Private Sub btnClear_Click()
Dim Q
Q = MsgBox("Are you sure want to clear the playlist entries?", vbQuestion + vbYesNo, "RENNSoft Virtual MegaBox. Version XT " + GetVersion)
If Q = 6 Then lstMain.Clear: lstSec.Clear: ZagalTime
End Sub

Private Sub btnCopy_Click()
 dlgCopyFile.Show 1, Me
End Sub

Private Sub btnDelete_Click()
On Error Resume Next
For ZU = lstSec.ListCount - 1 To 0 Step -1
  If lstSec.Selected(ZU) = True Then
    lstMain.RemoveItem (ZU)
    lstSec.RemoveItem (ZU)
    lstTimes.RemoveItem (ZU)
  End If
Next

ZagalTime

End Sub


Private Sub btnExit_Click()
Dim RetVal
RetVal = MsgBox("Are you sure want to exit?", vbQuestion + vbYesNo, "Virtual MegaBox XT " + GetVersion)
If RetVal = 6 Then
  Unload Me
  End
End If
End Sub

Private Sub btnExplore_Click()
 On Error Resume Next
 Dim SndVolume As Double
 SndVolume = Shell("explorer.exe", vbNormalFocus)
End Sub

Private Sub btnFav_Click()
  PopupMenu mnuFavour, 1, btnFav.Left + tblListCtrl.Left, btnFav.Top + tblListCtrl.Top + btnFav.Height
End Sub

Private Sub btnFind_Click()
If lstMain.ListCount = 0 Then Exit Sub

frmFind.Show 1, Me

End Sub

Private Sub btnFont_Click()
On Error Resume Next
Dim A
A = txtFileData.SelStart
txtFileData.SelStart = 0
txtFileData.SelLength = Len(txtFileData.Text)
txtFileData.SelFontSize = 8
txtFileData.SelStart = A
End Sub

Private Sub btnHide_Click()
 frmBorder.Visible = False
 wndTitle.Visible = False
 cSysTray.InTray = True
 Me.Hide
End Sub




Private Sub btnPlF_Click()
wndAddFiles.Tag = "Open"
wndAddFiles.Show 0, Me

End Sub

Private Sub btnPLMnu_Click()
  On Error Resume Next
  mnuRecQ.Visible = True
  mnuFavour.Visible = True
  PopupMenu mnuPM, , (tblListCtrl.Left + btnPLMnu.Left) + 50, (tblListCtrl.Top + btnPLMnu.Top)
End Sub

Private Sub btnRecPL_Click()
  PopupMenu mnuRecQ, 0, btnRecPL.Left + tblListCtrl.Left, btnRecPL.Top + tblListCtrl.Top + btnRecPL.Height
End Sub

Private Sub btnSC_Click()

SelectPlay 2

LC

End Sub

Private Sub Command1_Click()

   PopupMenu mSetx, , Command1.Left + 200, Command1.Top + 350

End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button Then PopupMenu abc, , Command1.Left + 200, Command1.Top + 350

End Sub



Private Sub Form_Activate()
UpdateRecentMenu
UpdateFavourites
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)


' Fxx Keys Definition
If KeyCode = 119 Then btnKill_Click
If KeyCode = 113 Then btnSave_Click
If KeyCode = 114 Then btnLoad_Click
If KeyCode = 120 Then btnKill_Click
If KeyCode = 116 Then btnCopy_Click

' SpecCode Definition
If KeyCode = 27 Then btnExit_Click
If KeyCode = 46 Then btnDelete_Click
If KeyCode = 45 Then btnAdd_Click

' Macros
If KeyCode = 93 Then Command1_MouseDown 1, 0, 0, 0

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SSPanel1_MouseUp Button, Shift, X, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveList LowPath(App.Path) + "default.mxt", wndMain.lstMain
' Save Programm settings
GlbSaveSettings

Unload frmBorder

End Sub

Private Sub gDown_Click()
On Error Resume Next
Dim ZU As Integer

For ZU = lstSec.ListCount - 1 To 0 Step -1
 If lstSec.Selected(ZU) = True Then
  If ZU < lstMain.ListCount - 1 Then
   ExchangeFiles ZU, ZU + 1, Me.lstMain
   ExchangeFiles ZU, ZU + 1, Me.lstSec
   ExchangeFiles ZU, ZU + 1, Me.lstTimes
' lstMain.ListIndex = lstMain.ListIndex + 1
' lstSec.ListIndex = lstSec.ListIndex + 1
  End If
 End If
Next

CenterList2

End Sub


Private Sub gUp_Click()
On Error Resume Next
Dim ZU As Integer

For ZU = 0 To lstSec.ListCount - 1 Step 1
 If lstSec.Selected(ZU) = True Then
  If ZU > 0 Then
   ExchangeFiles ZU, ZU - 1, Me.lstMain
   ExchangeFiles ZU, ZU - 1, Me.lstSec
   ExchangeFiles ZU, ZU - 1, Me.lstTimes
' lstMain.ListIndex = lstMain.ListIndex + 1
' lstSec.ListIndex = lstSec.ListIndex + 1
  End If
 End If
Next

CenterList2

End Sub

Private Sub htmlply_Click()

HTM.Show

End Sub

Private Sub idtaged_Click()
On Error Resume Next
Z = Shell(LowPath(App.Path) + "PLUGINS\idtag.exe " + lstMain.List(lstSec.ListIndex), vbNormalFocus)
If Err Then MsgBox "Error! idTAG Editor not installed!", vbCritical
If Err = 0 Then
 MsgBox "After idTAG editor stops, press 'Ok' to update the playlist entries information."
 UpdateTags -1
End If

End Sub


Private Sub lblName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SSPanel1_MouseDown Button, Shift, X, Y
End Sub


Private Sub lblName_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SSPanel1_MouseMove Button, Shift, X, Y
End Sub


Private Sub lblName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SSPanel1_MouseUp Button, Shift, X, Y
End Sub

Private Sub lblTitle_Click()
If lstSec.ListCount > 0 Then CenterList: lstSec.ListIndex = FindPlayDeal: lstSec.SetFocus
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SSPanel1_MouseDown Button, Shift, X, Y
End Sub


Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SSPanel1_MouseMove Button, Shift, X, Y
End Sub


Private Sub lblTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SSPanel1_MouseUp Button, Shift, X, Y
End Sub

Private Sub lblVersion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SSPanel1_MouseDown Button, Shift, X, Y
End Sub


Private Sub lblVersion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SSPanel1_MouseMove Button, Shift, X, Y
End Sub


Private Sub lblVersion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SSPanel1_MouseUp Button, Shift, X, Y
End Sub

Private Sub lstMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstSec.ListIndex = lstMain.ListIndex
lstSec_MouseDown Button, Shift, X, Y
End Sub


Private Sub lstMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstSec_MouseUp Button, Shift, X, Y

End Sub


Private Sub lstSec_Click()
On Error Resume Next
lstMain.ListIndex = lstSec.ListIndex

If lstSec.ListCount = 0 Then BackIndex = 0

Dim I As Currency, B

ZagalTime
PlayablePos

End Sub

Sub ZagalTime()

On Error Resume Next
I = 0
For B = 0 To lstSec.ListCount - 1
   If lstSec.Selected(B) = True Then I = I + Val(lstTimes.List(B))
Next

   Min = Fix(I / 60) Mod 60
   Hr = Fix(I / 3600)
   Sc = I Mod 60
   lblZag.Caption = Format$(Hr, "00") + ":" + Format$(Min, "00") + ":" + Format$(Sc, "00")

End Sub

Private Sub lstSec_DblClick()
On Error Resume Next
Dim NN As Integer
ResetStatus
BackIndex = lstSec.ListIndex
lstSec.List(lstSec.ListIndex) = GetMp3Song(lstMain.List(lstSec.ListIndex), NN)
lstTimes.List(lstSec.ListIndex) = Str(NN)
Call Command_Open: Command_Play
End Sub


Private Sub lstSec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And lstMain.ListIndex > -1 Then infForm.oFileName = lstMain.List(lstMain.ListIndex): infForm.Show 1, Me
End Sub


Private Sub lstSec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

If Button = 2 Or lstSec.ListIndex = -1 Then Exit Sub
lstMain.ListIndex = lstSec.ListIndex
lstMain.Selected(lstSec.ListIndex) = lstSec.Selected(lstSec.ListIndex)

End Sub


Private Sub lstSec_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim blt As String, NN As Integer

For bT = 1 To Data.Files.Count
  blt = Data.Files(bT)
  lstMain.AddItem blt
  lstSec.AddItem GetMp3Song(blt, NN)
  lstTimes.AddItem Str(NN)
Next bT


End Sub

Private Sub lstSec_Scroll()
PlayablePos
End Sub

Private Sub mnuAddAll_Click()

On Error Resume Next

Dim R, T

T = lstSec.ListIndex

For R = 0 To mnuItm.Count - 1
 lstMain.AddItem FileNameU(mnuItm(R).Caption)
 lstSec.AddItem mnuItm(R).Caption
Next R

lstSec.ListIndex = T

End Sub

Private Sub mnuAddAllFav_Click()

Dim X

For X = 0 To lstMain.ListCount - 1
  SaveFileHdr lstMain.List(X), lstSec.List(X)
Next

UpdateFavourites

End Sub

Private Sub mnuAddFav_Click()

SaveFileHdr lstMain.List(lstSec.ListIndex), lstSec.List(lstSec.ListIndex)

UpdateFavourites

End Sub

Private Sub mnuChPt_Click()
On Error Resume Next
Dim xPath As String
xPath = InputBox("Enter the new path here:", "Change path")
If xPath > "" Then
  ChDir xPath
  If Mid(xPath, 2, 1) = ":" Then
    ChDrive Mid(xPath, 1, 2)
  End If
End If

If Err Then MsgBox Err.Description, vbCritical, "Path change error"

End Sub

Private Sub mnuCP_Click()
GlbCHCP = 1 - GlbCHCP
UpdateSetup

End Sub

Private Sub mnuDelFav_Click()

frmDel.Show

End Sub

Private Sub mnuDirF_Click(Index As Integer)
SelectPlay (Index)
End Sub

Private Sub mnuFLF_Click()
GlbSeeker = 1 - GlbSeeker
UpdateSetup
End Sub

Private Sub mnuFS_Click()
GlbCHFS = 1 - GlbCHFS
UpdateSetup
End Sub

Private Sub mnuItm_Click(Index As Integer)

On Error Resume Next

lstMain.AddItem FileNameU(mnuItm(Index).Caption)
lstSec.AddItem mnuItm(Index).Caption
lstSec.ListIndex = lstSec.ListCount - 1
Command_Open
Command_Play

End Sub

Private Sub mnuLoadRec_Click()

lstRecent.Show


End Sub

Private Sub mnuOff_Click()
txtPad.Visible = False
End Sub

Private Sub mnuNoTags_Click()
GlbNoTags = 1 - GlbNoTags
UpdateSetup
End Sub

Private Sub mnuPau_Click()

On Error Resume Next
Dim E As String
E = InputBox("Change pause between tracks", "Settings", GlbPause)
If E > "" Then GlbPause = Val(E): UpdateSetup

End Sub

Private Sub mnuPlM_Click()
GlbCHPM = 1 - GlbCHPM
UpdateSetup
End Sub

Private Sub mnuPP_Click()
GlbCHPP = 1 - GlbCHPP
UpdateSetup

End Sub

Private Sub mnuprotime_Click()
On Error Resume Next
Dim E As Integer
E = Val(InputBox("1 minute = X bytes", "Proportional time set", GlbPropTime))
GlbPropTime = E
UpdateSetup
End Sub

Private Sub mnuRecItm_Click(Index As Integer)

If FileExists(LowPath(GlbRecentDir) + "r_" + mnuRecItm(Index).Caption + ".mxt") = True Then
 wndMain.LoadList LowPath(GlbRecentDir) + "r_" + mnuRecItm(Index).Caption + ".mxt", wndMain.lstMain, wndMain
End If

End Sub

Private Sub mnuRecLoc_Click()
recLoc.Show
End Sub

Private Sub mnuRepS_Click(Index As Integer)
SelectMode (Index)
End Sub

Private Sub mnuResetSet_Click()
Setting.GlbLoadSettings
UpdateSetup
End Sub

Private Sub mnuSaveRec_Click()

Dim rcName As String

rcName = InputBox("Enter the recent playlist name:", "Save recent playlist", xDefaut)
If rcName > "" Then
  xDefault = rcName
  SaveList LowPath(GlbRecentDir) + "r_" + rcName + ".mxt", wndMain.lstMain
End If

End Sub

Private Sub mnuSaveSet_Click()
Setting.GlbSaveSettings
End Sub

Private Sub mnuSTO_Click(Index As Integer)
GlbSTM = Index
UpdateSetup
End Sub

Private Sub mnuTAG_Click()
On Error Resume Next
Dim E As String
Dim V As String

V = vbCrLf
V = V + "%1 - Filename" + vbCrLf
V = V + "%2 - Filepath" + vbCrLf
V = V + "%3 - File ext. (w/o point)" + vbCrLf
V = V + "%TAB - Insert a TAB symbol" + vbCrLf
V = V + "%4 - Title" + vbCrLf
V = V + "%5 - Artist" + vbCrLf
V = V + "%6 - Album" + vbCrLf
V = V + "%7 - Year" + vbCrLf
V = V + "%8 - Gerne" + vbCrLf
V = V + "%9 - File type" + vbCrLf
V = V + "%10 - File Size" + vbCrLf
V = V + "%11 - Proportional time"

E = InputBox("Change TAG format: " + V, "Settings", GlbTitle)
If E > "" Then GlbTitle = E: UpdateSetup
End Sub

Private Sub mnuTAGLESS_Click()

On Error Resume Next
Dim E As String
Dim V As String

V = vbCrLf
V = V + "%1 - Filename" + vbCrLf
V = V + "%2 - Filepath" + vbCrLf
V = V + "%3 - File ext. (w/o point)" + vbCrLf
V = V + "%TAB - Insert a TAB symbol" + vbCrLf
V = V + "%4 - Title" + vbCrLf
V = V + "%5 - Artist" + vbCrLf
V = V + "%6 - Album" + vbCrLf
V = V + "%7 - Year" + vbCrLf
V = V + "%8 - Gerne" + vbCrLf
V = V + "%9 - File type" + vbCrLf
V = V + "%10 - File Size" + vbCrLf
V = V + "%11 - Proportional time"

E = InputBox("Change TAGless format: " + V, "Settings", GlbTagLess)
If E > "" Then GlbTagLess = E: UpdateSetup

End Sub

Private Sub mnuTS_Click()
GlbCHTS = 1 - GlbCHTS
UpdateSetup

End Sub

Private Sub mnuTupd_Click()
 UpdateTags -1
End Sub

Private Sub mPSE_Click()
Setting.GlbSaveSettings
Settings.Show

End Sub



Private Sub picCover_Click()

Dim U As Integer

If picCover.Top = 0 Then
 For U = 0 To 1020 Step 100
  picCover.Top = U
  SleepX 0.01
 Next
 picCover.Top = 1020
Else
 For U = 1020 To 0 Step -100
  picCover.Top = U
  SleepX 0.01
 Next
 picCover.Top = 0
End If

End Sub


Private Sub renameFile_Click()
On Error Resume Next
Dim UserDest As String, FullDst As String, FullSrc As String

UserDest = InputBox("Change file name ''" + FileHead(lstMain.List(lstSec.ListIndex)) + "'' to:", "Rename File")

If UserDest = "" Then Exit Sub

MMHeader.Command = "Close"
FullSrc = lstMain.List(lstSec.ListIndex)
FullDst = LowPath(PathHead(lstMain.List(lstSec.ListIndex))) + UserDest
Name FullSrc As FullDst

If Err Then MsgBox "Renaming error!!!", vbCritical
If Err = 0 Then lstMain.List(lstSec.ListIndex) = FullDst
If Err = 0 Then lstSec.List(lstSec.ListIndex) = "[Renamed to " + UserDest + "]"

End Sub

Private Sub Scroller_Scroll()

End Sub

Private Sub Scroller_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then

 If X > Scroller.ScaleWidth Then X = Scroller.ScaleWidth
 If X < 0 Then X = 0

 Scr_Scroll (X)
End If

End Sub

Private Sub SetTimeBut_Click(Index As Integer)


If opStart.Value = True Then

  If Index = 11 Then tmrStart.Reset: StartTime = "5555": Exit Sub
  
  StartTime = StartTime + Format(Index, "0")
  If Len(StartTime) > 4 Then
    StartTime = Mid(StartTime, 2, 4)
  Else
    StartTime = Mid(StartTime, 1, 4)
  End If
  tmrStart.TimeSet = Mid(StartTime, 1, 2) + "." + Mid(StartTime, 3, 2)
End If

If opStop.Value = True Then
  
  If Index = 11 Then tmrStop.Reset: StopTime = "5555": Exit Sub
   
  StopTime = StopTime + Format(Index)
  If Len(StopTime) > 4 Then
    StopTime = Mid(StopTime, 2, 4)
  Else
    StopTime = Mid(StopTime, 1, 4)
  End If
  tmrStop.TimeSet = Mid(StopTime, 1, 2) + "." + Mid(StopTime, 3, 2)
End If
  

End Sub

Private Sub SSPanel1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

CanMove = False

End Sub

Private Sub stngr_Click()
On Error Resume Next
Shell LowPath(App.Path) + "PLUGINS\stinger.exe", vbNormalFocus
If Err Then MsgBox "Error! Stinger not installed!", vbCritical
End Sub

Private Sub btnKill_Click()
On Error Resume Next
Dim RetVal, N

If lstSec.ListIndex < 0 Then Exit Sub

If lstSec.SelCount = 1 Then
  RetVal = MsgBox("Are you sure want to delete the ''" + lstMain.List(lstSec.ListIndex) + "''?", vbExclamation + vbYesNo, "Delete File From Disk")
  If RetVal = 6 Then
     If lstMain.List(lstSec.ListIndex) = MMHeader.Filename Then ResetStatus
     Kill lstMain.List(lstSec.ListIndex)
       If Err Then
         MsgBox "Can't delete the file" + Chr(13) + Chr$(13) + lstMain.List(lstMain.ListIndex), vbCritical
       Else
     lstMain.RemoveItem lstSec.ListIndex
     lstSec.RemoveItem lstSec.ListIndex
     lstTimes.RemoveItem lstSec.ListIndex
    End If
End If
End If

If lstSec.SelCount > 1 Then
  RetVal = MsgBox("Are you sure want to delete " + Format(lstMain.SelCount, "0") + " files?", vbExclamation + vbYesNo, "Delete Files From Disk")
  If RetVal = 6 Then
   For N = lstMain.ListCount - 1 To 0 Step -1
       If lstMain.List(N) = MMHeader.Filename Then ResetStatus
       If lstMain.Selected(N) = True Then Kill lstMain.List(N)
       
       If Err Then
         MsgBox "Can't delete the file" + Chr(13) + Chr$(13) + lstMain.List(N), vbCritical
       Else
         lstMain.RemoveItem N
         lstSec.RemoveItem N
         lstTimes.RemoveItem N
       End If
   
   Next
  End If
End If

End Sub

Private Sub btnLN_Click()
On Error Resume Next
Dim RV
RV = InputBox("Enter LN", , btnLN.Caption)
LN = RV
btnLN.Caption = Str(LN)

End Sub


Private Sub btnNext_Click()
Call Command_Next
End Sub

Private Sub btnPause_Click()
Call Command_Pause
End Sub


Private Sub btnPlay_Click()


If MMHeader.Mode = 529 Then
 Call Command_Pause
Else
 If MMHeader.Mode = 524 Then
  If lstMain.ListCount = 0 Then
    btnPlF_Click
  Else
    Call Command_Open: Command_Play
  End If
 Else
  Command_Play
 End If
End If

End Sub

Private Sub btnPrev_Click()
Call Command_Prev
End Sub

Private Sub btnRNDSort_Click()
Dim Z As Integer, Y As Currency, X As Integer
On Error Resume Next

If lstMain.ListCount = 0 Then Exit Sub
X = lstSec.ListIndex

lblMask.Caption = "Sorting playlist [ RANDOMIZED ]..."
lblMask.Visible = True
lstSec.Visible = False
DoEvents
For Z = 0 To lstSec.ListCount - 1
 Y = Rnd
 ExchangeFiles Z, Fix(Y * lstMain.ListCount - 1), wndMain.lstMain
 ExchangeFiles Z, Fix(Y * lstSec.ListCount - 1), wndMain.lstSec
 ExchangeFiles Z, Fix(Y * lstTimes.ListCount - 1), wndMain.lstTimes
Next

lstSec.ListIndex = X

lblMask.Visible = False
lstSec.Visible = True


End Sub

Private Sub btnSelect_Click()
Dim K, O
O = lstSec.ListIndex

lblMask.Caption = "Selecting..."
lblMask.Visible = True
lstSec.Visible = False
DoEvents

For K = 0 To lstMain.ListCount - 1
lstMain.Selected(K) = True
lstSec.Selected(K) = True
Next

lstSec.ListIndex = O

lblMask.Visible = False
lstSec.Visible = True


End Sub

Private Sub btnUnSelect_click()



 lblMask.Caption = "Selecting..."
 lblMask.Visible = True
 lstSec.Visible = False
 DoEvents


 Dim K
 For K = 0 To lstMain.ListCount - 1
  lstMain.Selected(K) = False
  lstSec.Selected(K) = False
 Next
 
 lblMask.Visible = False
 lstSec.Visible = True


End Sub


Private Sub btnSetup_Click()
 PopupMenu mSetx, , BtnSetup.Left + 100, BtnSetup.Top + BtnSetup.Height + 80
End Sub

Private Sub btnStop_Click()
If MMHeader.Mode = 526 Or MMHeader.Mode = 529 Then
  Command_Stop
Else
  ResetStatus
End If
End Sub

Private Sub btnTimer_Click()
 
If btnTimer.Value Then Tik = 5 Else Tik = 0

If btnTimer.Value Then
  gStart.Enabled = True
  ledStart.BackColor = RGB(0, 255, 0)
  gStop.Enabled = True
  ledStop.BackColor = RGB(0, 255, 0)
Else
  gStart.Enabled = False
  ledStart.BackColor = RGB(0, 64, 0)
  gStop.Enabled = False
  ledStop.BackColor = RGB(0, 64, 0)
End If

End Sub

Private Sub btnVolume_Click()
 On Error Resume Next
 Dim SndVolume As Double
 SndVolume = Shell("sndvol32.exe", vbNormalFocus)
 If SndVolume = 0 Then MsgBox "Can't initialize the Volume Control", vbCritical, "Windows Error"
End Sub

Private Sub btnWindow_Click()
 If wndTitle.Visible = True Then wndTitle.Visible = False: Exit Sub
 If wndTitle.Visible = False Then wndTitle.Visible = True:: Exit Sub
End Sub

Private Sub cSysTray_MouseDown(Button As Integer, Id As Long)
 Me.Show
 frmBorder.Visible = True
 cSysTray.InTray = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim KeyCode

KeyCode = Asc(UCase(Chr$(KeyAscii)))
Missed = True

If Chr(KeyCode) = "Z" Then btnPrev_Click
If Chr(KeyCode) = "C" Then btnNext_Click
If Chr(KeyCode) = "X" Then btnPlay_Click
If Chr(KeyCode) = "V" Then btnPause_Click
If Chr(KeyCode) = "B" Then btnStop_Click

If KeyCode = 39 Then Scr_Up ((MMHeader.Position / 1000) + 10)

Fucka = Fucka + Chr$(KeyCode)
Fucka = Right$(Fucka, 4)


If Fucka = "$-40" Then MsgBox "CONGRATULATIONS-> You Got The Cheat for LN!", vbExclamation: _
   btnLN.Visible = True: _
   Fucka = ""
   
If Fucka = "+486" Then MsgBox "CONGRATULATIONS-> You Got The Cheat for ENCH!", vbExclamation: _
   Ench = 1: Fucka = ""

If Fucka = "2011" Then
   MMHeader.DeviceType = InputBox("Enter the Custom Device Type", "(!WARNING!) Reprogramming (!WARNING!)", MMHeader.DeviceType)
End If
      
If Fucka = "MAIN" Then
  lstMain.Top = lstSec.Top
  lstMain.Left = lstSec.Left
  lstMain.Width = lstSec.Width
  lstMain.Height = lstSec.Height
  lstMain.Visible = True
  lstSec.Visible = False
End If

If Fucka = "@MIN" Then
  lstSec.Visible = True
  lstMain.Visible = False
End If


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SSPanel1_MouseDown Button, Shift, X, Y
Missed = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SSPanel1_MouseMove Button, Shift, X, Y

End Sub

Private Sub btnLoad_Click()
On Error Resume Next
dlgOpenSave.Filename = DefFile
dlgOpenSave.ShowOpen
If Err Then Exit Sub
lblMask.Caption = "Loading playlist..."
lblMask.Visible = True
lstSec.Visible = False
DoEvents
LoadList dlgOpenSave.Filename, Me.lstMain, Me
lblMask.Visible = False
lstSec.Visible = True

End Sub

Private Sub btnMode_Click()
SelectPlay (Setting.GlbPlayMode + 1) Mod 2

End Sub

Private Sub btnRepeat_Click()
SelectMode ((Setting.GlbRepeat + 1) Mod 3)
End Sub

Private Sub btnSave_Click()

' Ignore all errors
On Error Resume Next

' Show "SAVE" Dialog
dlgOpenSave.ShowSave

' If Cancel then Cancel Saving
If Err Then
  Exit Sub
Else
  SaveList dlgOpenSave.Filename, wndMain.lstMain
  DefFile = dlgOpenSave.Filename
End If

End Sub

Private Sub Form_Load()
' On Error Resume Next

frmBorder.Show

Facing
Ench = CCPU
PicCPUSub
tmrStart.Reset
tmrStop.Reset

' Load defaults
SelectMode (0)
SelectPlay (0)
ResetStatus
TotalChange 255, 0
' Welcome title
wndTitle.lab(0) = " WELCOME TO THE VIRTUAL MEGABOX VERSION XT " + GetVersion

ApplySettings
wndTitle.wndLS
UpdateSetup
Rots 0

If Command$ > "" Then
 If LCase(Right(Command$, 3)) = "mxt" Or LCase(Right(Command$, 3)) = "vmb" Then
   LoadList Command$, Me.lstMain, Me
 Else
   lstMain.AddItem Command$
   lstMain.ListIndex = 0
   lstSec.ListIndex = 0
   btnPlay_Click
 End If
Else
 If FileExists(LowPath(App.Path) + "default.mxt") Then LoadList LowPath(App.Path) + "default.mxt", Me.lstMain, Me
End If

On Error Resume Next
frmTip.Show vbModal, Me

' On Meters
' VU1.VU_ON

ledCont.BackColor = RGB(0, 0, 255)



End Sub

Private Sub Form_Resize()
Dim Y, MY, MX, X

frmBorder.Left = wndMain.Left
frmBorder.Top = wndMain.Top - frmBorder.Height
 

MY = tblTimer.Height / Screen.TwipsPerPixelY
For Y = 0 To MY
tblTimer.Line (0, Y * Screen.TwipsPerPixelY)-(tblTimer.Width, Y * Screen.TwipsPerPixelY), RGB((100 / MY * Y), (100 / MY * Y), 0)
Next

MY = (SSPanel1.Width + SSPanel1.Height) / Screen.TwipsPerPixelY
For Y = 0 To MY
  SSPanel1.Line (0, Y * Screen.TwipsPerPixelY)-(Y * Screen.TwipsPerPixelX, 0), RGB(50 + (150 / MY * Y), 50 + (150 / MY * Y), 0)
Next

 Dim A As Integer, CL As Integer
   For A = 0 To SSPanel1.Height
    CL = 70 + (40 * Sin(A))
    SSPanel1.Line (0, A + 1300)-(SSPanel1.Width, A + 1300), RGB(CL, CL, 0), BF
   Next

MY = tblListCtrl.Height / Screen.TwipsPerPixelY
For Y = 0 To MY
   tblListCtrl.Line (0, Y * Screen.TwipsPerPixelY)-(tblListCtrl.Width, Y * Screen.TwipsPerPixelY), RGB((100 / MY * Y), (100 / MY * Y), 0)
Next

' Dim a As Integer, CL As Integer
For A = 0 To tblListCtrl.Height Step 25
  CL = 70 + (40 * Sin(A))
  tblListCtrl.Line (0, A)-(tblListCtrl.Width, A), RGB(CL, CL, 0), BF
Next


MX = Fix(wndMain.Width / Screen.TwipsPerPixelX)
For X = 0 To MX
  Line (X * Screen.TwipsPerPixelX, 0)-(X * Screen.TwipsPerPixelX, wndMain.Height), RGB((100 / MX * (MX - X)), (100 / MX * (MX - X)), 0)
Next

Me.Line (0, 0)-(Me.Width - 25, 0), RGB(155, 155, 0)
Me.Line (0, Me.Height - 30)-(Me.Width - 25, Me.Height - 30), RGB(55, 55, 0)
Me.Line (Me.Width - 25, 0)-(Me.Width - 25, Me.Height - 25), RGB(55, 55, 0)
Me.Line (0, 0)-(0, Me.Height - 30), RGB(155, 155, 0)

Me.Line (15, 15)-(Me.Width - 15 - 25, 15), RGB(100, 100, 0)
Me.Line (15, Me.Height - 15 - 30)-(Me.Width - 25 - 15, Me.Height - 30 - 15), RGB(30, 30, 0)
Me.Line (Me.Width - 15 - 25, 15)-(Me.Width - 25 - 15, Me.Height - 30 - 15), RGB(30, 30, 0)
Me.Line (15, 15)-(15, Me.Height - 30 - 15), RGB(100, 100, 0)

 
End Sub

Private Sub gStart_Timer()


If Format(Val(StartTime), "####") = Format$(Now, "Hmm") Then
 Command_Open
 Command_Play
 gStart.Enabled = False
 ledStart.BackColor = RGB(0, 64, 0)
 tmrStart.Reset
End If

If Format(Val(StartTime), "####") = "5555" Then
  gStart.Enabled = False
  ledStart.BackColor = RGB(0, 64, 0)
  tmrStart.Reset
End If

If gStart.Enabled = False And gStop.Enabled = False Then btnTimer.Value = 0

End Sub

Private Sub gStop_Timer()
 

If Format(Val(StopTime), "####") = Format$(Now, "Hmm") Then
  ResetStatus
  gStop.Enabled = False
  ledStop.BackColor = RGB(0, 64, 0)
  tmrStop.Reset
  StopTime = "5555"
End If
 
If Format(Val(StopTime), "####") = "5555" Then
  gStop.Enabled = False
  ledStop.BackColor = RGB(0, 64, 0)
  tmrStop.Reset
End If
 
 
If gStart.Enabled = False And gStop.Enabled = False Then btnTimer.Value = 0

End Sub


Private Sub MMHeader_Done(NotifyCode As Integer)
On Error Resume Next



If NotifyCode = 1 Or NotifyCode = 8 Then


  Sleep (GlbPause / 1000)
  ' -----------ALL MODE REPEAT----------------
    ' If Selected "All" mode repeat
    If Setting.GlbRepeat = 1 Then
   
     ' If play direction is "Down"
     If Setting.GlbPlayMode = 0 Then
      If lstSec.ListIndex = lstSec.ListCount - 1 Then
         If lstSec.ListIndex >= 0 Then
              lstSec.ListIndex = 0
              Command_Open
              Command_Play
              Exit Sub
         Else
              ResetStatus
         End If
      Else
        Command_Next
      End If
     End If
   
     ' If play direction is "Up"
     If Setting.GlbPlayMode = 1 Then
      If lstSec.ListIndex = 0 And _
       lstSec.ListCount > 1 Then
       lstSec.ListIndex = lstSec.ListCount - 1
       Command_Open
       Command_Play
       Exit Sub
      Else
       Command_Prev
      End If
     End If
   
    Exit Sub
    End If
  ' --------------- END ALL ------------------
  
  
  ' -----------TRK MODE REPEAT----------------
  ' If Selected "Trk" mode repeat
  If Setting.GlbRepeat = 2 Then
      Command_Stop
      Command_Play
    Exit Sub
  End If
  ' --------------- END TRK ------------------


  ' ---------------NO REPEAT------------------
  If Setting.GlbRepeat = 0 Then
    ResetStatus
    ' -------------DIRECTION IS DOWN----------
    If Setting.GlbPlayMode = 0 Then
     If lstSec.ListIndex < (lstSec.ListCount - 1) Then
       btnNext_Click
     Else
       StopCode
     End If
    End If
    ' ----------------END DOWN----------------
    ' --------------DIRECTION IS UP-----------
    If Setting.GlbPlayMode = 1 Then
     If lstSec.ListIndex > 0 Then
      btnPrev_Click
     Else
      StopCode
     End If
    End If
    ' ----------------END UP------------------
  Exit Sub
  End If
  ' --------------END NO REPEAT---------------


End If

End Sub

Sub Moding(Mode)
Select Case Mode
Case 525: Model.Picture = Modes.GraphicCell(Tik + 0)
Case 524: Model.Picture = Modes.GraphicCell(Tik + 0)
Case 526: Model.Picture = Modes.GraphicCell(Tik + 1)
Case 529: Model.Picture = Modes.GraphicCell(Tik + 2)
Case 617: Model.Picture = Modes.GraphicCell(Tik + 3)
Case 618: Model.Picture = Modes.GraphicCell(Tik + 4)
End Select

End Sub

Private Sub MMHeader_StatusUpdate()
' Standart

Quanta = Quanta + 1
Quanta = Quanta Mod 4

If Quanta > 1 Then
 F_LED.BackColor = RGB(0, 255, 0)
Else
 F_LED.BackColor = RGB(0, 0, 0)
End If

On Error Resume Next
If UFlag = 1 Then Exit Sub

MMHeader.TimeFormat = 0
' Dim aSecs, aMins, aSecls, bSecs, bMins, bSecls
Dim STimer, GX

' Set the length and position as seconds
aSecls = Fix(MMHeader.Length / 1000)
bSecls = Fix(MMHeader.Position / 1000)

If aSecls - bSecls < 3 Then
 picEnd.BackColor = RGB(255, 0, 0)
Else
 picEnd.BackColor = RGB(0, 0, 0)
End If


' Prepare for update indicators
aMins = Fix(aSecls / 60)
bMins = Fix(bSecls / 60)

aSecs = aSecls Mod 60
bSecs = bSecls Mod 60

' Update indicators
If aMins <= 99 Then
  tmPos.TimeSet = Format$(bMins, "00") + ":" + Format$(bSecs, "00")
  tmLen.TimeSet = Format$(aMins, "00") + ":" + Format$(aSecs, "00")
Else
  tmPos.TimeSet = Format$(Fix(bMins / 60), "00") + ":" + Format$(bMins Mod 60, "00")
  tmLen.TimeSet = Format$(Fix(aMins / 60), "00") + ":" + Format$(aMins Mod 60, "00")
End If

' Me.Caption = lblTitle.Caption + " [" + Format$(bMins, "00") + ":" + Format$(bSecs, "00") + "]"


cSysTray.TrayTip = lblTitle.Caption

' Scroller update
If aSecls > 0 Then
 ScrChange CInt(aSecls), CInt(bSecls)
 If TT = 0 Then TT = 1
   PT = 100 / TT * ((BackIndex) + (1 / aSecls * bSecls))
   Rots PT
   TotalChange 100, PT
End If



End Sub

Private Sub RealUpdate_Timer()
Moding (MMHeader.Mode)
LC
If Setting.GlbSTM Then AST.ForeColor = RGB(255, 0, 0) Else AST.ForeColor = RGB(55, 0, 0)

clckItem(0).Picture = TSR.GraphicCell(Val(Mid(Format(Now, "HH"), 1, 1)))
clckItem(1).Picture = TSR.GraphicCell(Val(Mid(Format(Now, "HH"), 2, 1)))

clckItem(2).Picture = TSR.GraphicCell(Val(Mid(Format(Now, "hh:mm"), 4, 1)))
clckItem(3).Picture = TSR.GraphicCell(Val(Mid(Format(Now, "hh:mm"), 5, 1)))

' Timer1_Timer

End Sub


Private Sub Scroller_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
If X > Scroller.ScaleWidth Then X = Scroller.ScaleWidth
If X < 0 Then X = 0

Scr_Down

End Sub

Private Sub Scroller_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
If X > Scroller.ScaleWidth Then X = Scroller.ScaleWidth
If X < 0 Then X = 0

Scr_Up (X)

End Sub





Private Sub SSPanel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

OldX = X
OldY = Y
CanMove = True

End Sub


Private Sub SSPanel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If CanMove = False Then Exit Sub

Dim ML, MT

If Button = 1 Then

 ML = Me.Left + (X - OldX)
 MT = Me.Top + (Y - OldY)
 
 ' If ML <= 120 Then ML = 0
 ' If MT <= 120 Then MT = 0
 
 ' If ML > Screen.Width - Me.Width - 120 Then ML = Screen.Width - Me.Width
 ' If MT > Screen.Height - Me.Height - 120 Then MT = Screen.Height - Me.Height
 
 Me.Left = ML
 Me.Top = MT

 If wndTitle.Visible = True Then
    wndTitle.Top = Me.Top - wndTitle.Height
    wndTitle.Left = Me.Left + ((Me.Width - wndTitle.Width) / 2)
 End If
 
 frmBorder.Left = wndMain.Left
 frmBorder.Top = wndMain.Top - frmBorder.Height

End If


End Sub


Private Sub tubMeters_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub


Private Sub Timer1_Timer()

If MMHeader.Mode = 526 Then
  Rot1 = Rot1 + 1
  Rot2 = Rot2 + 1
  
  If Rot1 >= 9 Then Rot1 = 0
  If Rot2 >= 9 Then Rot2 = 0
  If Rot1 < 0 Then Rot1 = 0
  If Rot2 < 0 Then Rot2 = 0

  imgCas1.Picture = picRotate.GraphicCell(Fix(Rot1))
  imgCas2.Picture = picRotate.GraphicCell(Fix(Rot2))
End If

PlayablePos
If lstSec.ListCount = 0 And BackIndex > 0 Then BackIndex = 0

End Sub


