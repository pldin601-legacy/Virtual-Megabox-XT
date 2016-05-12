VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form wndMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   5865
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "mainWindow.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "vmbxt"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5865
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   240
      ScaleHeight     =   795
      ScaleWidth      =   3675
      TabIndex        =   23
      Top             =   6120
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
      Begin vmbxt401.cSysTray cSysTray 
         Left            =   1080
         Top             =   0
         _ExtentX        =   900
         _ExtentY        =   900
         InTray          =   0   'False
         TrayIcon        =   "mainWindow.frx":030A
         TrayTip         =   "Virtual MegaBox XT"
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
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
   End
   Begin VB.PictureBox SSPanel1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   2355
      Left            =   60
      ScaleHeight     =   2295
      ScaleWidth      =   10155
      TabIndex        =   19
      Top             =   60
      Width           =   10215
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   1215
         Left            =   5640
         ScaleHeight     =   1155
         ScaleWidth      =   4275
         TabIndex        =   31
         Top             =   60
         Width           =   4335
         Begin vmbxt401.MTrack tmtLen 
            Height          =   375
            Left            =   2520
            TabIndex        =   32
            Top             =   720
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   661
         End
         Begin vmbxt401.MTrack tmtPos 
            Height          =   315
            Left            =   2520
            TabIndex        =   33
            Top             =   180
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
         End
         Begin vmbxt401.MTimer tmLen 
            Height          =   330
            Left            =   1380
            Top             =   720
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin vmbxt401.MTimer tmPos 
            Height          =   330
            Left            =   1380
            Top             =   180
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin vmbxt401.MTimer tmClock 
            Height          =   330
            Left            =   240
            ToolTipText     =   "Wold Time"
            Top             =   180
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin VB.Label lblLedDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MODE"
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
            Index           =   6
            Left            =   3540
            TabIndex        =   50
            Top             =   600
            Width           =   300
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
            Left            =   3540
            TabIndex        =   49
            Top             =   120
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
            Left            =   3000
            TabIndex        =   48
            Top             =   60
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
            Left            =   2580
            TabIndex        =   47
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
            Left            =   1860
            TabIndex        =   46
            Top             =   600
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
            Left            =   1740
            TabIndex        =   45
            Top             =   60
            Width           =   495
         End
         Begin VB.Label lblLedDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIME"
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
            Index           =   0
            Left            =   840
            TabIndex        =   44
            Top             =   60
            Width           =   240
         End
         Begin VB.Image Model 
            Height          =   405
            Left            =   3540
            Picture         =   "mainWindow.frx":075C
            Top             =   720
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
            ForeColor       =   &H00808000&
            Height          =   165
            Left            =   360
            TabIndex        =   38
            Top             =   660
            Width           =   585
         End
         Begin VB.Line REPLine 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            X1              =   300
            X2              =   180
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line REPLine 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   180
            X2              =   180
            Y1              =   720
            Y2              =   960
         End
         Begin VB.Line REPLine 
            BorderColor     =   &H00FFFFFF&
            Index           =   3
            X1              =   1020
            X2              =   1140
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line REPLine 
            BorderColor     =   &H00FFFFFF&
            Index           =   4
            X1              =   1140
            X2              =   1140
            Y1              =   720
            Y2              =   960
         End
         Begin VB.Line REPLine 
            BorderColor     =   &H00FFFFFF&
            Index           =   2
            X1              =   180
            X2              =   300
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line REPLine 
            BorderColor     =   &H00FFFFFF&
            Index           =   5
            X1              =   1140
            X2              =   1020
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
            ForeColor       =   &H00808000&
            Height          =   165
            Left            =   360
            TabIndex        =   37
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
            ForeColor       =   &H00808000&
            Height          =   165
            Left            =   660
            TabIndex        =   36
            Top             =   900
            Width           =   285
         End
         Begin VB.Label ioUp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UP"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   165
            Left            =   3540
            TabIndex        =   35
            Top             =   240
            Width           =   195
         End
         Begin VB.Label ioDown 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DOWN"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   165
            Left            =   3540
            TabIndex        =   34
            Top             =   360
            Width           =   435
         End
      End
      Begin VB.PictureBox VUPANEL 
         BackColor       =   &H00000000&
         Height          =   1215
         Left            =   1800
         ScaleHeight     =   1155
         ScaleWidth      =   3675
         TabIndex        =   27
         Top             =   60
         Width           =   3735
         Begin vmbxt401.QSButton tubMeters 
            Height          =   135
            Left            =   3360
            Top             =   420
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   238
            Caption         =   ""
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
         Begin vmbxt401.VU VU1 
            Height          =   675
            Left            =   420
            TabIndex        =   28
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1191
         End
         Begin vmbxt401.VUMG VUMG1 
            Height          =   975
            Index           =   0
            Left            =   60
            TabIndex        =   40
            Top             =   60
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1720
         End
         Begin VB.PictureBox F_LED 
            BackColor       =   &H00000000&
            Height          =   135
            Left            =   3360
            ScaleHeight     =   75
            ScaleWidth      =   195
            TabIndex        =   29
            Top             =   240
            Width           =   255
         End
         Begin vmbxt401.VUMG VUMG1 
            Height          =   975
            Index           =   1
            Left            =   1680
            TabIndex        =   41
            Top             =   60
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STAT"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   165
            Left            =   3240
            TabIndex        =   30
            Top             =   60
            Width           =   375
         End
      End
      Begin VB.PictureBox panContPan 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   10125
         TabIndex        =   25
         Top             =   1320
         Width           =   10155
         Begin VB.PictureBox scrTotal 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   1800
            ScaleHeight     =   135
            ScaleWidth      =   3735
            TabIndex        =   39
            Top             =   60
            Width           =   3735
         End
         Begin vmbxt401.QSButton btnExplore 
            Height          =   255
            Left            =   9180
            Top             =   660
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   450
            BackColor       =   0
            Caption         =   "Explorer"
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
         Begin vmbxt401.QSButton btnCopy 
            Height          =   255
            Left            =   7920
            Top             =   660
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   450
            BackColor       =   16776960
            Caption         =   "COPY"
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
            ForeColor       =   0
         End
         Begin vmbxt401.QSButton btnKill 
            Height          =   255
            Left            =   5640
            ToolTipText     =   "Kill"
            Top             =   660
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            BackColor       =   255
            Caption         =   "KILL FILE"
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
         Begin vmbxt401.QSButton btnVolume 
            Height          =   315
            Left            =   8520
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            BackColor       =   0
            Caption         =   "VOLUME"
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
            ForeColor       =   8454016
         End
         Begin vmbxt401.QSButton btnMode 
            Height          =   315
            Left            =   7560
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BackColor       =   0
            Caption         =   "MODE"
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
            ForeColor       =   8454016
         End
         Begin vmbxt401.QSButton BtnSetup 
            Height          =   315
            Left            =   6660
            Top             =   360
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            Caption         =   "LANG."
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
         Begin vmbxt401.QSButton btnRepeat 
            Height          =   315
            Left            =   5640
            Top             =   360
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            BackColor       =   0
            Caption         =   "REPEAT"
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
            ForeColor       =   12632256
         End
         Begin vmbxt401.QSImgButton btnStop 
            Height          =   315
            Left            =   4080
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            Picture         =   "mainWindow.frx":11BE
            BackColor       =   0
         End
         Begin vmbxt401.QSImgButton btnPause 
            Height          =   315
            Left            =   3540
            Top             =   600
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            Picture         =   "mainWindow.frx":14B0
            BackColor       =   0
         End
         Begin vmbxt401.QSImgButton btnNext 
            Height          =   315
            Left            =   2940
            Top             =   600
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            Picture         =   "mainWindow.frx":16FA
            BackColor       =   0
         End
         Begin vmbxt401.QSImgButton btnPlay 
            Height          =   315
            Left            =   2400
            Top             =   600
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            Picture         =   "mainWindow.frx":1A58
            BackColor       =   0
         End
         Begin vmbxt401.QSImgButton btnPrev 
            Height          =   315
            Left            =   1800
            ToolTipText     =   "Prev"
            Top             =   600
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            Picture         =   "mainWindow.frx":1CB2
            BackColor       =   0
         End
         Begin vmbxt401.QSButton btnReset 
            Height          =   255
            Left            =   180
            Top             =   60
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   450
            BackColor       =   255
            Caption         =   "Reset"
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
         Begin vmbxt401.QSButton btnWindow 
            Height          =   255
            Left            =   180
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   450
            Caption         =   "Big Window"
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
         Begin vmbxt401.QSButton Command1 
            Height          =   255
            Left            =   960
            Top             =   660
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            Caption         =   "SOFT"
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
         Begin vmbxt401.QSButton btnExit 
            Height          =   255
            Left            =   180
            Top             =   660
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            Caption         =   "Exit"
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
         Begin ComctlLib.Slider Scroller 
            Height          =   315
            Left            =   1800
            TabIndex        =   26
            Top             =   240
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   327682
            BorderStyle     =   1
            Max             =   1
            TickFrequency   =   30
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Quadrosoft Virtual MegaBox XT 4.04"
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
            TabIndex        =   42
            Top             =   60
            Width           =   4335
         End
         Begin VB.Label lblTitleBack 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Quadrosoft Virtual MegaBox XT 4.04"
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
            Left            =   5665
            TabIndex        =   43
            Top             =   85
            Width           =   4335
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MegaBox"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   420
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Virtual"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   0
         Left            =   60
         TabIndex        =   22
         Top             =   120
         Width           =   1455
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   21
         Top             =   780
         Width           =   1755
      End
   End
   Begin vmbxt401.QSButton btnMinimize 
      Height          =   255
      Left            =   10020
      Top             =   5520
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Marlett"
      FontSize        =   8,25
   End
   Begin vmbxt401.QSButton btnHide 
      Height          =   255
      Left            =   9720
      Top             =   5520
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Marlett"
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
      TabIndex        =   11
      Top             =   5520
      Width           =   1875
   End
   Begin VB.PictureBox tblTimer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      Height          =   1155
      Left            =   60
      ScaleHeight     =   1095
      ScaleWidth      =   7635
      TabIndex        =   5
      Top             =   4620
      Width           =   7695
      Begin VB.CheckBox btnTimer 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   540
         TabIndex        =   51
         Top             =   600
         Width           =   195
      End
      Begin VB.PictureBox ledStop 
         BackColor       =   &H00008000&
         Height          =   135
         Left            =   7260
         ScaleHeight     =   75
         ScaleWidth      =   195
         TabIndex        =   12
         Top             =   540
         Width           =   255
      End
      Begin VB.PictureBox ledStart 
         BackColor       =   &H00008000&
         Height          =   135
         Left            =   7260
         ScaleHeight     =   75
         ScaleWidth      =   195
         TabIndex        =   13
         Top             =   180
         Width           =   255
      End
      Begin VB.PictureBox tblTmrIndicator 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   855
         Left            =   1320
         ScaleHeight     =   795
         ScaleWidth      =   3615
         TabIndex        =   8
         Top             =   120
         Width           =   3675
         Begin vmbxt401.MTimer tmrStop 
            Height          =   330
            Left            =   2580
            Top             =   360
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin vmbxt401.MTimer tmrStart 
            Height          =   330
            Left            =   120
            Top             =   360
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin VB.Label lblPro2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Старт         Стоп"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1140
            TabIndex        =   10
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label lblPro1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Программирование таймера :"
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
            Left            =   135
            TabIndex        =   9
            Top             =   60
            Width           =   3345
         End
      End
      Begin VB.HScrollBar scrStart 
         Height          =   255
         Left            =   5100
         Max             =   1440
         TabIndex        =   7
         Top             =   120
         Width           =   2055
      End
      Begin VB.HScrollBar scrStop 
         Height          =   255
         Left            =   5100
         Max             =   1440
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Timer On/Off"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   180
         TabIndex        =   16
         Top             =   120
         Width           =   930
      End
   End
   Begin VB.PictureBox tblListCtrl 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      Height          =   2055
      Left            =   60
      ScaleHeight     =   1995
      ScaleWidth      =   10170
      TabIndex        =   0
      Top             =   2460
      Width           =   10230
      Begin VB.ListBox lstSec 
         BackColor       =   &H00000000&
         Columns         =   1
         ForeColor       =   &H0000FF00&
         Height          =   1860
         Left            =   1800
         OLEDropMode     =   1  'Manual
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   60
         Width           =   5895
      End
      Begin vmbxt401.QSButton btnClear 
         Height          =   315
         Left            =   8040
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         BackColor       =   8421376
         Caption         =   "btnClear"
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
      Begin vmbxt401.QSButton btnSelect 
         Height          =   315
         Left            =   8040
         Top             =   1260
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         BackColor       =   8421376
         Caption         =   "Sel. All"
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
      Begin vmbxt401.QSButton btnDelete 
         Height          =   315
         Left            =   9120
         Top             =   900
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BackColor       =   8421376
         Caption         =   "Delete"
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
      Begin vmbxt401.QSButton btnAdd 
         Height          =   315
         Left            =   8040
         Top             =   900
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BackColor       =   8421376
         Caption         =   "Add..."
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
      Begin vmbxt401.QSButton btnLoad 
         Height          =   315
         Left            =   9000
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BackColor       =   8421376
         Caption         =   "Load..."
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
      Begin vmbxt401.QSButton btnSave 
         Height          =   315
         Left            =   8040
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BackColor       =   8421376
         Caption         =   "Save..."
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
      Begin vmbxt401.QSButton btnFind 
         Height          =   255
         Left            =   8040
         Top             =   300
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         BackColor       =   8421376
         Caption         =   "Find"
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
      Begin vmbxt401.QSButton btnRNDSort 
         Height          =   255
         Left            =   8040
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         BackColor       =   8421376
         Caption         =   "Random sort"
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
      Begin VB.CommandButton gDown 
         Height          =   975
         Left            =   7680
         Picture         =   "mainWindow.frx":2084
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   315
      End
      Begin VB.CommandButton gUp 
         Height          =   915
         Left            =   7680
         Picture         =   "mainWindow.frx":24C6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Width           =   315
      End
      Begin VB.CommandButton btnLN 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "LN"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1395
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
         ForeColor       =   &H0000FF00&
         Height          =   255
         ItemData        =   "mainWindow.frx":2908
         Left            =   120
         List            =   "mainWindow.frx":290A
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   1020
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MegaList"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   1
         Left            =   420
         TabIndex        =   2
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Virtual"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   60
         Width           =   1455
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1635
      End
   End
   Begin PicClip.PictureClip Modes 
      Left            =   7380
      Top             =   180
      _ExtentX        =   820
      _ExtentY        =   7144
      _Version        =   393216
      Rows            =   10
      Picture         =   "mainWindow.frx":290C
   End
   Begin VB.Image tblLogo 
      Height          =   795
      Left            =   7800
      Picture         =   "mainWindow.frx":8E9E
      Stretch         =   -1  'True
      Top             =   4620
      Width           =   2535
   End
   Begin VB.Menu abc 
      Caption         =   "ABC"
      Visible         =   0   'False
      Begin VB.Menu stngr 
         Caption         =   "1. Stinger"
      End
      Begin VB.Menu rmt 
         Caption         =   "2. Remote Control"
      End
      Begin VB.Menu idtaged 
         Caption         =   "3. Launch ""idTAG Editor"""
      End
      Begin VB.Menu renameFile 
         Caption         =   "4. Rename File"
      End
   End
End
Attribute VB_Name = "wndMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ModeSelected As Integer
Dim PlaySelected As Integer
Dim TimeCount As Integer
Dim OldX, OldY
Dim TimerState As String
Dim Tik
Dim Fucka As String, LN As Integer
Dim Ench As Integer
Dim Mozhna As Boolean
Dim UFlag As Integer
Dim DefFile As String

Const ME_FULL = 5925
Const ME_SMPL = 2490

Sub ApplySettings()
SelectPlay Val(GetSetting("vmbxt401", "Program", "PlayMode", "0"))
SelectMode Val(GetSetting("vmbxt401", "Program", "RepMode", "0"))
End Sub

Function CCPU() As Currency
Dim CPU_ST, CPU_ET, CPU_PRC

CPU_ST = Fix(Timer)
CPU_ET = 0
Do:  CPU_ET = CPU_ET + 1: Loop While Not Timer >= CPU_ST + 4

CCPU = Fix(CPU_ET / 4)

End Function

Sub Command_Next()
On Error Resume Next

If lstMain.ListIndex < lstMain.ListCount - 1 Then
 lstMain.ListIndex = lstMain.ListIndex + 1
 lstSec.ListIndex = lstSec.ListIndex + 1
 btnPlay_Click
End If


End Sub

Sub Command_Pause()
MMHeader.Command = "Pause"
End Sub

Sub Command_Play()
On Error Resume Next

 If lstMain.ListIndex >= 0 Then
 
  Scroller.Visible = True

  wpPlayer.seeker.Visible = True

  F_LED.BackColor = RGB(255, 255, 0)
  
  MMHeader.Command = "Close"
  MMHeader.Filename = lstMain.List(lstSec.ListIndex)

  lblTitle.Caption = GetMp3Song(MMHeader.Filename)
  lblTitleBack.Caption = GetMp3Song(MMHeader.Filename)

  wpPlayer.lblTrackTitle.Caption = GetMp3Song(MMHeader.Filename)
  wpPlayer.Caption = GetMp3Song(MMHeader.Filename)

 
  MMHeader.Command = "Open"

  wndTitle.lab(0).Caption = GetMp3Song(MMHeader.Filename) + " "
  
  MMHeader.Command = "Play"
  tmtLen.Track = lstMain.ListCount
  tmtPos.Track = lstMain.ListIndex + 1

  wpPlayer.TRK.Caption = Format(lstMain.ListIndex + 1)

  TotalChange lstMain.ListCount, lstMain.ListIndex
  If lstMain.ListIndex = lstMain.ListCount - 1 Then tmtPos.LastRec
  If MMHeader.Error Then tmtPos.ErrorFound: lblTitle.Caption = "(911) MCI ERROR: FILE FORMAT UNSUPPORTED! ": lblTitleBack.Caption = "(911) MCI ERROR: FILE FORMAT UNSUPPORTED! "
  F_LED.BackColor = RGB(0, 0, 0)
  If Err Then MsgBox "ERR" + Format(Err.Number) + ": " + Err.Description, vbCritical
  
  MMHeader_StatusUpdate
    
 End If

End Sub

Sub Command_Prev()
On Error Resume Next

If lstMain.ListIndex > 0 Then
 lstMain.ListIndex = lstMain.ListIndex - 1
 lstSec.ListIndex = lstSec.ListIndex - 1
 btnPlay_Click
End If


End Sub

Sub Facing()
Rem 'Lang32 Loading'
lblVersion(0).Caption = Language("VERSION:") + " " + GetVersion
lblVersion(1).Caption = Language("VERSION:") + " " + GetVersion
btnAdd.Caption = Language("BTNADD:")
btnAdd.ToolTipText = Language("BTNADDTIP:")
btnClear.Caption = Language("BTNCLEAR:")
btnClear.ToolTipText = Language("BTNCLEARTIP:")
btnCopy.Caption = Language("BTNCOPY:")
btnCopy.ToolTipText = Language("BTNCOPYTIP:")
btnDelete.Caption = Language("BTNDELETE:")
btnDelete.ToolTipText = Language("BTNDELETETIP:")
btnExit.Caption = Language("BTNEXIT:")
btnExit.ToolTipText = Language("BTNEXITTIP:")
btnFind.Caption = Language("BTNFIND:")
btnFind.ToolTipText = Language("BTNFINDTIP:")
btnKill.Caption = Language("BTNKILL:")
btnKill.ToolTipText = Language("BTNKILLTIP:")
btnLoad.Caption = Language("BTNLOAD:")
btnLoad.ToolTipText = Language("BTNLOADTIP:")
btnMode.Caption = Language("BTNMODE:")
btnMode.ToolTipText = Language("BTNMODETIP:")
btnRepeat.Caption = Language("BTNREPEAT:")
btnRepeat.ToolTipText = Language("BTNREPEATTIP:")
btnReset.Caption = Language("BTNRESET:")
btnReset.ToolTipText = Language("BTNRESETTIP:")
btnSave.Caption = Language("BTNSAVE:")
btnSave.ToolTipText = Language("BTNSAVETIP:")
btnSelect.Caption = Language("BTNSELECT:")
btnSelect.ToolTipText = Language("BTNSELECTTIP:")
btnVolume.Caption = Language("BTNVOLUME:")
btnVolume.ToolTipText = Language("BTNVOLUMETIP:")
btnWindow.Caption = Language("BTNWINDOW:")
btnWindow.ToolTipText = Language("BTNWINDOWTIP:")
Label3.Caption = Language("ON/OFF:")
Label3.ToolTipText = Language("ON/OFFTIP:")
lblPro1.Caption = Language("LBLPROGTIMER:")
lblPro1.ToolTipText = Language("LBLPROGTIMERTIP:")
btnPrev.ToolTipText = Language("BTNPREVTIP:")
btnPlay.ToolTipText = Language("BTNPLAYTIP:")
btnNext.ToolTipText = Language("BTNNEXTTIP:")
btnPause.ToolTipText = Language("BTNPAUSETIP:")
btnStop.ToolTipText = Language("BTNSTOPTIP:")

End Sub

Sub LoadList(ByRef ListFileName As String, Destination As ListBox, MainForm As Form)
' On Error Resume Next

' Definitions
Dim FN As String, Mired As Boolean, SP As String, Lines As Long
Dim ComaData As String, ComaValue As String, Misses As Boolean, AutoPlay As Boolean, SeekPlay As Currency

' Load File
DoEvents
Open ListFileName For Input As #1
If Err Then MsgBox "Error! ''" + ListFileName + "'' " + Language("NO_ACCESS:"), vbCritical: Exit Sub
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
  If Err Then MsgBox "Some error found!" + Chr$(13) + Chr$(13) + "Line " + Format$(Lines, "0"), vbInformation: Err = 0: Misses = True
  
  ' Command found, then
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#FILE     :" Then
   Err = 0
   SP = ComaValue
   Destination.AddItem LowPath(PathHead(SP)) + FileHead(SP)
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''FILE'' is used incorrect!", vbInformation: Misses = True
   Me.lstSec.AddItem GetMp3Song(LowPath(PathHead(SP)) + FileHead(SP)): Err = 0
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
   MainForm.scrStart.Value = Val(ComaValue)
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''START'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If ComaData = "#STOP     :" Then
   MainForm.scrStop.Value = Val(ComaValue)
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
   DefFile = Val(ComaValue)
   If Err Then MsgBox "Error!" + Chr$(13) + "Line " + Format$(Lines, "0") + Chr$(13) + Chr$(13) + "Function ''DEFAULT'' is used incorrect!", vbInformation: Err = 0: Misses = True
  End If
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  ' Reset Ifs
  ComaData = ""
  ComaValue = ""
  Mired = True
End If


If Mired = False Then Destination.AddItem FN: Me.lstSec.AddItem GetMp3Song(FN): Destination.Selected(Destination.ListCount - 1) = True: Me.lstSec.Selected(Destination.ListCount - 1) = True
Mired = False
Loop While Not EOF(1)

F_LED.BackColor = RGB(0, 0, 0)

If Misses Then MsgBox Language("LOAD_ERROR:"), vbInformation

If Err = 0 Then Beep

Close #1

If AutoPlay = True Then
  btnPlay_Click
  MMHeader.To = SeekPlay
  MMHeader.Command = "Seek"
  MMHeader.Command = "Play"
End If

End Sub


Sub ResetStatus()

 picCPU.Cls
 If Ench > 19232 Then picCPU.Scale (0, 0)-(Ench + 1, 2)
 If Ench <= 19232 Then picCPU.Scale (0, 0)-(19232, 2)

 picCPU.FontName = "Small Fonts"
 picCPU.FontSize = "5"
 picCPU.Line (0, 0)-(19232, 0.9), RGB(0, 0, 255), BF
 picCPU.Line (0, 1)-(Ench, 1.9), RGB(255, 255, 0), BF
 picCPU.ToolTipText = "Your system working up to " + Format((100 / 19232 * Ench) - 100, "0") + "% higher of my system"
' picCPU.ToolTipText = Ench
 
 picCPU.ForeColor = RGB(255, 255, 0)
 picCPU.CurrentX = 0: picCPU.CurrentY = 0: picCPU.Print "CEL500/INTEL810/64RAM/1024HDD"
 
 picCPU.ForeColor = RGB(0, 0, 255)
 picCPU.CurrentX = 0: picCPU.CurrentY = 1: picCPU.Print "ВАША  СИСТЕМА"

 Me.tmtLen.Waiting
 Me.tmtPos.Waiting
 Me.tmLen.Off
 Me.tmPos.Off

 wpPlayer.TimeX.Off
 wpPlayer.TRK.Caption = "0"
 wpPlayer.TOTL.Caption = "0"

 lblTitle.Caption = ""
 lblTitleBack.Caption = ""


 Scroller.Max = 1
 Scroller.Min = 0
 Scroller.Visible = False

 wpPlayer.seeker.Max = 1
 wpPlayer.seeker.Value = 0
 wpPlayer.seeker.Visible = False

 MMHeader.Command = "Close"
 MMHeader.Filename = ""

 TotalChange 1, 1
 F_LED.BackColor = RGB(0, 0, 0)

 wndTitle.lab(0).Caption = "CRD 4.07.2001. Quadrosoft. "

End Sub

Function SaveList(ListFile As String, Source As ListBox) As Boolean
' Standart
On Error Resume Next
Dim I, J

I = FreeFile
Open ListFile For Output As I
If Err Then Close I: SaveList = False: Exit Function

F_LED.BackColor = RGB(255, 128, 0)
 ' Some Comments
 Print #I, " Quadrosoft (R) Virtual MegaBox (TM) eXTended version Playlist File"
 Print #I, " Copyright (C) 2000-" + Format(Year(Now)) + "Quadrosoft Minicorporation. All Rights reserved."
 Print #I, " File format: MXT-File. Created " + Format$(Now, "dd.mm.yyyy") + " at " + Format$(Now, "hh:mm")
 Print #I, ""
 Print #I, ""
 Print #I, "#DEFAULT  : " + ListFile
 
 
' Files for add
For J = 0 To Source.ListCount - 1
 Print #I, "#FILE     : " + Source.List(J)
 Print #I, "#TITLE    : " + Me.lstSec.List(J)
Next
  
  Print #I, ""
' Files for select
For J = 0 To Source.ListCount - 1
 If Source.Selected(J) = True Then
   Print #I, "#SELECT   : " + Str(J)
 End If
Next

 ' Miscs
 Print #I, ""
 If MMHeader.Mode = 526 Then
  Print #I, "#AUTOPLAY : DO IT NOW!!! FROM SEEK POSITION."
  Print #I, "#SEEK     : " + Format(MMHeader.Position, "000 000 000 000.000") + " MSeconds"
 End If
 
 Print #I, "#TIMER    : " + Str$(CInt(btnTimer.Value))
 Print #I, "#START    : " + Str$(scrStart.Value)
 Print #I, "#STOP     : " + Str$(scrStop.Value)
 Print #I, ""
 Print #I, "#DIRECTION: " + Str$(PlaySelected)
 Print #I, "#FOCUS    : " + Str(lstMain.ListIndex)
 Print #I, ""
 Print #I, " End of playlist file --^--"
 Print #1, " Creater  : RMK"
 Print #1, " Tester   : EPSELON"
 Print #1, " PrgValue : " & App.Revision
 
Close I

F_LED.BackColor = RGB(0, 0, 0)
If Err = 0 Then SaveList = True

End Function

Sub Scr_Down()
MMHeader.UpdateInterval = 0
RealUpdate.Enabled = False
End Sub

Sub Scr_Scroll()
Dim aMins, aSecs

   aMins = Fix(Scroller.Value / 60)
   aSecs = Scroller.Value Mod 60
   tmPos.TimeSet = Format$(aMins, "00") + ":" + Format$(aSecs, "00")
   
   If Scroller.Value < MMHeader.Position / CCur(1000) Then
    Moding 617
   Else
    Moding 618
   End If

   If wpPlayer.Visible = True Then
     wpPlayer.TimeX.TimeSet = Format$(aMins, "00") + ":" + Format$(aSecs, "00")
   End If

End Sub

Sub Scr_Up()
Dim FUCK As Currency
MMHeader.To = CCur(Scroller.Value) * CCur(1000)
MMHeader.Command = "Seek"
MMHeader.Command = "Play"
MMHeader.UpdateInterval = 1000
RealUpdate.Enabled = True


End Sub

Sub SelectMode(Mode As Integer)

Select Case Mode
Case 0
 ioREPEAT.ForeColor = RGB(50, 50, 50)
 ioTRK.ForeColor = RGB(50, 50, 50)
 ioALL.ForeColor = RGB(50, 50, 50)
 ModeSelected = Mode
 
Case 1
 ioREPEAT.ForeColor = RGB(250, 250, 250)
 ioTRK.ForeColor = RGB(50, 50, 50)
 ioALL.ForeColor = RGB(250, 250, 250)
 ModeSelected = Mode
 
Case 2
 ioREPEAT.ForeColor = RGB(250, 250, 250)
 ioTRK.ForeColor = RGB(250, 250, 250)
 ioALL.ForeColor = RGB(50, 50, 50)
 ModeSelected = Mode
End Select

End Sub

Sub SelectPlay(Mode As Integer)

Select Case Mode
Case 0
 ioUp.ForeColor = RGB(50, 50, 50)
 ioDown.ForeColor = RGB(250, 250, 250)
 PlaySelected = Mode
 
Case 1
 ioUp.ForeColor = RGB(250, 250, 250)
 ioDown.ForeColor = RGB(50, 50, 50)
 PlaySelected = Mode
End Select

End Sub

Sub TotalChange(Max As Integer, Min As Integer)
Dim MOCbKA As Integer
scrTotal.Scale (0, 0)-(Max, 1)
For MOCbKA = 0 To Max
 scrTotal.Line (MOCbKA, 0)-(MOCbKA + 1, 1), RGB(0, 255 / Max * MOCbKA, 255), BF
Next

For MOCbKA = 0 To Min
 scrTotal.Line (MOCbKA, 0)-(MOCbKA + 1, 1), RGB(255 / Max * MOCbKA, 255, 0), BF
Next

End Sub

Private Sub btnAdd_Click()
wndAddFiles.Show 0, Me
End Sub

Private Sub btnClear_Click()
Dim Q
Q = MsgBox(Language("CLEAR_PLAYLIST:"), vbQuestion + vbYesNo, "Quadrosoft Virtual MegaBox. Version XT " + GetVersion)
If Q = 6 Then lstMain.Clear: lstSec.Clear
End Sub

Private Sub btnCopy_Click()
 dlgCopyFile.Show 1, Me
End Sub

Private Sub btnDelete_Click()
On Error Resume Next
lstMain.RemoveItem (lstMain.ListIndex)
lstSec.RemoveItem (lstSec.ListIndex)
End Sub


Private Sub btnExit_Click()
Dim retval
retval = MsgBox(Language("EXIT_PROGRAMM:"), vbQuestion + vbYesNo, "Virtual MegaBox XT " + GetVersion)
If retval = 6 Then
  SaveList LowPath(App.Path) + "default.mxt", wndMain.lstMain
  Unload Me
  End
End If
End Sub

Private Sub btnExplore_Click()
 On Error Resume Next
 Dim SndVolume As Double
 SndVolume = Shell("explorer.exe", vbNormalFocus)
End Sub

Private Sub btnFind_Click()
Dim Q As String, a As Integer, zopa
If lstMain.ListIndex > lstMain.ListCount - 4 Then Exit Sub

Q = InputBox(Language("FIND_PATTERN:"))
If Q = "" Then Exit Sub

For a = lstMain.ListIndex + 1 To lstMain.ListCount - 1
 If InStr(1, UCase(lstMain.List(a)), UCase(Q)) > 0 Then
  zopa = MsgBox(Language("FINDING_HA:") + Chr$(13) + "''" + lstMain.List(a) + "'' ?", vbQuestion + vbYesNoCancel)
  If zopa = 6 Then zopa = 0: lstSec.ListIndex = a: Exit Sub
  If zopa = 2 Then Exit Sub
 End If
Next

MsgBox Language("FINDING_HE:"), vbInformation, ":-("

End Sub

Private Sub btnHide_Click()
 cSysTray.InTray = True
 Me.Hide
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 119 Then btnKill_Click
If KeyCode = 113 Then btnSave_Click
If KeyCode = 114 Then btnLoad_Click
If KeyCode = 27 Then btnExit_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveList LowPath(App.Path) + "default.mxt", wndMain.lstMain
' Save Programm settings
SaveSetting "vmbxt401", "Program", "PlayMode", Format(PlaySelected)
SaveSetting "vmbxt401", "Program", "RepMode", Format(ModeSelected)


End Sub

Private Sub gDown_Click()
On Error Resume Next
If lstMain.ListIndex < lstMain.ListCount - 1 Then
 ExchangeFiles lstMain.ListIndex, lstMain.ListIndex + 1, Me.lstMain
 ExchangeFiles lstSec.ListIndex, lstSec.ListIndex + 1, Me.lstSec
 lstMain.ListIndex = lstMain.ListIndex + 1
 lstSec.ListIndex = lstSec.ListIndex + 1
End If
End Sub


Private Sub gUp_Click()
On Error Resume Next
If lstMain.ListIndex > 0 Then
 ExchangeFiles lstMain.ListIndex, lstMain.ListIndex - 1, Me.lstMain
 ExchangeFiles lstSec.ListIndex, lstSec.ListIndex - 1, Me.lstSec
 lstMain.ListIndex = lstMain.ListIndex - 1
 lstSec.ListIndex = lstSec.ListIndex - 1
End If
End Sub

Private Sub idtaged_Click()
On Error Resume Next
Shell LowPath(App.Path) + "PLUGINS\idtag.exe " + lstMain.List(lstSec.ListIndex), vbNormalFocus
If Err Then MsgBox "Ошибка! idTAG Editor не установлен!", vbCritical
End Sub


Private Sub Label5_Click()

End Sub

Private Sub lstSec_Click()
On Error Resume Next
lstMain.ListIndex = lstSec.ListIndex
lstMain.Selected(lstSec.ListIndex) = lstSec.Selected(lstSec.ListIndex)
tmtLen.Track = lstSec.ListCount
End Sub


Private Sub lstSec_DblClick()
On Error Resume Next
lstMain.ListIndex = lstSec.ListIndex
lstMain.Selected(lstSec.ListIndex) = lstSec.Selected(lstSec.ListIndex)
End Sub


Private Sub lstSec_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And lstMain.ListIndex > -1 Then infForm.oFileName = lstMain.List(lstMain.ListIndex): infForm.Show 1, Me
End Sub


Private Sub lstSec_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim blt As String
For Bt = 1 To Data.Files.Count
blt = Data.Files(Bt)

lstMain.AddItem blt
lstMain.Selected(lstMain.ListCount - 1) = True

lstSec.AddItem GetMp3Song(blt)
lstSec.Selected(lstSec.ListCount - 1) = True

Next


End Sub

Private Sub renameFile_Click()
On Error Resume Next
Dim UserDest As String, FullDst As String, FullSrc As String
UserDest = InputBox("Rename file " + lstMain.List(lstSec.ListIndex) + " to:", "Rename File")
If UserDest = "" Then Exit Sub
MMHeader.Command = "Close"
FullSrc = lstMain.List(lstSec.ListIndex)
FullDst = LowPath(PathHead(lstMain.List(lstSec.ListIndex))) + UserDest
Name FullSrc As FullDst

If Err Then MsgBox "Error renaming!!!", vbCritical
If Err = 0 Then lstMain.List(lstSec.ListIndex) = FullDst
If Err = 0 Then lstSec.List(lstSec.ListIndex) = "[Renamed to " + UserDest + "]"

End Sub

Private Sub rmt_Click()
wpPlayer.Show 0, Me
Me.Hide
End Sub

Private Sub Scroller_Scroll()
Scr_Scroll

End Sub

Private Sub stngr_Click()
On Error Resume Next
Shell LowPath(App.Path) + "PLUGINS\stinger.exe", vbNormalFocus
If Err Then MsgBox "Ошибка! Stinger не установлен!", vbCritical
End Sub

Private Sub tbllogo_Click()
Form1.Show

End Sub

Private Sub btnKill_Click()
On Error Resume Next
Dim retval

retval = MsgBox(Language("SURE_DELETE:") + " ''" + lstMain.List(lstMain.ListIndex) + "''?", vbExclamation + vbYesNo, "Delete File From Disk")
If retval = 6 Then
   ResetStatus
   Kill lstMain.List(lstSec.ListIndex)
    If Err Then
     MsgBox Language("ERROR_DELETE:") + Chr(13) + Chr$(13) + lstMain.List(lstMain.ListIndex), vbCritical
    Else
     lstMain.RemoveItem lstSec.ListIndex
     lstSec.RemoveItem lstSec.ListIndex
    End If
End If

End Sub

Private Sub btnLN_Click()
On Error Resume Next
Dim Rv
Rv = InputBox("Enter LN", , btnLN.Caption)
LN = Rv
btnLN.Caption = Str(LN)

End Sub

Private Sub btnMinimize_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub btnNext_Click()
Call Command_Next
End Sub

Private Sub btnPause_Click()
Call Command_Pause
End Sub


Private Sub btnPlay_Click()
Call Command_Play
End Sub

Private Sub btnPrev_Click()
Call Command_Prev
End Sub

Private Sub btnReset_Click()
Dim retval, Q
retval = MsgBox(Language("SURE_RESET:"), vbQuestion + vbYesNo, "Virtual MegaBox XT " + GetVersion)

If retval = 6 Then
frmSplash.Show
Unload wndMain
End If

End Sub

Private Sub btnRNDSort_Click()
Dim Z As Integer, y As Currency
On Error Resume Next

If lstMain.ListCount = 0 Then Exit Sub

For Z = 0 To lstSec.ListCount - 1
 y = Rnd
 ExchangeFiles Z, Fix(y * lstMain.ListCount - 1), wndMain.lstMain
 ExchangeFiles Z, Fix(y * lstSec.ListCount - 1), wndMain.lstSec
Next

End Sub

Private Sub btnSelect_Click()
Dim K
For K = 0 To lstMain.ListCount - 1
lstMain.Selected(K) = True
lstSec.Selected(K) = True
Next
End Sub

Private Sub btnSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then
 Dim K
 For K = 0 To lstMain.ListCount - 1
  lstMain.Selected(K) = False
 Next
End If

End Sub


Private Sub btnSetup_Click()
LangForm.Show 1, Me
Facing
End Sub

Private Sub btnStop_Click()
ResetStatus
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

Private Sub Command1_Click()
PopupMenu abc
End Sub

Private Sub cSysTray_MouseDown(Button As Integer, Id As Long)
 Me.Show
 cSysTray.InTray = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim KeyCode

KeyCode = Asc(UCase(Chr$(KeyAscii)))

If Chr(KeyCode) = "Z" Then btnPrev_Click
If Chr(KeyCode) = "C" Then btnNext_Click
If Chr(KeyCode) = "X" Then btnPlay_Click
If Chr(KeyCode) = "V" Then btnPause_Click
If Chr(KeyCode) = "B" Then btnStop_Click
If Chr(KeyCode) = "D" Then btnDelete_Click
If Chr(KeyCode) = "L" Then btnClear_Click

Fucka = Fucka + Chr$(KeyCode)
Fucka = Right$(Fucka, 4)


If Fucka = "$-40" Then MsgBox "CONGRATULATIONS-> You Got The Cheat for LN!", vbExclamation: _
   btnLN.Visible = True: _
   Fucka = ""
   
If Fucka = "+486" Then MsgBox "CONGRATULATIONS-> You Got The Cheat for ENCH!", vbExclamation: _
   Ench = 1: Fucka = ""

If Fucka = "2001" Then
   MMHeader.DeviceType = InputBox("Enter the Custom Device Type", "(!WARNING!) Reprogramming (!WARNING!)", MMHeader.DeviceType)
End If
   
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
OldX = x
OldY = y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim MoveX As Long
Dim MoveY As Long

If Button = 1 Then
    MoveX = x - OldX
    MoveY = y - OldY
    Me.Move Me.Left + MoveX, Me.Top + MoveY
End If

If wndTitle.Visible = True Then
    wndTitle.Top = Me.Top - wndTitle.Height
    wndTitle.Left = Me.Left + ((Me.Width - wndTitle.Width) / 2)
End If

End Sub

Private Sub btnLoad_Click()
On Error Resume Next
dlgOpenSave.Filename = DefFile
dlgOpenSave.ShowOpen
If Err Then Exit Sub
LoadList dlgOpenSave.Filename, Me.lstMain, Me
End Sub

Private Sub btnMode_Click()
SelectPlay (PlaySelected + 1) Mod 2
End Sub

Private Sub btnRepeat_Click()
SelectMode ((ModeSelected + 1) Mod 3)
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

Facing

Ench = CCPU

scrStart_Change
scrStop_Change

' Load defaults
SelectMode (0)
SelectPlay (0)
ResetStatus
TotalChange 255, 0
' Welcome title
wndTitle.lab(0) = " WELCOME TO THE VERSION XT " + GetVersion


If Command$ > "" Then
 If LCase(Right(Command$, 3)) = "mxt" Or LCase(Right(Command$, 3)) = "vmb" Then
   LoadList Command$, Me.lstMain, Me
 Else
   lstMain.AddItem Command$
   lstMain.ListIndex = 0
   btnPlay_Click
 End If
Else
 If FileExists(LowPath(App.Path) + "default.mxt") Then LoadList LowPath(App.Path) + "default.mxt", Me.lstMain, Me
End If

ApplySettings
On Error Resume Next
frmTip.Show vbModal, Me

' On Meters
' VU1.VU_ON
' VUMG1(0).VU_ON
' VUMG1(1).VU_ON


End Sub

Private Sub Form_Resize()
Dim y, MY

MY = tblListCtrl.Height / Screen.TwipsPerPixelY
For y = 0 To MY
tblListCtrl.Line (0, y * Screen.TwipsPerPixelY)-(tblListCtrl.Width, y * Screen.TwipsPerPixelY), RGB((100 / MY * y), (100 / MY * y), (100 / MY * y))
Next

MY = tblTimer.Height / Screen.TwipsPerPixelY
For y = 0 To MY
tblTimer.Line (0, y * Screen.TwipsPerPixelY)-(tblTimer.Width, y * Screen.TwipsPerPixelY), RGB((100 / MY * y), (100 / MY * y), (100 / MY * y))
Next

MY = (SSPanel1.Width + SSPanel1.Height) / Screen.TwipsPerPixelY
For y = 0 To MY
SSPanel1.Line (0, y * Screen.TwipsPerPixelY)-(y * Screen.TwipsPerPixelX, 0), RGB(100 + (100 / MY * y), 100 + (100 / MY * y), 100 + (100 / MY * y))
Next

MY = Fix(wndMain.Height / Screen.TwipsPerPixelY)
For y = 0 To MY
Line (0, (y * Screen.TwipsPerPixelY))-(wndMain.Width, (y * Screen.TwipsPerPixelY)), RGB(100 - (100 / MY * y), 100 - (100 / MY * y), 100 - (100 / MY * y))
Next

Me.Line (0, 0)-(Me.Width, 0), RGB(255, 255, 255)
Me.Line (0, Me.Height)-(Me.Width, Me.Height), RGB(155, 155, 155)
Me.Line (Me.Width, 0)-(Me.Width, Me.Height), RGB(255, 255, 255)
Me.Line (0, 0)-(0, Me.Height), RGB(155, 155, 155)

Me.Line (15, 15)-(Me.Width - 15, 15), RGB(200, 200, 200)
Me.Line (15, Me.Height - 15)-(Me.Width - 15, Me.Height - 15), RGB(100, 100, 100)
Me.Line (Me.Width - 15, 15)-(Me.Width - 15, Me.Height - 15), RGB(200, 200, 200)
Me.Line (15, 15)-(15, Me.Height - 15), RGB(100, 100, 100)

panContPan.Scale (0, 0)-(640, 64)
Dim a As Integer, CL As Integer
  For a = 0 To 64
    CL = 88 + (a * Sin(a))
    panContPan.Line (0, a)-(640, a), RGB(CL, CL, CL), BF
  Next
End Sub

Private Sub gStart_Timer()

If GetTimeFromMinutes(scrStart.Value - 1) = Format$(Now, "hh:mm") Then
 btnPlay_Click
 gStart.Enabled = False
 ledStart.BackColor = RGB(0, 64, 0)
 tmrStart.Reset
 scrStart.Value = 0
End If

If GetTimeFromMinutes(scrStart.Value) = "00:00" Then
  gStart.Enabled = False
  ledStart.BackColor = RGB(0, 64, 0)
  tmrStart.Reset
End If

If gStart.Enabled = False And gStop.Enabled = False Then btnTimer.Value = 0

End Sub

Private Sub gStop_Timer()
 

If GetTimeFromMinutes(scrStop.Value - 1) = Format$(Now, "hh:mm") Then
  btnStop_Click
  gStop.Enabled = False
  ledStop.BackColor = RGB(0, 64, 0)
  tmrStop.Reset
  scrStop.Value = 0
End If
 
If GetTimeFromMinutes(scrStop.Value) = "00:00" Then
  gStop.Enabled = False
  ledStop.BackColor = RGB(0, 64, 0)
  tmrStop.Reset
End If
 
 
If gStart.Enabled = False And gStop.Enabled = False Then btnTimer.Value = 0

End Sub


Private Sub MMHeader_Done(NotifyCode As Integer)
On Error Resume Next

If NotifyCode = 1 Then

  
  ' -----------ALL MODE REPEAT----------------
    ' If Selected "All" mode repeat
    If ModeSelected = 1 Then
   
     ' If play direction is "Down"
     If PlaySelected = 0 Then
      If lstMain.ListIndex = lstMain.ListCount - 1 And _
      lstMain.ListIndex >= 0 Then lstMain.ListIndex = 0: btnPlay_Click: Exit Sub
     End If
   
     ' If play direction is "Up"
     If PlaySelected = 1 Then
      If lstMain.ListIndex = 0 And _
      lstMain.ListCount > 1 Then lstMain.ListIndex = lstMain.ListCount - 1: btnPlay_Click: Exit Sub
     End If
   
    Exit Sub
    End If
  ' --------------- END ALL ------------------
  
  
  ' -----------TRK MODE REPEAT----------------
  ' If Selected "Trk" mode repeat
  If ModeSelected = 2 Then
    btnPlay_Click
    Exit Sub
  End If
  ' --------------- END TRK ------------------


  ' ---------------NO REPEAT------------------
  If ModeSelected = 0 Then
    ResetStatus
    ' -------------DIRECTION IS DOWN----------
    If PlaySelected = 0 Then
      btnNext_Click
    End If
    ' ----------------END DOWN----------------
    ' --------------DIRECTION IS UP-----------
    If PlaySelected = 1 Then
      btnPrev_Click
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
If wpPlayer.Visible = True Then wpPlayer.RTMODE.Picture = Model.Picture
End Sub

Private Sub MMHeader_StatusUpdate()
' Standart
On Error Resume Next
If UFlag = 1 Then Exit Sub
F_LED.BackColor = RGB(0, 255, 0)
MMHeader.TimeFormat = 0
Dim aSecs, aMins, aSecls, bSecs, bMins, bSecls
Dim STimer, GX

' Set the length and position as seconds
aSecls = Fix(MMHeader.Length / 1000)
bSecls = Fix(MMHeader.Position / 1000)

' DePlaying mode
If bSecls >= aSecls - LN Then
 MMHeader.Command = "Stop"
 MMHeader_Done (1)
 Exit Sub
End If

' Prepare for update indicators
aMins = Fix(aSecls / 60)
bMins = Fix(bSecls / 60)

aSecs = aSecls Mod 60
bSecs = bSecls Mod 60

If aMins <= 59 Then
' Update indicators
tmPos.TimeSet = Format$(bMins, "00") + ":" + Format$(bSecs, "00")
tmLen.TimeSet = Format$(aMins, "00") + ":" + Format$(aSecs, "00")

wpPlayer.TimeX.TimeSet = Format$(bMins, "00") + ":" + Format$(bSecs, "00")
wpPlayer.TOTL.Caption = Format$(aMins, "00") + ":" + Format$(aSecs, "00")

Else
' Update indicators
tmPos.TimeSet = Format$(Fix(bMins / 60), "00") + ":" + Format$(bMins Mod 60, "00")
tmLen.TimeSet = Format$(Fix(aMins / 60), "00") + ":" + Format$(aMins Mod 60, "00")

wpPlayer.TimeX.TimeSet = Format$(Fix(bMins / 60), "00") + ":" + Format$(bMins Mod 60, "00")
wpPlayer.TOTL.Caption = Format$(Fix(aMins / 60), "00") + ":" + Format$(aMins Mod 60, "00")

End If

' Me.Caption = lblTitle.Caption + " [" + Format$(bMins, "00") + ":" + Format$(bSecs, "00") + "]"


' Scroller update
If aSecls > 0 Then
 Scroller.Max = aSecls
 Scroller.Value = bSecls

   wpPlayer.seeker.Max = aSecls
   wpPlayer.seeker.Value = bSecls

End If

If Ench > 2000 Then
UFlag = 1
 For GX = 255 To 0 Step -0.05
  F_LED.BackColor = RGB(0, GX, 0)
  DoEvents
 Next GX
End If
UFlag = 0

If Ench <= 2000 Then F_LED.BackColor = RGB(0, 0, 0)

End Sub

Private Sub RealUpdate_Timer()
Moding (MMHeader.Mode)

If MMHeader.Mode = 526 Or MMHeader.Mode = 529 Then
 If Second(Now) Mod 10 >= 5 Then tmClock.TimeSet = Format$(Now, "hh:mm")
 If Second(Now) Mod 10 < 5 Then tmClock.TimeSet2 = Format$(Now, "dd.mm")
Else
 If Fix(Timer * 10) Mod 10 >= 5 Then tmClock.TimeSet = Format(Now, "hh:mm")
 If Fix(Timer * 10) Mod 10 < 5 Then tmClock.TimeSet2 = Format(Now, "hh:mm")
 tmPos.TimeSet2 = Format(Now, "dd.mm")
 tmtPos.TrackX = Year(Now)
End If


End Sub


Private Sub Scroller_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Scr_Down
End Sub

Private Sub Scroller_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 Scr_Up
End Sub

Private Sub scrStart_Change()
If scrStart.Value > 0 Then tmrStart.TimeSet = GetTimeFromMinutes(scrStart.Value - 1)
If scrStart.Value = 0 Then tmrStart.Reset
End Sub


Private Sub scrStart_Scroll()
scrStart_Change
End Sub


Private Sub scrStop_Change()
If scrStop.Value > 0 Then tmrStop.TimeSet = GetTimeFromMinutes(scrStop.Value - 1)
If scrStop.Value = 0 Then tmrStop.Reset
End Sub


Private Sub scrStop_Scroll()
scrStop_Change
End Sub


Private Sub SSPanel1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
OldX = x
OldY = y


End Sub


Private Sub SSPanel1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim MoveX As Long
Dim MoveY As Long

If Button = 1 Then
    MoveX = x - OldX
    MoveY = y - OldY
    Me.Move Me.Left + MoveX, Me.Top + MoveY
End If

If wndTitle.Visible = True Then
    wndTitle.Top = Me.Top - wndTitle.Height
    wndTitle.Left = Me.Left + ((Me.Width - wndTitle.Width) / 2)
End If
End Sub



Private Sub tubMeters_Click()
VUMG1(0).Visible = Not VUMG1(0).Visible
VUMG1(1).Visible = Not VUMG1(1).Visible
VU1.Visible = Not VU1.Visible
End Sub


