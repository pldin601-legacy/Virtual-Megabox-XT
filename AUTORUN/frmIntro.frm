VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RENNSoft Virtual MegaBox  Version XT 4.01"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox QINFO 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2220
      Left            =   2280
      TabIndex        =   1
      Top             =   1140
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2355
      Left            =   2220
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Image imgLogo 
      Height          =   2325
      Left            =   120
      Picture         =   "frmIntro.frx":030A
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   " Welcome to the RENNSoft Media Player"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3555
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   4080
      Picture         =   "frmIntro.frx":36D6
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
QINFO.AddItem "ʳ���� ��� ��� �� ����� ;-)"
QINFO.AddItem ""
QINFO.AddItem "����� ���������� � ����� �������� ������:"
QINFO.AddItem " 1. ��� �����"
QINFO.AddItem " 2. ��������"
QINFO.AddItem " 3. ������"
QINFO.AddItem ""
QINFO.AddItem "����� �� ��� ���������:"
QINFO.AddItem " - ���������� ����� ))"
QINFO.AddItem " - ������ ���� ����� �� �����"
QINFO.AddItem " - �������� �����"
QINFO.AddItem " - ����������� ���� ��������� � HTML. �������"
QINFO.AddItem " - �� ������������ ��������� ( ��� ;) )"
QINFO.AddItem " - �� ��������� ������������ ����� �� ���������"
QINFO.AddItem "    ����� � ��������"
QINFO.AddItem " �� ������� ������ ����� �� �����"
QINFO.AddItem " � �� ������ ���� ������..."
QINFO.AddItem ""
QINFO.AddItem "������ ������:"
QINFO.AddItem " <F3 > �����������"
QINFO.AddItem " <F2 > ��������"
QINFO.AddItem " <F8 > ������ ����"
QINFO.AddItem " <F5 > �������� �����"
QINFO.AddItem " <DEL> ���������� ����"
QINFO.AddItem " <INS> ������ �����"
QINFO.AddItem " <Z>   ����� ���������"
QINFO.AddItem " <X>   �����"
QINFO.AddItem " <C>   ����� ���������"
QINFO.AddItem " <V>   �����"
QINFO.AddItem " <B>   �� �����"
QINFO.AddItem " <MNU> ���� ����"

End Sub


