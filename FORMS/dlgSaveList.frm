VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form dlgSaveList 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������� ������ (*.vmb *.mxt)"
   ClientHeight    =   3225
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6510
   Icon            =   "dlgSaveList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   120
      Width           =   2355
   End
   Begin ComctlLib.StatusBar lblStatus 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   2955
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   476
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "�������"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.DriveListBox lstDrive 
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.DirListBox lstPath 
      Height          =   2340
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.FileListBox lstFiles 
      BackColor       =   &H00FFFFFF&
      Height          =   2430
      Left            =   120
      Pattern         =   "*.vmb;*.mxt"
      TabIndex        =   2
      Top             =   420
      Width           =   2355
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "������"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "dlgSaveList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
lblFileName.Text = ""
dlgLoadList.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub lstDrive_Change()
On Error Resume Next
lstPath.path = lstDrive.Drive
If Err Then MsgBox "������ ��������� �����. ���� �� ���������.", vbCritical
lstDrive.Drive = Left(lstPath.path, 2)
UpdateDisplay
End Sub

Private Sub lstFiles_Click()
txtFileName.Text = lstFiles.FileName
UpdateDisplay
End Sub

Private Sub lstFiles_DblClick()
OKButton_Click
End Sub

Private Sub lstPath_Change()
On Error Resume Next
lstFiles.path = lstPath.path
If Err Then MsgBox "������ ������ � �����. ���� �� ������.", vbCritical
UpdateDisplay
End Sub

Sub UpdateDisplay()
On Error Resume Next
Dim FN, LN As Integer, SP As Integer

' Other updates

lblStatus.SimpleText = LowPath(lstPath.path) + txtFileName.Text

End Sub

Private Sub OKButton_Click()
Dim a

If FileExists(LowPath(lstPath.path) + txtFileName.Text) = True Then
   a = MsgBox("���� ��� ����������. ������������?", vbQuestion + vbYesNo)
    If a = 7 Then Exit Sub
    If a = 6 Then UpdateDisplay: dlgSaveList.Hide
Else
    Call UpdateDisplay: dlgSaveList.Hide
End If

End Sub
