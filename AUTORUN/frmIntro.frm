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
QINFO.AddItem "Кілька слів про цю штуку ;-)"
QINFO.AddItem ""
QINFO.AddItem "Плейєр складається з трьох головних частин:"
QINFO.AddItem " 1. Сам Плейєр"
QINFO.AddItem " 2. Плейлист"
QINFO.AddItem " 3. Таймер"
QINFO.AddItem ""
QINFO.AddItem "Плейєр має такі можливості:"
QINFO.AddItem " - Програвати файли ))"
QINFO.AddItem " - Стерти файл прямо із диску"
QINFO.AddItem " - Копіювати файли"
QINFO.AddItem " - Перетворити зміст плейлисту в HTML. сторінку"
QINFO.AddItem " - Має багатомовний інтерфейс ( три ;) )"
QINFO.AddItem " - Має можливість перетягувати файли із провідника"
QINFO.AddItem "    прямо в плейлист"
QINFO.AddItem " Має функцію пошуку файлів на диску"
QINFO.AddItem " і ще багато чого іншого..."
QINFO.AddItem ""
QINFO.AddItem "Гарячі клавіші:"
QINFO.AddItem " <F3 > Завантажити"
QINFO.AddItem " <F2 > Зберегти"
QINFO.AddItem " <F8 > Стерти файл"
QINFO.AddItem " <F5 > Копіювати файли"
QINFO.AddItem " <DEL> Викреслити файл"
QINFO.AddItem " <INS> Додати файли"
QINFO.AddItem " <Z>   Грати попередній"
QINFO.AddItem " <X>   Грати"
QINFO.AddItem " <C>   Грати наступний"
QINFO.AddItem " <V>   Пауза"
QINFO.AddItem " <B>   Не грати"
QINFO.AddItem " <MNU> Софт Меню"

End Sub


