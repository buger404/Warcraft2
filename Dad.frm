VERSION 5.00
Begin VB.Form Dad 
   AutoRedraw      =   -1  'True
   BackColor       =   &H002D282A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "魔兽混战2"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14145
   Icon            =   "Dad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Dad.frx":1BCC2
   MousePointer    =   99  'Custom
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   943
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   13590
      Top             =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "小贴士：并不是所有魔兽都需要到GG商店购买，你也可以通关某个关卡获得这些魔兽。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0069665E&
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   8040
      Width           =   14145
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "加载中"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0B000&
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   4170
      Width           =   14145
   End
End
Attribute VB_Name = "Dad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IsTimeToClose As Boolean, CloseTime As Long
Private Sub Form_Load()
InitGDIPlus
BASS_Init -1, 44100, BASS_DEVICE_3D, Me.hWnd, 0
BASS_SetVolume 0.03
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
IsTimeToClose = True
CloseTime = GetTickCount
Ring "bye"
FadeOut 50, Dad
Unload FightWindow '乖徒儿，快快退下便是。
Unload MainWindow
Unload BookWindow
Unload BGMBox
Unload ChooseWindow
Unload ChooseWindow2
Unload GGShop
Unload CardWindow
Unload SetWindow
Unload MsgWindow
Unload AeroWindow
Unload testWindow
Unload EquipWindow
For i = 0 To UBound(GamePictures)
GamePictures(i).NextFrame.Delete
Next
    GdipDeleteBrush Brush1
    GdipDeletePen Pen1
    GdipDeleteFontFamily fontfam
    GdipDeleteStringFormat strformat
    GdipDeleteFont curFont
    GdipDeleteBrush Brush
TerminateGDIPlus
BASS_Free

End
End Sub

Private Sub Timer1_Timer()
If IsTimeToClose = True And GetTickCount - CloseTime >= 6000 Then
Msgbox "天啊，有一些过程拖着我的后退，不愿意我被关闭，按下确定后强制关闭。", 48, "魔兽混战2"
End
End If
End Sub
