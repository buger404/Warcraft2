VERSION 5.00
Begin VB.Form Dad 
   AutoRedraw      =   -1  'True
   BackColor       =   &H002D282A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ħ�޻�ս2"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   13590
      Top             =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "С��ʿ������������ħ�޶���Ҫ��GG�̵깺����Ҳ����ͨ��ĳ���ؿ������Щħ�ޡ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
Unload FightWindow '��ͽ����������±��ǡ�
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
Msgbox "�찡����һЩ���������ҵĺ��ˣ���Ը���ұ��رգ�����ȷ����ǿ�ƹرա�", 48, "ħ�޻�ս2"
End
End If
End Sub
