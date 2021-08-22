VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form BGMBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BGMBox"
   ClientHeight    =   690
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4245
   Icon            =   "BGMBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   4245
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer ForDad 
      Interval        =   30
      Left            =   1830
      Top             =   90
   End
   Begin VB.Timer onesreturn 
      Interval        =   1000
      Left            =   2430
      Top             =   90
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3720
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3090
      Top             =   90
   End
End
Attribute VB_Name = "BGMBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DoneNow As Boolean

Public Sub DrawScreen(TForm As Form)
On Error Resume Next
Dim p As POINTAPI
GetCursorPos p
p.X = p.X - Dad.Left / 15: p.Y = p.Y - Dad.top / 15 - 16

Dim Fuck() As Single, UI As Long, temp2 As Long

GdipCreateFromHDC TForm.HDC, UI

If Sets(3) = True Or TForm.name = "WelcomeWindow" Or TForm.name = "MainWindow" Then
Fuck = BGM.GetMusicBar
For i = 0 To 46
GamePictures(GetPic("blueeffect")).NextFrame.PresentWithClip TForm.HDC, i * 20, TForm.ScaleHeight - Fuck(i) / 7, 0, 0, 15, Fuck(i) / 7
temp2 = temp2 + Fuck(i)
Next
GamePictures(GetPic("frame")).NextFrame.Present TForm.HDC, Dad.ScaleWidth - 130, 227, 150
DrawTextRectUI UI, Dad.ScaleWidth - 22, 235, Int(temp2 / 1000) & " Hot Level", argb(150, 64, 64, 64), StringAlignmentFar, False
End If

If Sets(3) = True Then
GamePictures(GetPic("frame")).NextFrame.Present TForm.HDC, Dad.ScaleWidth - 130, 107, 150
DrawTextRectUI UI, Dad.ScaleWidth - 22, 115, TForm.DrawTimer.Interval & " Interval", argb(150, 64, 64, 64), StringAlignmentFar, False
GamePictures(GetPic("frame")).NextFrame.Present TForm.HDC, Dad.ScaleWidth - 130, 147, 150
DrawTextRectUI UI, Dad.ScaleWidth - 22, 155, UBound(GamePictures) & " Pictures", argb(150, 64, 64, 64), StringAlignmentFar, False
GamePictures(GetPic("frame")).NextFrame.Present TForm.HDC, Dad.ScaleWidth - 130, 187, 150
DrawTextRectUI UI, Dad.ScaleWidth - 22, 195, UBound(GameSounds) & " Sounds", argb(150, 64, 64, 64), StringAlignmentFar, False
End If

If GetTickCount - VoChangeTime <= 5000 Then
GamePictures(GetPic("frame")).NextFrame.Present TForm.HDC, Dad.ScaleWidth - 130, 107, 150
DrawTextRectUI UI, Dad.ScaleWidth - 22, 115, "音量 " & Int(GameVo / 0.15 * 100) & "%", argb(150, 64, 64, 64), StringAlignmentFar, False
End If

GamePictures(GetPic("mouse")).NextFrame.Present TForm.HDC, p.X - 16, p.Y - 16

If IsPreviewVersion = True Then GamePictures(GetPic("preview")).NextFrame.Present TForm.HDC, TForm.ScaleWidth / 2 - 230, TForm.ScaleHeight - 42

If Sets(8) = True Then
GamePictures(GetPic("frame")).NextFrame.Present TForm.HDC, Dad.ScaleWidth - 130, 67, 150
Select Case RFPS
    Case Is < 10
    DrawTextRectUI UI, Dad.ScaleWidth - 22, 75, RFPS & " FPS", argb(150, 255, 0, 0), StringAlignmentFar, False
    Case Is < 20
    DrawTextRectUI UI, Dad.ScaleWidth - 22, 75, RFPS & " FPS", argb(150, 255, 192, 0), StringAlignmentFar, False
    Case Is < 30
    DrawTextRectUI UI, Dad.ScaleWidth - 22, 75, RFPS & " FPS", argb(150, 64, 64, 64), StringAlignmentFar, False
    Case Else
    DrawTextRectUI UI, Dad.ScaleWidth - 22, 75, RFPS & " FPS", argb(150, 0, 176, 80), StringAlignmentFar, False
End Select
End If
End Sub

Private Sub ForDad_Timer()
If ActiveWindow.Visible = False Then
Dad.Cls
DrawScreen Dad
Dad.Refresh
End If

End Sub
Sub ChangeVo()
VoChangeTime = GetTickCount
If GameVo > 0.15 Then GameVo = 100
If GameVo < 0 Then GameVo = 0
BASS_SetVolume GameVo
End Sub
Private Sub Form_Load()
DoneNow = True
End Sub

Private Sub onesreturn_Timer()
RFPS = FPS
FPS = 0
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If GetAsyncKeyState(vbKeySubtract) Then
GameVo = GameVo - 0.01
ChangeVo
End If
If GetAsyncKeyState(vbKeyAdd) Then
GameVo = GameVo + 0.01
ChangeVo
End If

If BGM.PlayState = Stopped Then BGM.Play

If MainBGM = False Then

If DoneNow = True Then
Unload FightWindow
With BGM
    .SetPlayRate 1
    .StopMusic
    .LoadMusic App.Path & "\music\" & "Background" & UBound(NowLevel) + 1 & ".mp3"
    .Play
End With
DoneNow = False
End If

Else

DoneNow = True

End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim temp As String, data() As String, Run() As String

Winsock1.GetData temp
data = Split(temp, "灥")
For i = 0 To UBound(data)
    If data(i) <> "" Then
    Run = Split(data(i), " ")

    Select Case Run(0)
        Case "s"
        OnLineMode = 2
        NowWorld = Val(Run(1))
        CloseInfo = True
        Case "s2"
        OnLineMode = 1
        NowWorld = Val(Run(1))
        CloseInfo = True
        Case "t"
        Ring "warning"
        FightWindow.NewMsg Run(1)
        Case "fm"
        Ring "put"
        ReDim Preserve FMonster(UBound(FMonster) + 1)
        FMonster(UBound(FMonster)).X = Val(Run(2))
        FMonster(UBound(FMonster)).Y = Val(Run(3))
        FMonster(UBound(FMonster)).LoadMonster Run(1), UBound(FMonster)
        Case "em"
        ReDim Preserve EMonster(UBound(EMonster) + 1)
        EMonster(UBound(EMonster)).X = Val(Run(2))
        EMonster(UBound(EMonster)).Y = Val(Run(3))
        EMonster(UBound(EMonster)).MonsterName = Run(1)
        EMonster(UBound(EMonster)).LoadMonster EMonster(UBound(EMonster)).MonsterName, UBound(EMonster)
        Case "su"
        FightWindow.SuperMon Val(Run(1))
        Case "dm"
        MakeFMonsterDead Val(Run(1))
        Case "ef"
            Select Case Run(1)
            Case "ice"
            EMonster(Run(2)).IceTime = GetTickCount
            Case "fire"
            EMonster(Run(2)).FireTime = GetTickCount
            Case "che"
            EMonster(Run(2)).ChemicalMode = True
            End Select
    End Select
    
    End If '大判断，前面是If Data(i) <> "" Then
    
Next
End Sub

