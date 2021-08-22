VERSION 5.00
Begin VB.Form MainWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0049413A&
   BorderStyle     =   0  'None
   Caption         =   "主窗口"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MouseIcon       =   "MainWindow.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   949
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer DrawTimer 
      Interval        =   20
      Left            =   13710
      Top             =   870
   End
   Begin VB.Label BigIcon 
      BackColor       =   &H00F0B000&
      Height          =   5250
      Left            =   4470
      TabIndex        =   7
      Top             =   1650
      Width           =   5250
   End
   Begin VB.Label EquipButton 
      BackColor       =   &H00F0B000&
      Height          =   825
      Left            =   11910
      TabIndex        =   6
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Setting 
      BackColor       =   &H00F0B000&
      Height          =   825
      Left            =   11910
      TabIndex        =   5
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label OnlineButton 
      BackColor       =   &H00F0B000&
      Height          =   825
      Left            =   11910
      TabIndex        =   4
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Titlebar 
      BackColor       =   &H00F0B000&
      Height          =   645
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3345
   End
   Begin VB.Label fightbutton 
      BackColor       =   &H00F0B000&
      Height          =   825
      Left            =   11910
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label bookbutton 
      BackColor       =   &H00F0B000&
      Height          =   825
      Left            =   11910
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Shopbutton 
      BackColor       =   &H00F0B000&
      Height          =   825
      Left            =   11910
      TabIndex        =   0
      Top             =   3360
      Width           =   2295
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UI As Long
Dim lastobj As String
Dim NowBack As String
Dim CenterBig As String
Dim Fuck() As Single

Private Sub BigIcon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = 0 Then
Msgbox "游戏制作：Error 404 (QQ1361778219)" & vbCrLf & _
            "素材&音效：RPG MV+网络（部分来自超级幻影猫2）" & vbCrLf & _
            "音乐：滚动的天空（猎豹移动公司）+网络" & vbCrLf & _
            "技术支持：方程", , "关于"
Else
BGM.StopMusic
End If
End Sub

Private Sub bookbutton_Click()
Me.Hide
CreateAChild BookWindow
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub bookbutton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CenterBig = "bookbutton"
End Sub

Private Sub DrawTimer_Timer()
On Error Resume Next
GdipCreateFromHDC Me.HDC, UI
GdipSetTextRenderingHint UI, TextRenderingHintAntiAlias '平滑的字字

Dim p As POINTAPI
GetCursorPos p
p.X = p.X - Dad.Left / 15: p.Y = p.Y - Dad.top / 15 - 16
'If Sets(2) = False Then
'GamePictures(GetPic("Back" & UBound(NowLevel) + 1 & "blur")).NextFrame.Present Me.HDC, 0, 0
'Else
GamePictures(GetPic("MainBack")).NextFrame.Present Me.HDC, 0, 0, 100
'End I

'For i = 0 To UBound(MyMonster)
'If GetTickCount Mod 1000 < 100 Then '眨眼动画_(:з」∠)_
'GamePictures(GetPic(MyMonster(i) & "1")).NextFrame.Present Me.HDC, 80 + (i Mod 15) * 50, 120 + Int(i / 15) * 36
'Else
'GamePictures(GetPic(MyMonster(i) & "0")).NextFrame.Present Me.HDC, 80 + (i Mod 15) * 50, 120 + Int(i / 15) * 36
'End If
'Next
'Dim Fuck2() As Single
'Fuck2 = BGM.GetMusicBar
'For i = 0 To 5
'If Abs(Fuck2(i) - Fuck(i)) > 200 Then Fuck(i) = Fuck2(i)
'Next
Fuck = BGM.GetMusicBar
DrawPicNameControl BigIcon, CenterBig
'DrawPicNameControl Titlebar, "titlebar"

Dim NowObject As Object

Set NowObject = Setting
GamePictures(GetPic("whiteeffect")).NextFrame.PresentWithClip Me.HDC, NowObject.Left + 55, NowObject.top + 10, 0, 0, NowObject.Width - 55, NowObject.Height - 20, 30
NowObject.Left = Me.ScaleWidth - Fuck(4) / 10 - 100
NowObject.Width = Me.ScaleWidth - NowObject.Left + 100
DrawTextRect NowObject.Left + 65, NowObject.top + 18, "设置", argb(255, 255, 255, 255), StringAlignmentNear
DrawPicNameControl Setting, "settings"

Set NowObject = Shopbutton
GamePictures(GetPic("whiteeffect")).NextFrame.PresentWithClip Me.HDC, NowObject.Left + 55, NowObject.top + 10, 0, 0, NowObject.Width - 55, NowObject.Height - 20, 30
NowObject.Left = Me.ScaleWidth - Fuck(2) / 10 - 100
NowObject.Width = Me.ScaleWidth - NowObject.Left + 100
DrawTextRect NowObject.Left + 65, NowObject.top + 18, "商店", argb(255, 255, 255, 255), StringAlignmentNear
DrawPicNameControl Shopbutton, "shopbutton"

Set NowObject = bookbutton
GamePictures(GetPic("whiteeffect")).NextFrame.PresentWithClip Me.HDC, NowObject.Left + 55, NowObject.top + 10, 0, 0, NowObject.Width - 55, NowObject.Height - 20, 30
NowObject.Left = Me.ScaleWidth - Fuck(1) / 10 - 100
NowObject.Width = Me.ScaleWidth - NowObject.Left + 100
DrawTextRect NowObject.Left + 65, NowObject.top + 18, "图鉴", argb(255, 255, 255, 255), StringAlignmentNear
DrawPicNameControl bookbutton, "menubutton"

Set NowObject = Fightbutton
GamePictures(GetPic("whiteeffect")).NextFrame.PresentWithClip Me.HDC, NowObject.Left + 55, NowObject.top + 10, 0, 0, NowObject.Width - 55, NowObject.Height - 20, 30
NowObject.Left = Me.ScaleWidth - Fuck(0) / 10 - 100
NowObject.Width = Me.ScaleWidth - NowObject.Left + 100
DrawTextRect NowObject.Left + 65, NowObject.top + 18, "战斗", argb(255, 255, 255, 255), StringAlignmentNear
DrawPicNameControl Fightbutton, "fightbutton2"

Set NowObject = EquipButton
GamePictures(GetPic("whiteeffect")).NextFrame.PresentWithClip Me.HDC, NowObject.Left + 55, NowObject.top + 10, 0, 0, NowObject.Width - 55, NowObject.Height - 20, 30
NowObject.Left = Me.ScaleWidth - Fuck(5) / 10 - 100
NowObject.Width = Me.ScaleWidth - NowObject.Left + 100
DrawTextRect NowObject.Left + 65, NowObject.top + 18, "强化", argb(255, 255, 255, 255), StringAlignmentNear
DrawPicNameControl EquipButton, "superfight"

Set NowObject = OnlineButton
GamePictures(GetPic("whiteeffect")).NextFrame.PresentWithClip Me.HDC, NowObject.Left + 55, NowObject.top + 10, 0, 0, NowObject.Width - 55, NowObject.Height - 20, 30
NowObject.Left = Me.ScaleWidth - Fuck(3) / 10 - 100
NowObject.Width = Me.ScaleWidth - NowObject.Left + 100
DrawTextRect NowObject.Left + 65, NowObject.top + 18, "多人", argb(255, 255, 255, 255), StringAlignmentNear
DrawPicNameControl OnlineButton, "quickly"

'DrawTextRect 20, 10, "你好，" & PlayerName & "！（点击改名）", argb(255, 255, 255, 255), StringAlignmentNear
'DrawTextRect 50, Me.ScaleHeight - 23, "游戏制作：Error 404 (QQ1361778219)，素材&音效：RPG MV&网络，音乐：滚动的天空（猎豹移动公司）&网络，图形加速：方程", argb(255, 255, 255, 255), StringAlignmentNear

GamePictures(GetPic("money")).NextFrame.Present Me.HDC, 20, 20
DrawTextRect 70, 32, Money, argb(255, 255, 255, 255), StringAlignmentNear
GamePictures(GetPic("effecticon")).NextFrame.Present Me.HDC, 20, 70
DrawTextRect 70, 82, UBound(MyMonster) + 1 & "只", argb(255, 255, 255, 255), StringAlignmentNear
Call DrawEffect(UI, Me.name)
'==================================================================================
' UI ++
On Error Resume Next
For Each hey In Me.Controls
Err.Clear
If hey.name <> Me.name Then
    If p.X >= hey.Left And p.X <= hey.Left + hey.Width And p.Y >= hey.top And p.Y <= hey.top + hey.Height And hey.Visible = True Then
        If Err.Number = 0 Then
        If lastobj <> hey.name Then Ring "move": lastobj = hey.name
        'DrawShadowRectangle UI, argb(150, 38, 38, 38), argb(255, 38, 38, 38), hey.Left, hey.top, hey.Width, hey.Height
        GamePictures(GetPic("clickframe")).NextFrame.PresentWithClip Me.HDC, hey.Left, hey.top, 0, 0, hey.Width, hey.Height, 20
        Exit For
        End If
    End If
End If
Next
If hey.name <> lastobj Then lastobj = ""
'==================================================================================



FPS = FPS + 1
Me.Refresh
GdipDeleteGraphics UI
End Sub

Private Sub EquipButton_Click()
Me.Hide
CreateAChild BookWindow
Unload Me
End Sub

Private Sub EquipButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CenterBig = "equipbutton"
End Sub

Private Sub Fightbutton_Click()
Me.Hide
CreateAChild ChooseWindow
Unload Me
End Sub

Private Sub fightbutton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CenterBig = "fightbutton"
End Sub

Private Sub Form_Activate()
On Error Resume Next
'Dad.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next '错误错误快离开~~~
Dad.SetFocus
CenterBig = "mainicon"
For Each ClickArea In Me.Controls
ClickArea.BackStyle = 0 '变成偷懒用的点击检测吧~~~~~
Next
End Sub
Sub DrawPicNameControl(Control As Object, PicName As String)
If Control.Visible = False Then Exit Sub
GamePictures(GetPic(PicName)).NextFrame.Present Me.HDC, Control.Left, Control.top
End Sub
Sub DrawRectangleRect(X As Single, Y As Single, W As Single, H As Single, BackColor As Long)
'Exit Sub
If LastPenColor <> BackColor Then
LastPenColor = BackColor
GdipSetSolidFillColor Brush1, BackColor
End If
GdipFillRectangle UI, Brush1, X, Y, W, H
End Sub
Sub DrawRectangleControl(Control As Object, BackColor As Long)
'Exit Sub
If Control.Visible = False Then Exit Sub
If LastPenColor <> BackColor Then
LastPenColor = BackColor
GdipSetSolidFillColor Brush1, BackColor
End If
GdipFillRectangle UI, Brush1, Control.Left, Control.top, Control.Width, Control.Height
End Sub
Sub DrawTextControl(Control As Object, ByVal Text As String, ForeColor As Long, mode As StringAlignment, Optional BigSize As Boolean = False)
If Control.Visible = False Then Exit Sub
If LastPenColor <> ForeColor Then
LastPenColor = ForeColor
GdipSetSolidFillColor Brush1, ForeColor
End If
GdipSetStringFormatAlign strformat, mode
If BigSize = False Then
GdipDrawString UI, StrPtr(Text), -1, curFont, NewRectF(Control.Left, Control.top, Control.Width, Control.Height), strformat, Brush1
Else
GdipDrawString UI, StrPtr(Text), -1, curFontBig, NewRectF(Control.Left, Control.top, Control.Width, Control.Height), strformat, Brush1
End If
End Sub
Sub DrawTextRect(ByVal X As Single, ByVal Y As Single, ByVal Text As String, ForeColor As Long, mode As StringAlignment, Optional BigSize As Boolean = False)
If LastPenColor <> ForeColor Then
LastPenColor = ForeColor
GdipSetSolidFillColor Brush1, ForeColor
End If
GdipSetStringFormatAlign strformat, mode
If BigSize = False Then
GdipDrawString UI, StrPtr(Text), -1, curFont, NewRectF(X, Y, 0, 0), strformat, Brush1
Else
GdipDrawString UI, StrPtr(Text), -1, curFontBig, NewRectF(X, Y, 0, 0), strformat, Brush1
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < Me.ScaleWidth / 5 * 4 Then CenterBig = "mainicon"
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'NewEffect Me.HDC, Me.name, X, Y, FadeInPic, "lost", 30
'NewEffect Me.HDC, Me.name, X - 40, Y - 25, FadeInPic, "lighting", 5
BGM.SetPlayRate Int(Rnd * 3)
NewEffect Me.HDC, Me.name, X - 40, Y - 30, TaketurnsPic, "Hit", 1
NewEffect Me.HDC, Me.name, X - 40, Y - 20, MagicText, "Boom", 5
If Button = 4 And Sets(3) = True Then
If VBA.Environ("Error404_Key") = "1FCB6793" Then testWindow.Show
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
GdipDeleteGraphics UI
End Sub

Private Sub GetMusic_Timer()

End Sub

Private Sub OnlineButton_Click()
Me.Hide
CreateAChild ChooseWindow2
Unload Me
End Sub

Private Sub OnlineButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CenterBig = "onlinebutton"
End Sub

Private Sub Setting_Click()
ShowToolWindow SetWindow
End Sub

Private Sub Setting_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CenterBig = "settingsbig"
End Sub

Private Sub Shopbutton_Click()
Me.Hide
CreateAChild GGShop
Unload Me
End Sub

Private Sub Shopbutton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CenterBig = "shopbuttonbig"
End Sub

Private Sub Titlebar_Click()
oldName = PlayerName
PlayerName = Inputbox("给你自己起一个帅气的新名字吧！")
If PlayerName = "" Then Msgbox "不起新名字就算了吧。", , "：（":  PlayerName = oldName: Exit Sub
Call WriteSave
End Sub
