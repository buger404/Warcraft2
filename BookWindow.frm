VERSION 5.00
Begin VB.Form BookWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0049413A&
   BorderStyle     =   0  'None
   Caption         =   "BookBook"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MouseIcon       =   "BookWindow.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   949
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   11400
      Top             =   420
   End
   Begin VB.PictureBox ListFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFD973&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   510
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   873
      TabIndex        =   14
      Top             =   4500
      Width           =   13095
      Begin VB.PictureBox CardList 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFD973&
         BorderStyle     =   0  'None
         Height          =   2925
         Left            =   0
         ScaleHeight     =   195
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   457
         TabIndex        =   16
         Top             =   0
         Width           =   6855
         Begin VB.Label Cards 
            BackColor       =   &H00E9E4E9&
            Height          =   765
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   90
            Width           =   1425
         End
         Begin VB.Label CardsPre 
            BackColor       =   &H00CC7A00&
            Height          =   1095
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label CardsPre 
            BackColor       =   &H00CC7A00&
            Height          =   1095
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label CardsPre 
            BackColor       =   &H00BAB539&
            Height          =   765
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   2070
            Width           =   1425
         End
         Begin VB.Label CardsPre 
            BackColor       =   &H00BAB539&
            Height          =   765
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   1425
         End
         Begin VB.Label CardsPre 
            BackColor       =   &H00BAB539&
            Height          =   765
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   90
            Width           =   1425
         End
      End
   End
   Begin VB.Label previewframe 
      BackColor       =   &H00FFD973&
      Height          =   480
      Left            =   3990
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label scrollbutton 
      BackColor       =   &H00BAB539&
      Height          =   225
      Left            =   720
      TabIndex        =   6
      Top             =   8130
      Width           =   645
   End
   Begin VB.Label scrollbar 
      BackColor       =   &H00FFD973&
      Height          =   225
      Left            =   720
      TabIndex        =   5
      Top             =   8130
      Width           =   13500
   End
   Begin VB.Label backbutton 
      BackColor       =   &H00F0B000&
      Height          =   720
      Left            =   0
      TabIndex        =   25
      Top             =   7620
      Width           =   720
   End
   Begin VB.Label speakbutton 
      BackColor       =   &H00F0B000&
      Height          =   630
      Left            =   12900
      TabIndex        =   24
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label PreviewMonster 
      BackColor       =   &H00F0B000&
      Height          =   1095
      Left            =   3660
      TabIndex        =   20
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label buybutton 
      BackColor       =   &H00BAB539&
      Height          =   645
      Left            =   11670
      TabIndex        =   15
      Top             =   3630
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Label infotext 
      BackColor       =   &H00F0B000&
      Height          =   1065
      Left            =   7710
      TabIndex        =   13
      Top             =   3030
      Width           =   5715
   End
   Begin VB.Label cdicon 
      BackColor       =   &H00F0B000&
      Height          =   630
      Left            =   9660
      TabIndex        =   12
      Top             =   2280
      Width           =   630
   End
   Begin VB.Label duringicon 
      BackColor       =   &H00F0B000&
      Height          =   630
      Left            =   7650
      TabIndex        =   11
      Top             =   2280
      Width           =   630
   End
   Begin VB.Label speedicon 
      BackColor       =   &H00F0B000&
      Height          =   630
      Left            =   11550
      TabIndex        =   10
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label hpicon 
      BackColor       =   &H00F0B000&
      Height          =   630
      Left            =   9660
      TabIndex        =   9
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label attackicon 
      BackColor       =   &H00F0B000&
      Height          =   630
      Left            =   7650
      TabIndex        =   8
      Top             =   1650
      Width           =   630
   End
   Begin VB.Label effecticon 
      BackColor       =   &H00F0B000&
      Height          =   630
      Left            =   480
      TabIndex        =   7
      Top             =   1140
      Width           =   630
   End
   Begin VB.Label infoframe 
      BackColor       =   &H00FFD973&
      Height          =   3105
      Left            =   7470
      TabIndex        =   4
      Top             =   1170
      Width           =   6135
   End
   Begin VB.Label bigframe 
      BackColor       =   &H00F0B000&
      Height          =   3165
      Left            =   510
      TabIndex        =   2
      Top             =   4500
      Visible         =   0   'False
      Width           =   13095
   End
   Begin VB.Label bookbutton2 
      BackColor       =   &H00FFD973&
      Height          =   825
      Left            =   13200
      TabIndex        =   1
      Top             =   180
      Width           =   825
   End
   Begin VB.Label bookbutton1 
      BackColor       =   &H00F0B000&
      Height          =   825
      Left            =   12330
      TabIndex        =   0
      Top             =   180
      Width           =   825
   End
End
Attribute VB_Name = "BookWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UI As Long
Dim FBMonster() As New FriendlyMonster, EBMonster() As New EvilMonster
Dim BookMode As Integer, NowMonster As Integer
Dim StartX As Single
Public lastobj As String

Private Sub Fightbutton_Click()

End Sub

Private Sub backbutton_Click()
Me.Hide
CreateAChild MainWindow
Unload Me
End Sub

Private Sub bookbutton1_Click()
If BookMode <> 0 Then

DrawTimer.Enabled = False
NowMonster = 0
scrollbutton.Left = scrollbar.Left
For i = 1 To Cards.UBound
Unload Cards(i)
Next
Cards(0).Move CardsPre(0).Left, CardsPre(0).top, CardsPre(0).Width, CardsPre(0).Height
For i = 1 To UBound(FBMonster)
Load Cards(i)
With Cards(i)
    .Left = CardsPre(0).Left + Int(i / 3) * (Cards(0).Width + 20)
    .top = CardsPre(i Mod 3).top
    .Visible = True
    .ZOrder
End With
Next
scrollcount = Int(UBound(FBMonster) / 3) - 6 + 1
If scrollcount <= 0 Then
scrollbutton.Width = scrollbar.Width
Else
scrollbutton.Width = scrollbar.Width / scrollcount
End If
CardList.Width = (Int(UBound(FBMonster) / 3) + 1) * (Cards(0).Width + 20)
BookMode = 0
'spendicon.Visible = True
cdicon.Visible = True
duringicon.Visible = True
DrawTimer.Enabled = True

End If
End Sub

Private Sub bookbutton2_Click()
If BookMode <> 1 Then

DrawTimer.Enabled = False
NowMonster = 0
scrollbutton.Left = scrollbar.Left
For i = 1 To Cards.UBound
Unload Cards(i)
Next
Cards(0).Move CardsPre(3).Left, CardsPre(3).top, 73, 73
For i = 1 To UBound(EBMonster)
Load Cards(i)
With Cards(i)
    .Left = CardsPre(3).Left + Int(i / 2) * (Cards(0).Width + 20)
    .top = CardsPre((i Mod 2) + 3).top
    .Visible = True
    .ZOrder
End With
Next
scrollcount = Int(UBound(EBMonster) / 2) - 8 + 1
If scrollcount <= 0 Then
scrollbutton.Width = scrollbar.Width
Else
scrollbutton.Width = scrollbar.Width / scrollcount
End If
CardList.Width = (Int(UBound(EBMonster) / 2) + 1) * (Cards(0).Width + 20)
BookMode = 1
'spendicon.Visible = False
cdicon.Visible = False
duringicon.Visible = False
DrawTimer.Enabled = True

End If
End Sub

Private Sub buybutton_Click()
If buybutton.Tag = "equip" Then
needMoneys = (MyMonsterLevel(FindMonster(FBMonster(NowMonster).MonsterName)) + 1) * 3000
If Msgbox("确定要强化吗？" & vbCrLf & "你将要花费" & needMoneys & "游戏币", , "强化？") = 1 Then Exit Sub
If Money >= needMoneys Then
Money = Money - needMoneys
MyMonsterLevel(FindMonster(FBMonster(NowMonster).MonsterName)) = MyMonsterLevel(FindMonster(FBMonster(NowMonster).MonsterName)) + 1
ShowInfo "花费 " & needMoneys & " 游戏币", "money", "coin"
Call WriteSave
Else
Ring "warning"
Msgbox "你的钱不够啦！", , "GG"
End If

Else

Me.Hide
Dim tempUI As Long '临时储存用
NowWorld = UBound(NowLevel)
CreateAChild FightWindow: GdipCreateFromHDC FightWindow.HDC, tempUI
FightWindow.UI = tempUI: FightWindow.DrawTimer.Enabled = True
Dim i As Integer
For i = 0 To 11
FightWindow.SetCard i, FBMonster(NowMonster).MonsterName, 0
Next
FunnyCounts = 9999
FightWindow.SuperCounts = 9999
FightWindow.DebugMode = True
FightWindow.SetLevel 50
Unload Me

End If
End Sub

Private Sub Cards_Click(Index As Integer)
NowMonster = Index
End Sub

Private Sub DrawTimer_Timer()
    Dim p As POINTAPI
    GetCursorPos p: p.X = p.X - Dad.Left / 15: p.Y = p.Y - Dad.top / 15 - 16

GdipCreateFromHDC Me.HDC, UI
GdipCreateFromHDC CardList.HDC, UI2
GdipSetTextRenderingHint UI, TextRenderingHintAntiAlias '平滑的字字
GdipSetTextRenderingHint UI2, TextRenderingHintAntiAlias '平滑的字字

If BookMode = 0 Then
GamePictures(GetPic(FBMonster(NowMonster).BookPic)).NextFrame.Present Me.HDC, 0, 0, 100
Else
GamePictures(GetPic(EBMonster(NowMonster).BookPic)).NextFrame.Present Me.HDC, 0, 0, 100
End If

DrawPicNameControl bookbutton1, "bookbutton1"
DrawPicNameControl bookbutton2, "bookbutton2"
FillControl BigFrame, "whiteeffect"
FillControl infoframe, "whiteeffect"
If p.Y >= scrollbar.top Then
FillControl scrollbar, "whiteeffect"
FillControl scrollbutton, "blueeffect"
End If
DrawPicNameControl attackicon, "attackicon"
DrawPicNameControl hpicon, "hpicon"
DrawPicNameControl speedicon, "speedicon"
DrawPicNameControl cdicon, "cdicon"
DrawPicNameControl duringicon, "duringicon"
DrawPicNameControl backbutton, "backbutton"
DrawPicNameControl speakbutton, "speakbutton"


BitBlt ListFrame.HDC, 0, 0, ListFrame.Width, ListFrame.Height, Me.HDC, ListFrame.Left, ListFrame.top, vbSrcCopy
ListFrame.Refresh
CardList.Left = -((scrollbutton.Left - scrollbar.Left) / scrollbutton.Width * (Cards(0).Width + 30)) '* 15
BitBlt CardList.HDC, -CardList.Left, -CardList.top, ListFrame.Width, ListFrame.Height, ListFrame.HDC, 0, 0, vbSrcCopy

buybutton.Visible = False
buybutton.Tag = ""
If BookMode = 0 Then

        DrawPicNameControl effecticon, "effecticon"
        
        If (GetTickCount - FBMonster(NowMonster).LastFireTime) Mod 1000 < 100 Then
        DrawPicNameControl PreviewMonster, FBMonster(NowMonster).MonsterName & "1"
        Else
        DrawPicNameControl PreviewMonster, FBMonster(NowMonster).MonsterName & "0"
        End If
        
        DrawPicNameControl buybutton, "buybutton0"
        DrawTextRect effecticon.Left + 50, effecticon.top + 12, GetEffectName(FBMonster(NowMonster).MonsterType), argb(255, 252, 252, 252), StringAlignmentNear
        DrawTextRect attackicon.Left + 50, attackicon.top + 12, "攻击 " & GetLevelStr("Attack", FBMonster(NowMonster).Attack), argb(255, 32, 32, 32), StringAlignmentNear
        DrawTextRect hpicon.Left + 50, hpicon.top + 12, "血量 " & GetLevelStr("HP", FBMonster(NowMonster).HP), argb(255, 32, 32, 32), StringAlignmentNear
        DrawTextRect speedicon.Left + 50, speedicon.top + 12, "速度 " & GetLevelStr("Speed", FBMonster(NowMonster).Speed), argb(255, 32, 32, 32), StringAlignmentNear
        DrawTextRect duringicon.Left + 50, duringicon.top + 12, "间隔 " & GetLevelStr("During", FBMonster(NowMonster).During), argb(255, 32, 32, 32), StringAlignmentNear
        DrawTextRect cdicon.Left + 50, cdicon.top + 12, "冷却时间 " & GetLevelStr("CD", FBMonster(NowMonster).CDTime), argb(255, 32, 32, 32), StringAlignmentNear
'        DrawTextRect spendicon.Left + 36, spendicon.top + 12, "花费  " & FBMonster(NowMonster).Spend & "个滑稽果", argb(255, 32, 32, 32), StringAlignmentNear
        DrawTextControl infotext, FBMonster(NowMonster).info, argb(255, 64, 64, 64), StringAlignmentNear
        

        If OwnMonster(FBMonster(NowMonster).MonsterName) = False Then
        'FillControl PreviewFrame, "blackeffect"
        'GamePictures(GetPic("padlock")).NextFrame.Present Me.hDC, PreviewFrame.Left, PreviewFrame.top, 200
        DrawTextRect attackicon.Left + 3, attackicon.top - 21, FBMonster(NowMonster).MonsterName, argb(255, 0, 176, 240), StringAlignmentNear
        buybutton.Visible = True
        Else
        DrawTextRect attackicon.Left + 26, attackicon.top - 21, FBMonster(NowMonster).MonsterName, argb(255, 0, 176, 240), StringAlignmentNear
        GamePictures(GetPic("level" & MyMonsterLevel(FindMonster(FBMonster(NowMonster).MonsterName)) & "")).NextFrame.Present Me.HDC, infoframe.Left + 2, infoframe.top
        DrawTextRect PreviewFrame.Left + 50, PreviewFrame.top + 50, "x" & CanUpLevel(FBMonster(NowMonster).MonsterName), argb(255, 0, 176, 240), StringAlignmentNear
        If MyMonsterLevel(FindMonster(FBMonster(NowMonster).MonsterName)) < 5 And Money >= (MyMonsterLevel(FindMonster(FBMonster(NowMonster).MonsterName)) + 1) * 3000 Then
        buybutton.Visible = True
        buybutton.Tag = "equip"
        DrawPicNameControl buybutton, "equipbutton1"
        End If
        End If

        LastPenColor = argb(255, 32, 32, 32)
        GdipSetSolidFillColor Brush1, argb(255, 32, 32, 32)

       For i = 0 To Cards.UBound
            If Cards(i).Left + CardList.Left <= ListFrame.Width And Cards(i).Left + CardList.Left >= -Cards(0).Width Then
                
                GamePictures(GetPic("card1")).NextFrame.PresentWithClip CardList.HDC, Cards(i).Left, Cards(i).top, 0, 0, Cards(i).Width, Cards(i).Height
                GamePictures(GetPic(FBMonster(i).MonsterName & "Icon")).NextFrame.Present CardList.HDC, Cards(i).Left + 20, Cards(i).top + 1
                
                GamePictures(GetPic("monsterframe")).NextFrame.Present CardList.HDC, Cards(i).Left, Cards(i).top
                GdipDrawString UI2, StrPtr(FBMonster(i).Spend), -1, curFont, NewRectF(Cards(i).Left + 62, Cards(i).top + 31, 0, 0), strformat, Brush1
                If OwnMonster(FBMonster(i).MonsterName) = False Then
                GamePictures(GetPic("blackeffect")).NextFrame.PresentWithClip CardList.HDC, Cards(i).Left, Cards(i).top, 0, 0, Cards(i).Width, Cards(i).Height
                GamePictures(GetPic("padlock")).NextFrame.Present CardList.HDC, Cards(i).Left + 31, Cards(i).top + 9, 200
                End If
            End If
        Next

Else

        DrawPicNameControl effecticon, "effecticon"
        
        If GetTickCount Mod 200 < 100 Then
        DrawPicNameControl PreviewMonster, EBMonster(NowMonster).MonsterName & "1"
        Else
        DrawPicNameControl PreviewMonster, EBMonster(NowMonster).MonsterName & "0"
        End If
        
        DrawTextRect effecticon.Left + 50, effecticon.top + 12, GetEffectName(EBMonster(NowMonster).MonsterType), argb(255, 252, 252, 252), StringAlignmentNear
        DrawTextRect attackicon.Left + 50, attackicon.top + 12, "攻击 " & GetLevelStr("Attack", EBMonster(NowMonster).Attack), argb(255, 32, 32, 32), StringAlignmentNear
        DrawTextRect hpicon.Left + 50, hpicon.top + 12, "血量 " & GetLevelStr("HP", EBMonster(NowMonster).HP), argb(255, 32, 32, 32), StringAlignmentNear
        DrawTextRect speedicon.Left + 50, speedicon.top + 12, "速度 " & GetLevelStr("Speed", EBMonster(NowMonster).Speed, True), argb(255, 32, 32, 32), StringAlignmentNear
        DrawTextControl infotext, EBMonster(NowMonster).info, argb(255, 64, 64, 64), StringAlignmentNear
        
        DrawTextRect attackicon.Left + 3, attackicon.top - 23, EBMonster(NowMonster).MonsterName, argb(255, 0, 176, 240), StringAlignmentNear

        For i = 0 To Cards.UBound
        If Cards(i).Left + CardList.Left <= ListFrame.Width And Cards(i).Left + CardList.Left >= -Cards(0).Width Then
                GamePictures(GetPic("card1")).NextFrame.PresentWithClip CardList.HDC, Cards(i).Left, Cards(i).top, 0, 0, Cards(i).Width, Cards(i).Height
                With GamePictures(GetPic(EBMonster(i).MonsterName & "0"))
                .NextFrame.PresentWithClip CardList.HDC, Cards(i).Left, Cards(i).top, .NextFrame.Width / 2 - Cards(i).Width / 2, 0, Cards(i).Width, Cards(i).Height
                End With
        End If
        Next

End If

'DrawTextRect 20, 10, "你好，404！", argb(255, 255, 255, 255), StringAlignmentNear
'DrawTextRect 795, 10, "0", argb(255, 255, 255, 255), StringAlignmentNear
'DrawTextRect 875, 10, "0只", argb(255, 255, 255, 255), StringAlignmentNear
TheEnd:
If buybutton.Tag = "" Then DrawPicNameControl buybutton, "buybutton0"
'==================================================================================
' UI ++
On Error Resume Next
For Each hey In Me.Controls
Err.Clear
If hey.name <> Me.name Then
    If p.X >= hey.Left And p.X <= hey.Left + hey.Width And p.Y >= hey.top And p.Y <= hey.top + hey.Height And hey.Visible = True And hey.name <> "CardList" And hey.name <> "Cards" And hey.name <> "CardsPre" And InStr(hey.name, "icon") = 0 Then
        If Err.Number = 0 Then
        If lastobj <> hey.name Then Ring "move": lastobj = hey.name
        GamePictures(GetPic("clickframe")).NextFrame.PresentWithClip Me.HDC, hey.Left, hey.top, 0, 0, hey.Width, hey.Height, 20
        Exit For
        End If
    End If
End If
Next
If hey.name <> lastobj Then lastobj = ""
'==================================================================================

Call DrawEffect(UI, Me.name)

CardList.Refresh
FPS = FPS + 1
Me.Refresh

'FightWindow.PaintOnce Picture1.HDC, Picture1.ScaleWidth, Picture1.ScaleHeight
'Picture1.Refresh

GdipDeleteGraphics UI
GdipDeleteGraphics UI2
End Sub

Private Sub Form_Activate()
On Error Resume Next
'Dad.SetFocus
End Sub
Private Sub ScrollBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call scrollbutton_MouseDown(1, 0, 0, 0)
End Sub

Private Sub ScrollBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call scrollbutton_MouseMove(1, 0, 0, 0)
End Sub
Private Sub Form_Load()
On Error Resume Next '错误错误快离开~~~
Dad.SetFocus
NowMonster = 0
BookMode = 0
Dim temp() As String, file As String

PrevWndProc = SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf WndProcForBookWindow)

For Each ClickArea In Me.Controls
ClickArea.BackStyle = 0 '变成偷懒用的点击检测吧~~~~~
Next

ReDim FBMonster(0), EBMonster(0)

file = Dir(App.Path & "\monster\fmonster\")
temp = Split(file, ".")
FBMonster(0).LoadMonster temp(0), 0
Do While file <> ""
file = Dir()
If file <> "" Then
temp = Split(file, ".")
ReDim Preserve FBMonster(UBound(FBMonster) + 1)
FBMonster(UBound(FBMonster)).LoadMonster temp(0), 0
End If
DoEvents
Loop

file = Dir(App.Path & "\monster\emonster\")
temp = Split(file, ".")
EBMonster(0).LoadMonster temp(0), 0
Do While file <> ""
file = Dir()
If file <> "" Then
temp = Split(file, ".")
ReDim Preserve EBMonster(UBound(EBMonster) + 1)
EBMonster(UBound(EBMonster)).LoadMonster temp(0), 0
End If
DoEvents
Loop

For i = 1 To UBound(FBMonster)
Load Cards(i)
With Cards(i)
    .Left = CardsPre(i Mod 3).Left + Int(i / 3) * (Cards(0).Width + 20)
    .top = CardsPre(i Mod 3).top
    .Visible = True
    .ZOrder
End With
Next

CardList.Width = (Int(UBound(FBMonster) / 3) + 1) * (Cards(0).Width + 20)

scrollbutton.Left = scrollbar.Left
scrollcount = Int(UBound(FBMonster) / 3) - 6 + 1
If scrollcount <= 0 Then
scrollbutton.Width = scrollbar.Width
Else
scrollbutton.Width = scrollbar.Width / scrollcount
End If

DrawTimer.Enabled = True

End Sub
Sub FillControl(Control As Object, PicName As String)
If Control.Visible = False Then Exit Sub
GamePictures(GetPic(PicName)).NextFrame.PresentWithClip Me.HDC, Control.Left, Control.top, 0, 0, Control.Width, Control.Height
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
SetWindowLongA Me.hWnd, GWL_WNDPROC, PrevWndProc
GdipDeleteGraphics UI
End Sub

Private Sub scrollbutton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartX = X / 15
End Sub

Public Sub scrollbutton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Dim p As POINTAPI
GetCursorPos p
p.X = p.X - StartX
tmpx = p.X - Dad.Left / 15
If tmpx < scrollbar.Left Then tmpx = scrollbar.Left
If tmpx > scrollbar.Left + scrollbar.Width - scrollbutton.Width Then tmpx = scrollbar.Left + scrollbar.Width - scrollbutton.Width
'tmpx = scrollbar.Left + Round((tmpx - scrollbar.Left) / scrollbutton.Width) * scrollbutton.Width
scrollbutton.Left = tmpx
End If
End Sub

Private Sub speakbutton_Click()
On Error Resume Next
If BookMode = 0 Then
Ring FBMonster(NowMonster).MonsterName
Else
Ring EBMonster(NowMonster).MonsterName
End If
End Sub

