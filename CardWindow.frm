VERSION 5.00
Begin VB.Form CardWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H002D282A&
   BorderStyle     =   0  'None
   Caption         =   "CardWindow"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14235
   Icon            =   "CardWindow.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "CardWindow.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   949
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PreviewFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFD973&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   1860
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   213
      TabIndex        =   13
      Top             =   1440
      Width           =   3195
   End
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   12300
      Top             =   300
   End
   Begin VB.PictureBox ListFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFD973&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   5880
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   487
      TabIndex        =   0
      Top             =   1530
      Width           =   7305
      Begin VB.PictureBox CardList 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFD973&
         BorderStyle     =   0  'None
         Height          =   5505
         Left            =   0
         ScaleHeight     =   367
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   487
         TabIndex        =   1
         Top             =   0
         Width           =   7305
         Begin VB.Label CardsPre 
            BackColor       =   &H00BAB539&
            Height          =   765
            Index           =   5
            Left            =   150
            TabIndex        =   12
            Top             =   4650
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label CardsPre 
            BackColor       =   &H00BAB539&
            Height          =   765
            Index           =   4
            Left            =   150
            TabIndex        =   11
            Top             =   3750
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label CardsPre 
            BackColor       =   &H00BAB539&
            Height          =   765
            Index           =   3
            Left            =   150
            TabIndex        =   6
            Top             =   2850
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label CardsPre 
            BackColor       =   &H00BAB539&
            Height          =   765
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   150
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label CardsPre 
            BackColor       =   &H00BAB539&
            Height          =   765
            Index           =   1
            Left            =   150
            TabIndex        =   4
            Top             =   1050
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label CardsPre 
            BackColor       =   &H00BAB539&
            Height          =   765
            Index           =   2
            Left            =   150
            TabIndex        =   3
            Top             =   1950
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label Cards 
            BackColor       =   &H00E9E4E9&
            Height          =   765
            Index           =   0
            Left            =   150
            TabIndex        =   2
            Top             =   150
            Width           =   1425
         End
      End
   End
   Begin VB.Label UseFrame 
      Height          =   555
      Left            =   5880
      TabIndex        =   17
      Top             =   1020
      Width           =   7305
   End
   Begin VB.Label backbutton 
      BackColor       =   &H00F0B000&
      Height          =   720
      Left            =   0
      TabIndex        =   16
      Top             =   7650
      Width           =   720
   End
   Begin VB.Label SuperFightButton 
      BackColor       =   &H00F0B000&
      Height          =   795
      Left            =   12090
      TabIndex        =   15
      Tag             =   "超级模式"
      Top             =   7230
      Width           =   795
   End
   Begin VB.Label Fightbutton 
      BackColor       =   &H00F0B000&
      Height          =   795
      Left            =   12990
      TabIndex        =   14
      Tag             =   "开始战斗"
      Top             =   7230
      Width           =   795
   End
   Begin VB.Label scrollbutton 
      BackColor       =   &H00BAB539&
      Height          =   225
      Left            =   6180
      TabIndex        =   7
      Top             =   7050
      Width           =   645
   End
   Begin VB.Label scrollbar 
      BackColor       =   &H00F0B000&
      Height          =   225
      Left            =   5880
      TabIndex        =   8
      Top             =   7050
      Width           =   7305
   End
   Begin VB.Label PhoneFrame 
      BackColor       =   &H00F0B000&
      Height          =   6315
      Left            =   1530
      TabIndex        =   9
      Top             =   1020
      Width           =   3885
   End
   Begin VB.Label CardFrame 
      BackColor       =   &H00F0B000&
      Height          =   6255
      Left            =   5880
      TabIndex        =   10
      Top             =   1020
      Width           =   7305
   End
End
Attribute VB_Name = "CardWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UI As Long, UI2 As Long
Dim FBMonster() As New FriendlyMonster
Dim IsChoose() As Boolean
Dim ChooseCount As Integer
Dim DownTime As Long
Dim StartX As Single
Public lastobj As String
Private Sub backbutton_Click()
OnLineMode = 0
On Error Resume Next
OnLine.close
Me.Hide
CreateAChild MainWindow
Unload Me
End Sub

Private Sub Cards_Click(Index As Integer)
If OnLineMode <> 0 Then
If FBMonster(Index).MonsterType = CopyHP Or FBMonster(Index).MonsterType = RanAttack Or FBMonster(Index).MonsterType = BeatOut Then
Msgbox "多人游戏不能使用" & FBMonster(Index).MonsterName & "哦。", , ":("
Exit Sub
End If
End If

If IsChoose(Index) = True Then
IsChoose(Index) = False
ChooseCount = ChooseCount - 1
ElseIf ChooseCount < 12 Then
IsChoose(Index) = True
ChooseCount = ChooseCount + 1
Else
Msgbox "不能再选了哦。", , ":("
End If
End Sub

Private Sub DrawTimer_Timer()
On Error Resume Next
Dim p As POINTAPI
GetCursorPos p
p.X = p.X - Dad.Left / 15: p.Y = p.Y - Dad.top / 15 - 16

GdipCreateFromHDC Me.HDC, UI
GdipCreateFromHDC CardList.HDC, UI2
GdipSetTextRenderingHint UI, TextRenderingHintAntiAlias '平滑的字字
GdipSetTextRenderingHint UI2, TextRenderingHintAntiAlias '平滑的字字

If Sets(2) = False Then
GamePictures(GetPic("Back" & NowWorld + 1 & "blur")).NextFrame.Present Me.HDC, 0, 0
Else
GamePictures(GetPic("Back" & NowWorld + 1 & "blur")).NextFrame.Present Me.HDC, 0, 0, 100
End If
DrawPicNameControl PhoneFrame, "phoneframe"
FillControl CardFrame, "whiteeffect"
FillControl scrollbar, "whiteeffect"
FillControl scrollbutton, "blueeffect"
FillControl Fightbutton, "fightbutton2"
FillControl SuperFightButton, "superfight"
FillControl UseFrame, "blueeffect"
DrawPicNameControl backbutton, "backbutton"

BitBlt ListFrame.HDC, 0, 0, ListFrame.Width, ListFrame.Height, Me.HDC, ListFrame.Left, ListFrame.top, vbSrcCopy
ListFrame.Refresh
CardList.Left = -((scrollbutton.Left - scrollbar.Left) / scrollbutton.Width * (Cards(0).Width + 30)) '* 15
BitBlt CardList.HDC, -CardList.Left, -CardList.top, ListFrame.Width, ListFrame.Height, ListFrame.HDC, 0, 0, vbSrcCopy

        LastPenColor = argb(255, 32, 32, 32)
        GdipSetSolidFillColor Brush1, argb(255, 32, 32, 32)
        GdipSetStringFormatAlign strformat, StringAlignmentCenter
        For i = 0 To Cards.UBound
                If Cards(i).Left + CardList.Left >= -Cards(i).Width And Cards(i).Left + CardList.Left <= ListFrame.ScaleWidth Then
                GamePictures(GetPic("card1")).NextFrame.PresentWithClip CardList.HDC, Cards(i).Left, Cards(i).top, 0, 0, Cards(i).Width, Cards(i).Height
                GamePictures(GetPic(FBMonster(i).MonsterName & "Icon")).NextFrame.Present CardList.HDC, Cards(i).Left + 20, Cards(i).top + 1
                
                GamePictures(GetPic("monsterframe")).NextFrame.Present CardList.HDC, Cards(i).Left - Cards(i).Width + 60, Cards(i).top, IIf(Fuck2 = True, 80, 180)
                GamePictures(GetPic("level" & MyMonsterLevel(i))).NextFrame.Present CardList.HDC, Cards(i).Left + Cards(i).Width - 35, Cards(i).top - 5, 180
                
                GdipDrawString UI2, StrPtr(FBMonster(i).Spend), -1, curFont, NewRectF(Cards(i).Left + 32.5, Cards(i).top + Cards(i).Height - 20.5, 0, 0), strformat, Brush1
                If IsChoose(i) = True Then GamePictures(GetPic("choose")).NextFrame.Present CardList.HDC, Cards(i).Left, Cards(i).top
                End If
        Next
        


GamePictures(GetPic("滑稽之花Icon")).NextFrame.PresentWithClip Me.HDC, UseFrame.Left, UseFrame.top, 0, 0, 999, UseFrame.Height
DrawTextRect UseFrame.Left + 55, UseFrame.top + 10, "已选魔兽 " & ChooseCount & "/12", argb(255, 0, 0, 0), StringAlignmentNear


'==================================================================================
' UI ++
On Error Resume Next
For Each hey In Me.Controls
Err.Clear
If hey.name <> Me.name Then
    If p.X >= hey.Left And p.X <= hey.Left + hey.Width And p.Y >= hey.top And p.Y <= hey.top + hey.Height And hey.Visible = True And hey.name <> "CardList" And hey.name <> "Cards" And hey.name <> "CardsPre" Then
        If Err.Number = 0 Then
        If lastobj <> hey.name Then Ring "move": lastobj = hey.name
        GamePictures(GetPic("clickframe")).NextFrame.PresentWithClip Me.HDC, hey.Left, hey.top, 0, 0, hey.Width, hey.Height, 20
        DrawTextRect hey.Left + hey.Width / 2, hey.top + hey.Height + 2, hey.Tag, argb(185, 255, 255, 255), StringAlignmentCenter
        Exit For
        End If
    End If
End If
Next
If hey.name <> lastobj Then lastobj = ""
'==================================================================================

Call DrawEffect(UI, Me.name)

GdipDeleteGraphics UI2
GdipDeleteGraphics UI
CardList.Refresh
FPS = FPS + 1
Me.Refresh
End Sub

Private Sub Fightbutton_Click()

If ChooseCount = 0 Then
If Msgbox("什么也不带？？？！", , "害怕") = 1 Then Exit Sub
End If

If OnLineMode = 1 Then
Me.Hide
OnLine.Connect "free.ngrok.cc", 19168
Do While OnLine.state <> 7
DoEvents
Loop
CloseInfo = True
PostData "c " & NowWorld
PostData "n " & Replace(PlayerName, " ", "_")
Me.Show
DoEvents
InfoBox "成功创建房间，正在等待其它玩家的加入。", , "多人游戏已开启"
If CloseInfo = False Then Exit Sub
End If

Me.Hide
'===============================================
Dim tempUI As Long '临时储存用
CreateAChild FightWindow: GdipCreateFromHDC FightWindow.HDC, tempUI
FightWindow.UI = tempUI: FightWindow.DrawTimer.Enabled = True
Dim s As Integer
s = 0
Open "C:\Monster2\Card.rsdata" For Output As #1
For i = 0 To UBound(IsChoose)
If IsChoose(i) = True Then

If OnLineMode = 0 Then
FightWindow.SetCard s, FBMonster(i).MonsterName, MyMonsterLevel(i): s = s + 1
Else
FightWindow.SetCard s, FBMonster(i).MonsterName, 0: s = s + 1
End If

Print #1, FBMonster(i).MonsterName
End If
Next
Close #1
If OnLineMode <> 0 Then
FightWindow.Text1.Visible = True
FightWindow.SetLevel 50
End If
Unload Me
'===============================================
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
GdipDeleteGraphics UI
GdipDeleteGraphics UI2
End Sub
Private Sub Form_Load()
On Error Resume Next '错误错误快离开~~~
Dad.SetFocus
ChooseCount = 0
For Each ClickArea In Me.Controls
ClickArea.BackStyle = 0 '变成偷懒用的点击检测吧~~~~~
Next

ReDim FBMonster(UBound(MyMonster))
ReDim IsChoose(UBound(MyMonster))
FBMonster(0).LoadMonster MyMonster(0), 0

For i = 1 To UBound(FBMonster)
FBMonster(i).LoadMonster MyMonster(i), 0
Load Cards(i)
With Cards(i)
    .Left = Cards(0).Left + Int(i / 6) * (Cards(0).Width + 20)
    .top = CardsPre(i Mod 6).top
    .Visible = True
    .ZOrder
End With
Next

scrollbutton.Left = scrollbar.Left
scrollcount = Int(UBound(FBMonster) / 6) - 2 + 1
If scrollcount <= 0 Then
scrollbutton.Width = scrollbar.Width
Else
scrollbutton.Width = scrollbar.Width / scrollcount
End If
CardList.Width = (Int(UBound(FBMonster) / 6) + 1) * (Cards(0).Width + 20)

Dim FaMonsters(3) As String
FaMonsters(0) = ReadINI("Level" & NowLevel(NowWorld), "Monster1", App.Path & "\level\world" & NowWorld & ".ini")
FaMonsters(1) = ReadINI("Level" & NowLevel(NowWorld), "Monster2", App.Path & "\level\world" & NowWorld & ".ini")
FaMonsters(2) = ReadINI("Level" & NowLevel(NowWorld), "Monster3", App.Path & "\level\world" & NowWorld & ".ini")
FaMonsters(3) = ReadINI("Level" & NowLevel(NowWorld), "Monster4", App.Path & "\level\world" & NowWorld & ".ini")
ProgressPresent = ReadINI("Level" & NowLevel(NowWorld), "Present", App.Path & "\level\world" & NowWorld & ".ini")
Presents2 = Split(ProgressPresent, ";")
GamePictures(GetPic("Back" & NowWorld + 1)).NextFrame.Present PreviewFrame.HDC, 0, 0
For i = 0 To 3
temp = Split(FaMonsters(i), ";")
For s = 0 To UBound(temp)
GamePictures(GetPic(temp(s) & "0")).NextFrame.Present PreviewFrame.HDC, Int(Rnd * (PreviewFrame.Width - 73)), Int(Rnd * (PreviewFrame.Height - 73))
Next
Next
For i = 0 To UBound(Presents2)
GamePictures(GetPic(PresentIcon(Presents2(i)))).NextFrame.Present PreviewFrame.HDC, PreviewFrame.Width - 48 - i * 10, PreviewFrame.Height - 48
Next
PreviewFrame.Refresh
DrawTimer.Enabled = True

On Error Resume Next
If Dir("C:\Monster2\Card.rsdata") <> "" Then
    Open "C:\Monster2\Card.rsdata" For Input As #1
    Do While Not EOF(1)
    Line Input #1, a
    For i = 0 To UBound(FBMonster)
    If FBMonster(i).MonsterName = a Then Call Cards_Click(Val(i)): Exit For
    Next
    Loop
    Close #1
End If
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

Private Sub PreviewFrame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DownTime = GetTickCount
End Sub

Private Sub scrollbutton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartX = X / 15
End Sub

Private Sub scrollbutton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub SuperFightButton_Click()
If Msgbox("确定要选择超级模式吗？超级模式不计通关且没有任何战利品。", , "超级模式") = 1 Then Exit Sub

If ChooseCount = 0 Then
If Msgbox("什么也不带？？？！", , "害怕") = 1 Then Exit Sub
End If
Me.Hide
'===============================================
Dim tempUI As Long '临时储存用
CreateAChild FightWindow: GdipCreateFromHDC FightWindow.HDC, tempUI
FightWindow.UI = tempUI: FightWindow.DrawTimer.Enabled = True
FightWindow.DebugMode = True
Dim s As Integer
s = 0
For i = 0 To UBound(IsChoose)
If IsChoose(i) = True Then FightWindow.SetCard s, FBMonster(i).MonsterName, 5: s = s + 1
Next
FightWindow.SuperCounts = 9999
FunnyCounts = 9999
Unload Me
'===============================================
End Sub
