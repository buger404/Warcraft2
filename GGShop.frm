VERSION 5.00
Begin VB.Form GGShop 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0049413A&
   BorderStyle     =   0  'None
   Caption         =   "GGShop"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MouseIcon       =   "GGShop.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   949
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer ChangeBack 
      Interval        =   5000
      Left            =   13650
      Top             =   570
   End
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   13650
      Top             =   1050
   End
   Begin VB.PictureBox ListFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFD973&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   390
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   899
      TabIndex        =   2
      Top             =   4260
      Width           =   13485
      Begin VB.PictureBox CardList 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFD973&
         BorderStyle     =   0  'None
         Height          =   2595
         Left            =   90
         ScaleHeight     =   173
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   457
         TabIndex        =   3
         Top             =   120
         Width           =   6855
         Begin VB.Label CardsPre 
            BackColor       =   &H00BAB539&
            Height          =   765
            Index           =   1
            Left            =   0
            TabIndex        =   6
            Top             =   900
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label CardsPre 
            BackColor       =   &H00BAB539&
            Height          =   765
            Index           =   2
            Left            =   0
            TabIndex        =   5
            Top             =   1800
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label Cards 
            BackColor       =   &H00E9E4E9&
            Height          =   765
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   1425
         End
         Begin VB.Label CardsPre 
            BackColor       =   &H00BAB539&
            Height          =   765
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   1425
         End
      End
   End
   Begin VB.Label buybutton 
      Height          =   630
      Left            =   12450
      TabIndex        =   15
      Top             =   3150
      Width           =   1215
   End
   Begin VB.Label needMoney 
      BackColor       =   &H00F0B000&
      Height          =   630
      Left            =   7920
      TabIndex        =   14
      Top             =   1350
      Width           =   630
   End
   Begin VB.Label SpeechText 
      BackColor       =   &H00F0B000&
      Height          =   945
      Left            =   7950
      TabIndex        =   12
      Top             =   2190
      Width           =   5715
   End
   Begin VB.Label PreviewMonster 
      BackColor       =   &H00F0B000&
      Height          =   1095
      Left            =   3480
      TabIndex        =   11
      Top             =   1770
      Width           =   1095
   End
   Begin VB.Label backbutton 
      BackColor       =   &H00F0B000&
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   7620
      Width           =   720
   End
   Begin VB.Label scrollbutton 
      BackColor       =   &H00BAB539&
      Height          =   225
      Left            =   660
      TabIndex        =   8
      Top             =   7230
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label scrollbar 
      BackColor       =   &H00FFD973&
      Height          =   225
      Left            =   300
      TabIndex        =   9
      Top             =   7230
      Visible         =   0   'False
      Width           =   13695
   End
   Begin VB.Label MoneyIcon 
      BackColor       =   &H00F0B000&
      Height          =   630
      Left            =   300
      TabIndex        =   0
      Top             =   450
      Width           =   630
   End
   Begin VB.Label BigFrame 
      BackColor       =   &H00F0B000&
      Height          =   3075
      Left            =   300
      TabIndex        =   10
      Top             =   4140
      Visible         =   0   'False
      Width           =   13695
   End
   Begin VB.Label infoframe 
      BackColor       =   &H00FFD973&
      Height          =   3105
      Left            =   7740
      TabIndex        =   13
      Top             =   840
      Width           =   6135
   End
End
Attribute VB_Name = "GGShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UI As Long, UI2 As Long
Dim DownTime As Long
Dim StartX As Single
Dim Presents() As String, PresentImg() As String
Dim NowShop As Integer
Dim NowBack As String
Public lastobj As String
Private Sub backbutton_Click()
Me.Hide
CreateAChild MainWindow
Unload Me
End Sub


Private Sub buybutton_Click()
Index = NowShop
If Msgbox("确定要买下" & PresentImg(Index) & "吗？" & vbCrLf & "你将要花费" & Presents(Index) & "游戏币", , "购买？") = 1 Then Exit Sub
If Money >= Val(Presents(Index)) Then
Money = Money - Val(Presents(Index))
ShowInfo "花费 " & Presents(Index) & " 游戏币", "money", "coin"
temp = GetPresent(Cards(Index).Tag)
Call WriteSave
Else
Ring "warning"
Msgbox "你的钱不够啦！", , "GG"
End If
End Sub

Private Sub Cards_Click(Index As Integer)
NowShop = Index
End Sub

Private Sub ChangeBack_Timer()
Randomize
a = Int(Rnd * 4)
Select Case a
    Case Is <= 3
    NowBack = "Back" & a + 1 & "blur"
    Case Else
    NowBack = "shopBack"
End Select
End Sub

Private Sub DrawTimer_Timer()
Dim p As POINTAPI
GetCursorPos p
p.X = p.X - Dad.Left / 15: p.Y = p.Y - Dad.top / 15 - 16

GdipCreateFromHDC Me.HDC, UI
GdipCreateFromHDC CardList.HDC, UI2
GdipSetTextRenderingHint UI, TextRenderingHintAntiAlias '平滑的字字
GdipSetTextRenderingHint UI2, TextRenderingHintAntiAlias '平滑的字字

GamePictures(GetPic(NowBack)).NextFrame.Present Me.HDC, 0, 0, 100

FillControl infoframe, "whiteeffect"
FillControl BigFrame, "whiteeffect"
FillControl scrollbar, "whiteeffect"
FillControl scrollbutton, "blueeffect"
FillControl MoneyIcon, "money"
FillControl needMoney, "money"
DrawPicNameControl backbutton, "backbutton"

If GetTickCount Mod 1000 < 100 Then '眨眼动画_(:з」∠)_
DrawPicNameControl PreviewMonster, PresentImg(NowShop) & "1"
Else
DrawPicNameControl PreviewMonster, PresentImg(NowShop) & "0"
End If

DrawPicNameControl buybutton, "buybutton1"


BitBlt ListFrame.HDC, 0, 0, ListFrame.Width, ListFrame.Height, Me.HDC, ListFrame.Left, ListFrame.top, vbSrcCopy
ListFrame.Refresh
CardList.Left = -((scrollbutton.Left - scrollbar.Left) / scrollbutton.Width * (Cards(0).Width + 30)) '* 15
BitBlt CardList.HDC, -CardList.Left, -CardList.top, ListFrame.Width, ListFrame.Height, ListFrame.HDC, 0, 0, vbSrcCopy

        LastPenColor = argb(255, 32, 32, 32)
        GdipSetSolidFillColor Brush1, argb(255, 32, 32, 32)
        For i = 0 To Cards.UBound
            GamePictures(GetPic("blackeffect")).NextFrame.PresentWithClip CardList.HDC, Cards(i).Left, Cards(i).top, 0, 0, Cards(i).Width, Cards(i).Height
            GamePictures(GetPic(PresentImg(i) & "Icon")).NextFrame.Present CardList.HDC, Cards(i).Left + 20, Cards(i).top + 1
        Next

DrawTextRect MoneyIcon.Left + 50, MoneyIcon.top + 12, "持有 " & Money, argb(255, 255, 255, 255), StringAlignmentNear
DrawTextControl SpeechText, Cards(NowShop).ToolTipText, argb(255, 32, 32, 32), StringAlignmentNear
DrawTextRect needMoney.Left + 50, needMoney.top + 12, "花费 " & Presents(NowShop), argb(255, 32, 32, 32), StringAlignmentNear
DrawTextRect infoframe.Left + 15, infoframe.top + 10, PresentImg(NowShop), argb(255, 0, 176, 240), StringAlignmentNear

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
        Exit For
        End If
    End If
End If
Next
If hey.name <> lastobj Then lastobj = ""
'==================================================================================

CardList.Refresh
Call DrawEffect(UI, Me.name)
FPS = FPS + 1
Me.Refresh

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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
GdipDeleteGraphics UI
GdipDeleteGraphics UI2
End Sub
Private Sub Form_Load()
Dad.SetFocus
NowBack = "shopBack"
ReDim Presents(0), PresentImg(0)
temp = Dir(App.Path & "\monster\shop\")
Presents(0) = ReadINI("Present", "Cost", App.Path & "\monster\shop\" & temp)
Cards(0).ToolTipText = ReadINI("Present", "Info", App.Path & "\monster\shop\" & temp)
Cards(0).Tag = ReadINI("Present", "Code", App.Path & "\monster\shop\" & temp)
temp2 = Split(temp, ".")
PresentImg(0) = temp2(0)

Do While temp <> ""
temp = Dir()
If temp = "" Then Exit Do
ReDim Preserve Presents(UBound(Presents) + 1)
ReDim Preserve PresentImg(UBound(Presents))
Load Cards(UBound(Presents))
With Cards(UBound(Presents))
    .Left = Cards(0).Left + Int(UBound(Presents) / 3) * (Cards(0).Width + 20)
    .top = CardsPre(UBound(Presents) Mod 3).top
    .Visible = True
    .ZOrder
End With
Presents(UBound(Presents)) = ReadINI("Present", "Cost", App.Path & "\monster\shop\" & temp)
Cards(UBound(Presents)).ToolTipText = ReadINI("Present", "Info", App.Path & "\monster\shop\" & temp)
Cards(UBound(Presents)).Tag = ReadINI("Present", "Code", App.Path & "\monster\shop\" & temp)
temp2 = Split(temp, ".")
PresentImg(UBound(Presents)) = temp2(0)
DoEvents
Loop

scrollbutton.Left = scrollbar.Left
scrollcount = Int(UBound(Presents) / 3) - 7 + 1
If scrollcount <= 0 Then
scrollbutton.Width = scrollbar.Width
Else
scrollbutton.Width = scrollbar.Width / scrollcount
End If
CardList.Width = (Int(UBound(Presents) / 3) + 1) * (Cards(0).Width + 20)

On Error Resume Next '错误错误快离开~~~
ChooseCount = 0
For Each ClickArea In Me.Controls
ClickArea.BackStyle = 0 '变成偷懒用的点击检测吧~~~~~
Next

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

