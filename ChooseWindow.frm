VERSION 5.00
Begin VB.Form ChooseWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "ChooseWindow"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MouseIcon       =   "ChooseWindow.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   949
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer DrawTimer 
      Interval        =   30
      Left            =   120
      Top             =   150
   End
   Begin VB.Label Maps 
      BackColor       =   &H00F0B000&
      Height          =   1635
      Index           =   4
      Left            =   11100
      TabIndex        =   5
      Tag             =   "难度  ★★★★★★"
      Top             =   3240
      Width           =   1635
   End
   Begin VB.Label Maps 
      BackColor       =   &H00F0B000&
      Height          =   1635
      Index           =   3
      Left            =   8835
      TabIndex        =   4
      Tag             =   "难度  ★★★★★"
      Top             =   3240
      Width           =   1635
   End
   Begin VB.Label backbutton 
      BackColor       =   &H00F0B000&
      Height          =   720
      Left            =   0
      TabIndex        =   3
      Top             =   7620
      Width           =   720
   End
   Begin VB.Label Maps 
      BackColor       =   &H00F0B000&
      Height          =   1635
      Index           =   2
      Left            =   6570
      TabIndex        =   2
      Tag             =   "难度  ★★★"
      Top             =   3240
      Width           =   1635
   End
   Begin VB.Label Maps 
      BackColor       =   &H00F0B000&
      Height          =   1635
      Index           =   1
      Left            =   4305
      TabIndex        =   1
      Tag             =   "难度  ★★"
      Top             =   3240
      Width           =   1635
   End
   Begin VB.Label Maps 
      BackColor       =   &H00F0B000&
      Height          =   1635
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Tag             =   "难度  ★"
      Top             =   3240
      Width           =   1635
   End
End
Attribute VB_Name = "ChooseWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UI As Long, lastobj As String
Dim nowPWorld As Integer
Private Sub backbutton_Click()
Me.Hide
CreateAChild MainWindow
Unload Me
End Sub

Private Sub DrawTimer_Timer()
GdipCreateFromHDC Me.HDC, UI

Dim p As POINTAPI
GetCursorPos p
p.X = p.X - Dad.Left / 15: p.Y = p.Y - Dad.top / 15 - 16

GamePictures(GetPic("Back" & nowPWorld + 1 & "blur")).NextFrame.Present Me.HDC, 0, 0, 80

DrawPicNameControl backbutton, "backbutton"
For i = 0 To Maps.UBound
DrawPicNameControl Maps(i), "map" & i + 1
Next
For i = 0 To UBound(NowLevel)
DrawTextRect Maps(i).Left + 56, Maps(i).top + 68, Int(NowLevel(i) / 20 * 100) & "%", argb(195, 255, 255, 255), StringAlignmentCenter
Next

'==================================================================================
' UI ++
On Error Resume Next
For Each hey In Me.Controls
Err.Clear
If hey.name <> Me.name Then
    If p.X >= hey.Left And p.X <= hey.Left + hey.Width And p.Y >= hey.top And p.Y <= hey.top + hey.Height Then
        If Err.Number = 0 Then
        If lastobj <> hey.name Then Ring "move": lastobj = hey.name
        GamePictures(GetPic("clickframe")).NextFrame.PresentWithClip Me.HDC, hey.Left, hey.top, 0, 0, hey.Width, hey.Height, 20
        DrawTextRect hey.Left + hey.Width / 2, hey.top + hey.Height + 2, hey.Tag, argb(185, 255, 255, 255), StringAlignmentCenter
        If hey.name = "Maps" Then
        'GamePictures(GetPic("nextbutton")).NextFrame.Present Me.HDC, hey.Left + hey.Width - 55, hey.top + hey.Height + 20
        'GamePictures(GetPic("menubutton")).NextFrame.Present Me.HDC, hey.Left + hey.Width - 115, hey.top + hey.Height + 20
        End If
        
        Exit For
        End If
    End If
End If
Next
If hey.name <> lastobj Then lastobj = ""
'==================================================================================


Call DrawEffect(UI, Me.name)

FPS = FPS + 1
Me.Refresh
GdipDeleteGraphics UI
End Sub

Private Sub Form_Load()
On Error Resume Next '错误错误快离开~~~
Dad.SetFocus
For Each ClickArea In Me.Controls
ClickArea.BackStyle = 0 '变成偷懒用的点击检测吧~~~~~
Next
End Sub

Private Sub Maps_Click(Index As Integer)
NowWorld = Index
Me.Hide
CreateAChild CardWindow
Unload Me
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
GdipDeleteGraphics UI
End Sub

Private Sub Maps_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
nowPWorld = Index
End Sub
