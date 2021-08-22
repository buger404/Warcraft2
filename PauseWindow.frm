VERSION 5.00
Begin VB.Form PauseWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "PauseWindow"
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   486
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer DrawTimer 
      Interval        =   20
      Left            =   6630
      Top             =   360
   End
   Begin VB.Label nextbutton 
      BackColor       =   &H00FFD973&
      Height          =   825
      Left            =   5070
      TabIndex        =   3
      Tag             =   "继续"
      Top             =   2820
      Width           =   825
   End
   Begin VB.Label retrybutton 
      BackColor       =   &H00FFD973&
      Height          =   825
      Left            =   3330
      TabIndex        =   2
      Tag             =   "重新开始"
      Top             =   2820
      Width           =   825
   End
   Begin VB.Label closebutton 
      BackColor       =   &H00FFD973&
      Height          =   825
      Left            =   1530
      TabIndex        =   1
      Tag             =   "退出游戏"
      Top             =   2820
      Width           =   825
   End
   Begin VB.Label PauseText 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   3720
      TabIndex        =   0
      Top             =   1860
      Width           =   45
   End
End
Attribute VB_Name = "PauseWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastobj As String
Private Sub Label1_Click()

End Sub

Private Sub closebutton_Click()
Call FightWindow.pauseframe_MouseUp(2, 0, 0, 0)
Unload Me
End Sub

Private Sub DrawTimer_Timer()
BitBlt Me.HDC, 0, 0, Me.ScaleWidth, _
            Me.ScaleHeight, AeroWindow.HDC, AeroWindow.ToolFrame.Left, AeroWindow.ToolFrame.top, vbSrcCopy

GdipCreateFromHDC Me.HDC, UI

LastPenColor = argb(255, 255, 255, 255)
GdipSetSolidFillColor Brush1, argb(255, 255, 255, 255)

GdipSetStringFormatAlign strformat, StringAlignmentCenter
GdipDrawString UI, StrPtr("已暂停"), -1, curFontBig, NewRectF(PauseText.Left, PauseText.top, 0, 0), strformat, Brush1

GamePictures(GetPic("closebutton")).NextFrame.Present Me.HDC, closebutton.Left, closebutton.top
GamePictures(GetPic("nextbutton")).NextFrame.Present Me.HDC, nextbutton.Left, nextbutton.top
GamePictures(GetPic("retrybutton")).NextFrame.Present Me.HDC, retrybutton.Left, retrybutton.top

LastPenColor = argb(185, 255, 255, 255)
GdipSetSolidFillColor Brush1, argb(185, 255, 255, 255)
Dim p As POINTAPI
GetCursorPos p
p.X = p.X - Dad.Left / 15 - AeroWindow.ToolFrame.Left: p.Y = p.Y - Dad.top / 15 - 16 - AeroWindow.ToolFrame.top
'DrawEffect UI, Me.name
'==================================================================================
' UI ++
On Error Resume Next
For Each hey In Me.Controls
Err.Clear
If hey.name <> Me.name Then
    If p.X >= hey.Left And p.X <= hey.Left + hey.Width And p.Y >= hey.top And p.Y <= hey.top + hey.Height And hey.Visible = True Then
        If Err.Number = 0 Then
        If lastobj <> hey.name Then Ring "move": lastobj = hey.name
        GamePictures(GetPic("clickframe")).NextFrame.PresentWithClip Me.HDC, hey.Left, hey.top, 0, 0, hey.Width, hey.Height, 20
        GdipDrawString UI, StrPtr(hey.Tag), -1, curFont, NewRectF(hey.Left + hey.Width / 2, hey.top + hey.Height + 2, 0, 0), strformat, Brush1
        Exit For
        End If
    End If
End If
Next
If hey.name <> lastobj Then lastobj = ""
'==================================================================================

FPS = FPS + 1


GdipDeleteGraphics UI
Me.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next '错误错误快离开~~~
Dad.SetFocus
For Each ClickArea In Me.Controls
ClickArea.BackStyle = 0 '变成偷懒用的点击检测吧~~~~~
Next
Me.Show
BitBlt Me.HDC, 0, 0, Me.ScaleWidth, _
            Me.ScaleHeight, AeroWindow.HDC, AeroWindow.ScaleWidth / 2 - Me.ScaleWidth / 2, AeroWindow.ScaleHeight / 2 - Me.ScaleHeight / 2, vbSrcCopy

GdipCreateFromHDC Me.HDC, UI

LastPenColor = argb(255, 255, 255, 255)
GdipSetSolidFillColor Brush1, argb(255, 255, 255, 255)

GdipSetStringFormatAlign strformat, StringAlignmentCenter
GdipDrawString UI, StrPtr("已暂停"), -1, curFontBig, NewRectF(PauseText.Left, PauseText.top, 0, 0), strformat, Brush1

GamePictures(GetPic("closebutton")).NextFrame.Present Me.HDC, closebutton.Left, closebutton.top
GamePictures(GetPic("nextbutton")).NextFrame.Present Me.HDC, nextbutton.Left, nextbutton.top
GamePictures(GetPic("retrybutton")).NextFrame.Present Me.HDC, retrybutton.Left, retrybutton.top
FPS = FPS + 1
GdipDeleteGraphics UI
Me.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
DrawTimer.Enabled = False
Dad.SetFocus
GdipDeleteGraphics UI
AeroWindow.Hide
End Sub

Private Sub nextbutton_Click()
Call FightWindow.pauseframe_MouseUp(1, 0, 0, 0)
Unload Me
End Sub

Private Sub retrybutton_Click()
Call FightWindow.pauseframe_MouseUp(4, 0, 0, 0)
Unload Me
End Sub
