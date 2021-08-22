VERSION 5.00
Begin VB.Form InfoWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "InfoWindow"
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MouseIcon       =   "InfoWindow.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   62
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer DrawTimer 
      Interval        =   100
      Left            =   7590
      Top             =   90
   End
   Begin VB.Label PauseText 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   4170
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "InfoWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public infotext As String, InfoPic As String

Private Sub DrawTimer_Timer()
BitBlt Me.HDC, 0, 0, Me.ScaleWidth, _
            Me.ScaleHeight, AeroWindow.HDC, AeroWindow.ScaleWidth / 2 - Me.ScaleWidth / 2, AeroWindow.ScaleHeight / 2 - Me.ScaleHeight / 2, vbSrcCopy

GdipCreateFromHDC Me.HDC, UI

LastPenColor = argb(255, 255, 255, 255)
GdipSetSolidFillColor Brush1, argb(255, 255, 255, 255)

GdipSetStringFormatAlign strformat, StringAlignmentCenter
GdipDrawString UI, StrPtr(infotext), -1, curFontBig, NewRectF(Me.ScaleWidth / 2, Me.ScaleHeight / 2, 0, 0), strformat, Brush1

GamePictures(GetPic(InfoPic)).NextFrame.Present Me.HDC, 10, Me.ScaleHeight / 2 - GamePictures(GetPic(InfoPic)).NextFrame.Height / 4

FPS = FPS + 1
GdipDeleteGraphics UI
Me.Refresh
End Sub

Sub Draw()
Me.Show
BitBlt Me.HDC, 0, 0, Me.ScaleWidth, _
            Me.ScaleHeight, AeroWindow.HDC, AeroWindow.ScaleWidth / 2 - Me.ScaleWidth / 2, AeroWindow.ScaleHeight / 2 - Me.ScaleHeight / 2, vbSrcCopy

GdipCreateFromHDC Me.HDC, UI

LastPenColor = argb(255, 255, 255, 255)
GdipSetSolidFillColor Brush1, argb(255, 255, 255, 255)

GdipSetStringFormatAlign strformat, StringAlignmentCenter
GdipDrawString UI, StrPtr(infotext), -1, curFontBig, NewRectF(Me.ScaleWidth / 2, Me.ScaleHeight / 2, 0, 0), strformat, Brush1

GamePictures(GetPic(InfoPic)).NextFrame.Present Me.HDC, 10, Me.ScaleHeight / 2

FPS = FPS + 1
GdipDeleteGraphics UI
Me.Refresh
End Sub

Private Sub Form_Load()
Dad.SetFocus
End Sub
