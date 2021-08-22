VERSION 5.00
Begin VB.Form EquipWindow 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "EquipWindow"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MouseIcon       =   "EquipWindow.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   949
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer DrawTimer 
      Interval        =   30
      Left            =   210
      Top             =   150
   End
End
Attribute VB_Name = "EquipWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Step As Integer, LastEffect As Integer
Private Sub DrawTimer_Timer()
Dim UI As Long
GdipCreateFromHDC Me.HDC, UI
GamePictures(GetPic("MainBackground")).NextFrame.Present Me.HDC, 0, 0
Call DrawEffect(UI, Me.name)


Me.Refresh
GdipDeleteGraphics UI
End Sub

Private Sub Form_Click()
Me.Hide
CreateAChild MainWindow
Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next '´íÎó´íÎó¿ìÀë¿ª~~~
Dad.SetFocus
Step = 0
LastEffect = NewEffect(Me.HDC, Me.name, Me.ScaleWidth / 2, Me.ScaleHeight / 2, MagicText, "It is not time for use !", 30)
End Sub

