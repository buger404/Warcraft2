VERSION 5.00
Begin VB.Form WelcomeWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00262626&
   BorderStyle     =   0  'None
   Caption         =   "WelcomeWindow"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MouseIcon       =   "WelcomeWindow.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   949
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer DrawTimer 
      Interval        =   30
      Left            =   390
      Top             =   420
   End
End
Attribute VB_Name = "WelcomeWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Step As Integer, LastEffect As Integer
Private Sub DrawTimer_Timer()
Dim UI As Long
GdipCreateFromHDC Me.HDC, UI
GamePictures(GetPic("MainBackground")).NextFrame.Present Me.HDC, 0, 0
Select Case Step
Case 0
    If LastEffect > UBound(GameEffect) Then
    LastEffect = NewEffect(Me.HDC, Me.name, Me.ScaleWidth / 2 - 266 / 2, Me.ScaleHeight / 2 - 290 / 2, FadeInPic, "LOGO", 20)
    Step = Step + 1
    End If
Case 1
    If LastEffect > UBound(GameEffect) Then
    LastEffect = NewEffect(Me.HDC, Me.name, Me.ScaleWidth / 2, Me.ScaleHeight / 2, MagicText, "Waitting for load assets 100", 60)
    Step = Step + 1
    End If
Case 2
    If LastEffect > UBound(GameEffect) Then
    If LoadOK = True Then Step = Step + 1
    Else
    If LoadOK = False And GameEffect(LastEffect).EffectCount >= 45 Then GameEffect(LastEffect).EffectCount = 45
    GameEffect(LastEffect).ChangeText "Waitting for load assets " & Int(UBound(GamePictures) / 755 * 100)
    End If
Case 3
    If NowLevel(0) = 0 Then
    NowWorld = 0
    '===============================================
    Dim tempUI As Long '临时储存用
    CreateAChild FightWindow: GdipCreateFromHDC FightWindow.HDC, tempUI
    FightWindow.UI = tempUI: FightWindow.DrawTimer.Enabled = True
    '===============================================
    'FadeIn 20, Dad
    Else
    CreateAChild MainWindow

    'FadeIn 20, Dad
    Wait 1000
    Ring "play"
    
    If Dir("C:\Monster2\Version.txt") <> "" Then LastVersion = Val(ReadFile("C:\Monster2\Version.txt"))
        If LastVersion < 185250 Then '18年份406日期1编译次数
        Open "C:\Monster2\Version.txt" For Output As #1
        Print #1, 185250
        Close #1
        Call ShowLastShow
        Msgbox ReadFile(App.Path & "\Update.txt"), , "更新内容"
        End If
    
    End If
    
    If OwnMonster("云") = False Then GetPresent "m 云"
    Step = Step + 1
    Unload WelcomeWindow
Case 4
    Unload Me
    Exit Sub
End Select

Call DrawEffect(UI, Me.name)
FPS = FPS + 1
Me.Refresh
GdipDeleteGraphics UI
End Sub

Private Sub Form_Load()
LastEffect = NewEffect(Me.HDC, Me.name, Me.ScaleWidth / 2, Me.ScaleHeight / 2, MagicText, "Welcome to Monster Fight 2", 20)
Dad.SetFocus
End Sub
