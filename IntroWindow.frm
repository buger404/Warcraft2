VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form IntroWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "IntroWindow"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   949
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   6330
      Left            =   1470
      TabIndex        =   0
      Top             =   960
      Width           =   11520
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   20320
      _cy             =   11165
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   285
      Index           =   0
      Left            =   14670
      Top             =   6660
      Width           =   285
   End
End
Attribute VB_Name = "IntroWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
wmp.url = App.Path & "\video\intro.mp4"
End Sub

Private Sub wmp_StatusChange()
If wmp.playState = wmppsStopped Then
Me.Hide
'===============================================
Dim tempUI As Long '临时储存用
CreateAChild FightWindow: GdipCreateFromHDC FightWindow.Hdc, tempUI
FightWindow.UI = tempUI: FightWindow.DrawTimer.Enabled = True
Unload Me
'===============================================
End If
End Sub
