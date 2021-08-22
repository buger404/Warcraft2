VERSION 5.00
Begin VB.Form FightWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Fight♂"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MouseIcon       =   "FightWindow.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   949
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer FPSPrinter 
      Interval        =   1000
      Left            =   13650
      Top             =   1980
   End
   Begin VB.Timer FireTimer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   13650
      Top             =   4290
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00262626&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F2F2F2&
      Height          =   285
      Left            =   0
      TabIndex        =   76
      Text            =   "发送信息..."
      Top             =   8070
      Visible         =   0   'False
      Width           =   6315
   End
   Begin VB.Timer FunnyTimer 
      Interval        =   20000
      Left            =   13650
      Top             =   3750
   End
   Begin VB.Timer EffectTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13650
      Top             =   3150
   End
   Begin VB.Timer MoveTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   13650
      Top             =   2550
   End
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   13650
      Top             =   1350
   End
   Begin VB.Label musicbutton 
      BackColor       =   &H00FFD973&
      Height          =   825
      Left            =   9240
      TabIndex        =   82
      Tag             =   "跟着节奏!"
      Top             =   150
      Width           =   825
   End
   Begin VB.Label ratebutton 
      BackColor       =   &H00FFD973&
      Height          =   825
      Left            =   10050
      TabIndex        =   81
      Tag             =   "游戏加速/减速"
      Top             =   150
      Width           =   825
   End
   Begin VB.Label BigSpeak 
      BackColor       =   &H00262626&
      Height          =   2925
      Left            =   0
      TabIndex        =   72
      Top             =   5430
      Visible         =   0   'False
      Width           =   14265
   End
   Begin VB.Label quickbutton 
      BackColor       =   &H00FFD973&
      Height          =   825
      Left            =   8430
      TabIndex        =   71
      Tag             =   "下一波"
      Top             =   150
      Width           =   825
   End
   Begin VB.Label Cards 
      BackColor       =   &H00FFD973&
      Height          =   765
      Index           =   11
      Left            =   6870
      TabIndex        =   80
      Tag             =   "滑稽之花"
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Cards 
      BackColor       =   &H00FFD973&
      Height          =   765
      Index           =   10
      Left            =   5370
      TabIndex        =   79
      Tag             =   "滑稽之花"
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Cards 
      BackColor       =   &H00FFD973&
      Height          =   765
      Index           =   9
      Left            =   3870
      TabIndex        =   78
      Tag             =   "滑稽之花"
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Cards 
      BackColor       =   &H00FFD973&
      Height          =   765
      Index           =   8
      Left            =   2370
      TabIndex        =   77
      Tag             =   "滑稽之花"
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label winFrame 
      Height          =   375
      Left            =   12780
      TabIndex        =   75
      Top             =   1050
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label hpview 
      BackColor       =   &H00FFD973&
      Height          =   825
      Left            =   11670
      TabIndex        =   74
      Tag             =   "显示/隐藏血量"
      Top             =   150
      Width           =   825
   End
   Begin VB.Label Cards 
      BackColor       =   &H00FFD973&
      Height          =   765
      Index           =   7
      Left            =   330
      TabIndex        =   70
      Tag             =   "滑稽之花"
      Top             =   6840
      Width           =   1425
   End
   Begin VB.Label Cards 
      BackColor       =   &H00FFD973&
      Height          =   765
      Index           =   6
      Left            =   330
      TabIndex        =   69
      Tag             =   "GG"
      Top             =   6030
      Width           =   1425
   End
   Begin VB.Label DelButton 
      BackColor       =   &H00FFD973&
      Height          =   825
      Left            =   12480
      TabIndex        =   13
      Tag             =   "删除"
      Top             =   150
      Width           =   825
   End
   Begin VB.Label MoneyText 
      BackColor       =   &H00FFD973&
      Height          =   45
      Left            =   6870
      TabIndex        =   68
      Top             =   990
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Speaking 
      BackColor       =   &H00F0B000&
      Height          =   555
      Left            =   0
      TabIndex        =   8
      Top             =   6660
      Visible         =   0   'False
      Width           =   14205
   End
   Begin VB.Label BOOMButton 
      BackColor       =   &H00FFD973&
      Height          =   975
      Left            =   6780
      TabIndex        =   17
      Top             =   7290
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label InfoText 
      BackColor       =   &H00FFD973&
      Height          =   165
      Left            =   6540
      TabIndex        =   16
      Top             =   -90
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label SuperFunnyFrame 
      BackColor       =   &H00FFD973&
      Height          =   825
      Left            =   10860
      TabIndex        =   15
      Tag             =   "滑稽能量"
      Top             =   150
      Width           =   825
   End
   Begin VB.Label MoneyIcon 
      BackColor       =   &H00FFD973&
      Height          =   75
      Left            =   7800
      TabIndex        =   14
      Top             =   990
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label PauseButton 
      BackColor       =   &H00FFD973&
      Height          =   825
      Left            =   13290
      TabIndex        =   12
      Tag             =   "暂停"
      Top             =   150
      Width           =   825
   End
   Begin VB.Label LevelText 
      BackColor       =   &H00F0B000&
      Height          =   375
      Left            =   8760
      TabIndex        =   11
      Top             =   7860
      Width           =   5175
   End
   Begin VB.Label LevelProgress 
      BackColor       =   &H00FFD973&
      Height          =   195
      Left            =   8940
      TabIndex        =   10
      Top             =   7440
      Width           =   4635
   End
   Begin VB.Label LevelFrame 
      BackColor       =   &H00F0B000&
      Height          =   375
      Left            =   8670
      TabIndex        =   9
      Top             =   7350
      Width           =   5175
   End
   Begin VB.Label Cards 
      BackColor       =   &H00F0B000&
      Height          =   765
      Index           =   5
      Left            =   330
      TabIndex        =   7
      Tag             =   "SWL"
      Top             =   5220
      Width           =   1425
   End
   Begin VB.Label Cards 
      BackColor       =   &H00F0B000&
      Height          =   765
      Index           =   4
      Left            =   330
      TabIndex        =   6
      Tag             =   "地狱犬"
      Top             =   4410
      Width           =   1425
   End
   Begin VB.Label Cards 
      BackColor       =   &H00F0B000&
      Height          =   765
      Index           =   3
      Left            =   330
      TabIndex        =   5
      Tag             =   "谜草"
      Top             =   3600
      Width           =   1425
   End
   Begin VB.Label Cards 
      BackColor       =   &H00F0B000&
      Height          =   765
      Index           =   2
      Left            =   330
      TabIndex        =   4
      Tag             =   "神烦狗"
      Top             =   2790
      Width           =   1425
   End
   Begin VB.Label Cards 
      BackColor       =   &H00F0B000&
      Height          =   765
      Index           =   1
      Left            =   330
      TabIndex        =   3
      Tag             =   "阳光加冰"
      Top             =   1980
      Width           =   1425
   End
   Begin VB.Label Cards 
      BackColor       =   &H00F0B000&
      Height          =   765
      Index           =   0
      Left            =   330
      TabIndex        =   2
      Tag             =   "黑嘴"
      Top             =   1170
      Width           =   1425
   End
   Begin VB.Label FunnyCount 
      BackColor       =   &H00FFD973&
      Height          =   285
      Left            =   630
      TabIndex        =   1
      Top             =   435
      Width           =   1425
   End
   Begin VB.Label FunnyFrame 
      BackColor       =   &H00F0B000&
      Height          =   660
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   2010
   End
   Begin VB.Label pauseframe 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   5850
      TabIndex        =   73
      Top             =   750
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   44
      Left            =   11640
      TabIndex        =   67
      Top             =   5910
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   43
      Left            =   10500
      TabIndex        =   66
      Top             =   5910
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   42
      Left            =   9360
      TabIndex        =   65
      Top             =   5910
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   41
      Left            =   8220
      TabIndex        =   64
      Top             =   5910
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   40
      Left            =   7080
      TabIndex        =   63
      Top             =   5910
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   39
      Left            =   5940
      TabIndex        =   62
      Top             =   5910
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   38
      Left            =   4800
      TabIndex        =   61
      Top             =   5910
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   37
      Left            =   3660
      TabIndex        =   60
      Top             =   5910
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   36
      Left            =   2520
      TabIndex        =   59
      Top             =   5910
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   35
      Left            =   11640
      TabIndex        =   58
      Top             =   4770
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   34
      Left            =   10500
      TabIndex        =   57
      Top             =   4770
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   33
      Left            =   9360
      TabIndex        =   56
      Top             =   4770
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   32
      Left            =   8220
      TabIndex        =   55
      Top             =   4770
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   31
      Left            =   7080
      TabIndex        =   54
      Top             =   4770
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   30
      Left            =   5940
      TabIndex        =   53
      Top             =   4770
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   29
      Left            =   4800
      TabIndex        =   52
      Top             =   4770
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   28
      Left            =   3660
      TabIndex        =   51
      Top             =   4770
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   27
      Left            =   2520
      TabIndex        =   50
      Top             =   4770
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   26
      Left            =   11640
      TabIndex        =   49
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   25
      Left            =   10500
      TabIndex        =   48
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   24
      Left            =   9360
      TabIndex        =   47
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   23
      Left            =   8220
      TabIndex        =   46
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   22
      Left            =   7080
      TabIndex        =   45
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   21
      Left            =   5940
      TabIndex        =   44
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   20
      Left            =   4800
      TabIndex        =   43
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   19
      Left            =   3660
      TabIndex        =   42
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   18
      Left            =   2520
      TabIndex        =   41
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   17
      Left            =   11640
      TabIndex        =   40
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   16
      Left            =   10500
      TabIndex        =   39
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   15
      Left            =   9360
      TabIndex        =   38
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   14
      Left            =   8220
      TabIndex        =   37
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   13
      Left            =   7080
      TabIndex        =   36
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   12
      Left            =   5940
      TabIndex        =   35
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   11
      Left            =   4800
      TabIndex        =   34
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   10
      Left            =   3660
      TabIndex        =   33
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   9
      Left            =   2520
      TabIndex        =   32
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   8
      Left            =   11640
      TabIndex        =   31
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   7
      Left            =   10500
      TabIndex        =   30
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   6
      Left            =   9360
      TabIndex        =   29
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   5
      Left            =   8220
      TabIndex        =   28
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   4
      Left            =   7080
      TabIndex        =   27
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   3
      Left            =   5940
      TabIndex        =   26
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   2
      Left            =   4800
      TabIndex        =   25
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   1
      Left            =   3660
      TabIndex        =   24
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Friends 
      BackColor       =   &H00404040&
      Height          =   1095
      Index           =   0
      Left            =   2520
      TabIndex        =   23
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Dogs 
      BackColor       =   &H001D1D1D&
      Height          =   1095
      Index           =   4
      Left            =   1380
      TabIndex        =   22
      Top             =   5910
      Width           =   1095
   End
   Begin VB.Label Dogs 
      BackColor       =   &H001D1D1D&
      Height          =   1095
      Index           =   3
      Left            =   1380
      TabIndex        =   21
      Top             =   4770
      Width           =   1095
   End
   Begin VB.Label Dogs 
      BackColor       =   &H001D1D1D&
      Height          =   1095
      Index           =   2
      Left            =   1380
      TabIndex        =   20
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Dogs 
      BackColor       =   &H001D1D1D&
      Height          =   1095
      Index           =   1
      Left            =   1380
      TabIndex        =   19
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Label Dogs 
      BackColor       =   &H001D1D1D&
      Height          =   1095
      Index           =   0
      Left            =   1380
      TabIndex        =   18
      Top             =   1350
      Width           =   1095
   End
End
Attribute VB_Name = "FightWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UI As Long
Dim RfWidth As Long, RfHeight As Long
Dim WaitMonster As Integer, WaitIndex As Integer  '临时变量
Public NowProgress As Integer, MaxProgress As Integer, BigProgress As Integer, ProgressMonster As Long, ProgressDuring As Single '进度数据
Public LastProgress As Long, PauseProgress As Long, FirstDuring As Single
Dim ProgressPresent As String, ProgressMonsters(3) As String '进度数据
Public DebugMode As Boolean

Dim lastobj As String
Public FirstDanger As Boolean, FirstDangerTime As Long '魔兽警报数据
Public SuperCounts As Long, WaitSuper As Boolean
Dim Fires() As New FireManager  '蛋蛋蛋蛋药
Dim FaceName As String, Speach As String
Dim DelMode As Boolean, IsHPView As Boolean
Dim MyLevel As New LevelManager
Dim LastFPS As Long
Public MusicMode As Boolean
Public SpicalBack As String


Private Sub BigSpeak_Click()
BigSpeak.Visible = False
End Sub

Private Sub Cards_Click(Index As Integer)
If MCards(Index).name = "" Then Exit Sub
If Sets(3) = False Then On Error GoTo ShowErr
If DelMode = True Or WaitSuper = True Then Exit Sub
If WaitMonster = Index Then WaitMonster = -1: Exit Sub
If MCards(Index).NowCD >= MCards(Index).CDTime And FunnyCounts >= MCards(Index).Spend Then
WaitMonster = Index: WaitIndex = -1 'With Wait
ElseIf MCards(Index).NowCD < MCards(Index).CDTime Then
Ring "warning"
NewMsg "卡片正在冷却哦"
Exit Sub
End If
If FunnyCounts < MCards(Index).Spend Then
Ring "warning"
NewMsg "你的滑稽果不够啦"
Exit Sub
End If
ShowErr:
If Err.Number <> 0 Then Call ShowError("Cards_Click")
End Sub



Private Sub Command1_Click()

End Sub

Private Sub DelButton_Click()
If Sets(3) = False Then
    If Sets(4) = False Then
    On Error Resume Next
    Else
    On Error GoTo ShowErr
    End If
End If
If WaitSuper = True Or WaitMonster <> -1 Then Exit Sub
DelMode = Not DelMode
ShowErr:
If Err.Number <> 0 Then Call ShowError("DelButton_Click")
End Sub

Private Sub EffectTimer_Timer()
If Sets(3) = False Then
    If Sets(4) = False Then
    On Error Resume Next
    Else
    On Error GoTo ShowErr
    End If
End If

Call MonsterBorn

For i = 0 To UBound(MCards)
If MCards(i).NowCD < MCards(i).CDTime Then MCards(i).NowCD = MCards(i).NowCD + 1
Next

If UBound(EMonster) > 0 Then
    For i = 1 To UBound(EMonster)
    If i > UBound(EMonster) Then Exit For
    EMonster(i).UpdateEffect
    Next
End If

ShowErr:
If Err.Number <> 0 Then Call ShowError("EffectTimer_Timer")
End Sub



Private Sub FireTimer_Timer()
If Sets(3) = False Then
    If Sets(4) = False Then
    On Error Resume Next
    Else
    On Error GoTo ShowErr
    End If
End If
For i = 0 To UBound(FMonster)
    If i <= UBound(FMonster) Then
    FMonster(i).Update
    End If
Next

ShowErr:
If Err.Number <> 0 Then Call ShowError("FireTimer_Timer")
End Sub

Private Sub Form_Activate()
On Error Resume Next
'If OnLineMode = 0 Then Dad.SetFocus

End Sub
Sub SetCard(Index As Integer, MonsterN As String, Optional ByVal NLevel As Long = 0)
    MCards(Index).name = MonsterN
    MCards(Index).CDTime = Val(ReadINI("Monster", "CD", App.Path & "\monster\fmonster\" & MCards(Index).name & ".ini"))
    MCards(Index).Spend = Val(ReadINI("Monster", "Spend", App.Path & "\monster\fmonster\" & MCards(Index).name & ".ini"))
    MCards(Index).Level = NLevel
    If DebugMode = True Then
    MCards(Index).CDTime = 0
    MCards(Index).Spend = 0
    End If
    MCards(Index).NowCD = MCards(Index).CDTime
    Cards(Index).Tag = MonsterN
End Sub
Sub SetLevel(ByVal Levels As Integer)
Set MyLevel = Nothing
MyLevel.Level = Levels
MyLevel.World = NowWorld
NowProgress = 0
MaxProgress = Val(ReadINI("Level" & Levels, "AttackCount", App.Path & "\level\world" & NowWorld & ".ini"))
BigProgress = Val(ReadINI("Level" & Levels, "AttackExEvery", App.Path & "\level\world" & NowWorld & ".ini"))
ProgressMonster = Val(ReadINI("Level" & Levels, "MonsterCount", App.Path & "\level\world" & NowWorld & ".ini"))
ProgressDuring = Val(ReadINI("Level" & Levels, "AttackDuring", App.Path & "\level\world" & NowWorld & ".ini"))
FirstDuring = Val(ReadINI("Level" & Levels, "FirstAttack", App.Path & "\level\world" & NowWorld & ".ini"))
ProgressPresent = ReadINI("Level" & Levels, "Present", App.Path & "\level\world" & NowWorld & ".ini")
For i = 0 To 3
ProgressMonsters(i) = ReadINI("Level" & Levels, "Monster" & i + 1, App.Path & "\level\world" & NowWorld & ".ini")
Next
LastProgress = GetTickCount
End Sub
Private Sub Form_Load()
Dad.SetFocus
If Sets(3) = False Then
    If Sets(4) = False Then
    On Error Resume Next
    Else
    On Error GoTo ShowErr
    End If
End If
SpicalBack = ""
Text1.Visible = False
DebugMode = False
Set MyLevel = Nothing
MusicMode = False
MyLevel.Level = NowLevel(NowWorld)
MyLevel.World = NowWorld
NowProgress = 0
PaintHDC = 0
MaxProgress = Val(ReadINI("Level" & NowLevel(NowWorld), "AttackCount", App.Path & "\level\world" & NowWorld & ".ini"))
BigProgress = Val(ReadINI("Level" & NowLevel(NowWorld), "AttackExEvery", App.Path & "\level\world" & NowWorld & ".ini"))
ProgressMonster = Val(ReadINI("Level" & NowLevel(NowWorld), "MonsterCount", App.Path & "\level\world" & NowWorld & ".ini"))
ProgressDuring = Val(ReadINI("Level" & NowLevel(NowWorld), "AttackDuring", App.Path & "\level\world" & NowWorld & ".ini"))
FirstDuring = Val(ReadINI("Level" & NowLevel(NowWorld), "FirstAttack", App.Path & "\level\world" & NowWorld & ".ini"))
ProgressPresent = ReadINI("Level" & NowLevel(NowWorld), "Present", App.Path & "\level\world" & NowWorld & ".ini")
For i = 0 To 3
ProgressMonsters(i) = ReadINI("Level" & NowLevel(NowWorld), "Monster" & i + 1, App.Path & "\level\world" & NowWorld & ".ini")
Next
LastProgress = GetTickCount
'===========================BaBaBaBaBa前方无异常===========================
ReDim Fires(0): ReDim EMonster(0): ReDim Funnys(0): ReDim Presents(0): FunnyCounts = AtFirstFunny: WaitMonster = -1: SuperCounts = 1
ReDim FMonster(0)
RfWidth = Me.ScaleWidth: RfHeight = Me.ScaleHeight
DelMode = False: IsHPView = False: WaitSuper = False
For i = 0 To UBound(MCards)
MCards(i).name = ""
Cards(i).Tag = ""
Next
'====================================================================

ShowErr:
If Err.Number <> 0 Then Call ShowError("Form_Load")

On Error Resume Next '错误错误快离开~~~
For Each ClickArea In Me.Controls
ClickArea.BackStyle = 0 '变成偷懒用的点击检测吧~~~~~
Next
Err.Clear
MainBGM = True
Wait 1000
With BGM
    .StopMusic
    .LoadMusic App.Path & "\music\Background" & NowWorld + 1 & ".mp3"
    .Play
End With
FMonster(0).MonsterName = "Angle": FMonster(0).Speed = 10: FMonster(0).MonsterType = OhAngle '设置我们的小天使【噗】
FPSPrinter.Enabled = True

BigSpeak.ZOrder
FirstDanger = False: FirstDangerTime = 0: MoveTimer.Enabled = True: EffectTimer.Enabled = True: FunnyTimer.Enabled = True: FireTimer.Enabled = True
End Sub
Sub CheatSpace()
On Error Resume Next
CloseCEandOD
SaveSetting LockString(FuckReg1, "! you are cheatting !", 1), LockString(FuckReg2, "! you are cheatting !", 1), "0", "0"
Ring "error"
FirstDuring = 0
ProgressDuring = 0
CheatTime = GetTickCount
End Sub
Private Sub DrawTimer_Timer()
If Sets(3) = False Then
    If Sets(4) = False Then
    On Error Resume Next
    Else
    On Error GoTo ShowErr
    End If
End If

Dim ReOnce As Boolean
ReFire:
If UBound(Fires) > 0 Then
For i = 1 To UBound(Fires)
If i <= UBound(Fires) Then
Fires(i).X = Fires(i).X + Fires(i).Speed
If Fires(i).X >= Me.ScaleWidth + 73 Or Fires(i).X < -73 Then
Set Fires(i) = Fires(UBound(Fires))
ReDim Preserve Fires(UBound(Fires) - 1)
i = i - 1
End If
End If
Next
End If
If BGM.Rate = 88200 And ReOnce = False Then ReOnce = True: GoTo ReFire
    
    GdipCreateFromHDC Me.HDC, UI
    GdipSetTextRenderingHint UI, TextRenderingHintAntiAlias '平滑的字字

    Dim p As POINTAPI
    GetCursorPos p: p.X = p.X - Dad.Left / 15: p.Y = p.Y - Dad.top / 15 - 16 '获得鼠标坐标

    'GdipGraphicsClear UI, argb(255, 0, 0, 0) '清空画布~~~~~
    
    If SpicalBack = "" Then
    GamePictures(GetPic("Back" & NowWorld + 1)).NextFrame.Present Me.HDC, 0, 0, 100 '绘制背景图
    Else
    GamePictures(GetPic(SpicalBack)).NextFrame.Present Me.HDC, 0, 0, 100 '绘制关卡特别指定的背景图
    End If
    
    
    '=============================================================
    '绘制各种图标
    DrawPicNameControl PauseButton, "pause"                                      '暂停按钮岂不滑至？
    DrawPicNameControl FunnyFrame, "funnyframe"                             '滑稽框不可或缺
    DrawPicNameControl MoneyIcon, "money"                                      '没有钱钱无法战斗
    DrawPicNameControl SuperFunnyFrame, "superframe"                    '超级能量自然不可少
    DrawPicNameControl DelButton, "dustbin"                                       '没有垃圾桶还能行么
    DrawPicNameControl LevelFrame, "progressback"                         '进度条
    DrawPicNameControl quickbutton, "quickbutton"
    If BGM.Rate = 88200 Then
    DrawPicNameControl ratebutton, "quickly"
    Else
    DrawPicNameControl ratebutton, "slowly"
    End If
    If MusicMode = True Then
    DrawPicNameControl musicbutton, "musicbutton1"
    Else
    DrawPicNameControl musicbutton, "musicbutton0"
    End If
    '=============================================================
    
    '=============================================================
    '进度条
    GamePictures(GetPic("progressbar")).NextFrame.PresentWithClip Me.HDC, LevelProgress.Left, LevelProgress.top + 1.5, 0, 0, _
                                                                        NowProgress / MaxProgress * LevelProgress.Width, LevelProgress.Height * 2
    GamePictures(GetPic("progressball")).NextFrame.PresentWithClip Me.HDC, LevelProgress.Left + NowProgress / MaxProgress * LevelProgress.Width, LevelProgress.top - 5.5, 0, 0, _
                                                                        90, 90
    '=============================================================
    
    '=============================================================
    '血量显示
    If IsHPView = True Then
    DrawPicNameControl hpview, "hpview0"
    Else
    DrawPicNameControl hpview, "hpview1"
    End If
    '=============================================================
    
    '=============================================================
    '扫黄车动画
    For i = 0 To 4
        If FPS Mod 4 <> 0 Then
        DrawPicNameControl Dogs(i), "Angle0"
        Else
        DrawPicNameControl Dogs(i), "Angle1"
        End If
    Next
    '=============================================================
    
    '=============================================================
    '绘制友好类型魔兽
    If UBound(FMonster) > 0 Then
    For i = 1 To UBound(FMonster)
            If FMonster(i).SpicalPic <> "" Then DrawPicNameMonster FMonster(i), FMonster(i).SpicalPic: GoTo NextOne '如果存在指明的图片则画出
            If GetTickCount - FMonster(i).LastFireTime < 1500 Then DrawPicNameMonster FMonster(i), FMonster(i).MonsterName & "Attack": GoTo NextOne
            '↑攻击状态
            If (GetTickCount - FMonster(i).LastFireTime) Mod 1000 < 100 Then '眨眼动画_(:з」∠)_
            DrawPicNameMonster FMonster(i), FMonster(i).MonsterName & "1"
            Else
            DrawPicNameMonster FMonster(i), FMonster(i).MonsterName & "0"
            End If
            
NextOne:     '好，下一个~
            If FMonster(i).AttackEx = True Then DrawPicNameMonster FMonster(i), "attack"
            If FMonster(i).SuperMode = True Then DrawPicNameMonster FMonster(i), "supereffect"
            If FMonster(i).NeedWater = True Then GamePictures(GetPic("NeedWater")).NextFrame.Present Me.HDC, FMonster(i).X + 66, FMonster(i).Y - 10
    Next
    End If
    '=============================================================
    
    '=============================================================
    '绘制敌对类型魔兽
    If UBound(EMonster) > 0 Then '是否有魔兽需要被绘制
        For i = 1 To UBound(EMonster)
            If FPS Mod 6 > 2 Then '交替动画
                If EMonster(i).Eating = True Then '在吃东西么
                GamePictures(GetPic(EMonster(i).MonsterName & "Attack0")).NextFrame.Present Me.HDC, EMonster(i).X, EMonster(i).Y - EMonster(i).H * 73
                Else
                GamePictures(GetPic(EMonster(i).MonsterName & "0")).NextFrame.Present Me.HDC, EMonster(i).X, EMonster(i).Y - EMonster(i).H * 73
                End If
                If GetTickCount - EMonster(i).IceTime <= 5000 Then GamePictures(GetPic("ice0")).NextFrame.Present Me.HDC, EMonster(i).X, EMonster(i).Y
                If EMonster(i).ChemicalMode = True Or EMonster(i).Speed < EMonster(i).MaxSpeed Then GamePictures(GetPic("chemical0")).NextFrame.Present Me.HDC, EMonster(i).X, EMonster(i).Y
            Else
                If EMonster(i).Eating = True Then '在吃东西么
                GamePictures(GetPic(EMonster(i).MonsterName & "Attack1")).NextFrame.Present Me.HDC, EMonster(i).X, EMonster(i).Y - EMonster(i).H * 73
                Else
                GamePictures(GetPic(EMonster(i).MonsterName & "1")).NextFrame.Present Me.HDC, EMonster(i).X, EMonster(i).Y - EMonster(i).H * 73
                End If
                If GetTickCount - EMonster(i).IceTime <= 5000 Then GamePictures(GetPic("ice1")).NextFrame.Present Me.HDC, EMonster(i).X, EMonster(i).Y
                If EMonster(i).ChemicalMode = True Or (EMonster(i).Speed < EMonster(i).MaxSpeed And MusicMode = False) Then GamePictures(GetPic("chemical1")).NextFrame.Present Me.HDC, EMonster(i).X, EMonster(i).Y
            End If
            If GetTickCount - EMonster(i).FireTime <= 5000 Then GamePictures(GetPic("fire0")).NextFrame.Present Me.HDC, EMonster(i).X, EMonster(i).Y
            If EMonster(i).DarkMode = True Then GamePictures(GetPic("dark0")).NextFrame.Present Me.HDC, EMonster(i).X, EMonster(i).Y
            If EMonster(i).ThunderMode = True Then GamePictures(GetPic("thunder0")).NextFrame.Present Me.HDC, EMonster(i).X, EMonster(i).Y
        Next
    End If
    '=============================================================
    
    If UBound(Funnys) > 0 Then
        For i = 1 To UBound(Funnys)
            If i > UBound(Funnys) Then Exit For
            GamePictures(GetPic("funny")).NextFrame.Present Me.HDC, Funnys(i).X, Funnys(i).Y
            If Funnys(i).MoveY = 0 Then
                If p.X >= Funnys(i).X And p.Y >= Funnys(i).Y And p.X <= Funnys(i).X + 48 And p.Y <= Funnys(i).Y + 48 Then
                Funnys(i).MoveX = (FunnyFrame.Left - Funnys(i).X) / 10
                Funnys(i).MoveY = (FunnyFrame.top - Funnys(i).Y) / 10
                Ring "get"
                NewEffect Me.HDC, Me.name, Funnys(i).X, Funnys(i).Y, FadeInPic, "lighting", 5
                End If
                Funnys(i).Y = Funnys(i).Y + 2
                If Funnys(i).Y > Me.ScaleHeight Then Funnys(i) = Funnys(UBound(Funnys)): ReDim Preserve Funnys(UBound(Funnys) - 1)
            Else
                Funnys(i).Y = Funnys(i).Y + Funnys(i).MoveY
                Funnys(i).X = Funnys(i).X + Funnys(i).MoveX
                Funnys(i).MoveCount = Funnys(i).MoveCount + 1
                If Funnys(i).MoveCount = 10 Then Funnys(i) = Funnys(UBound(Funnys)): FunnyCounts = FunnyCounts + 1: ReDim Preserve Funnys(UBound(Funnys) - 1)
            End If
        Next
    End If
    
    If UBound(Presents) > 0 Then
        For i = 1 To UBound(Presents)
            If i > UBound(Presents) Then Exit For
            GamePictures(GetPic(Presents(i).Icon)).NextFrame.Present Me.HDC, Presents(i).X, Presents(i).Y
            If p.X >= Presents(i).X And p.Y >= Presents(i).Y And p.X <= Presents(i).X + 48 And p.Y <= Presents(i).Y + 48 Then
            ptt = GetPresent(Presents(i).Code)
            NewEffect Me.HDC, Me.name, p.X, p.Y, MagicText, ptt, 10
            NewEffect Me.HDC, Me.name, p.X, p.Y, FadeInPic, "lighting", 5
            Ring "get"
            Presents(i) = Presents(UBound(Presents)): ReDim Preserve Presents(UBound(Presents) - 1)
            End If
        Next
    End If
    
    If UBound(Fires) > 0 Then
        For i = 1 To UBound(Fires)
        If Fires(i).MonsterName <> "" Then GamePictures(GetPic(Fires(i).MonsterName & "Fire")).NextFrame.Present Me.HDC, Fires(i).X, Fires(i).Y
        Next
    End If
    
    If UBound(EMonster) > 0 And IsHPView = True Then
        For i = 1 To UBound(EMonster)
        GamePictures(GetPic("blackheart")).NextFrame.PresentWithClip Me.HDC, EMonster(i).X, EMonster(i).Y + 73 - 32, 0, 0, 73, 31
        GamePictures(GetPic("whiteheart")).NextFrame.PresentWithClip Me.HDC, EMonster(i).X, EMonster(i).Y + 73 - 32, 0, 0, 73, 31 - (EMonster(i).HP / EMonster(i).MaxHP * 31)
        Next
    End If
    
    If UBound(FMonster) > 0 And IsHPView = True Then
        For i = 1 To UBound(FMonster)
        If FMonster(i).MonsterName <> "" Then
        GamePictures(GetPic("blackheart")).NextFrame.PresentWithClip Me.HDC, FMonster(i).X, FMonster(i).Y + 73 - 32, 0, 0, 73, 31
        GamePictures(GetPic("whiteheart")).NextFrame.PresentWithClip Me.HDC, FMonster(i).X, FMonster(i).Y + 73 - 32, 0, 0, 73, 31 - (FMonster(i).HP / FMonster(i).MaxHP * 31)
        End If
        Next
    End If
    
    Dim Fuck2 As Boolean
    
    For i = 0 To Cards.UBound
    If MCards(i).name <> "" Then
        Fuck2 = (WaitMonster <> -1 Or MCards(i).NowCD < MCards(i).CDTime Or FunnyCounts < MCards(i).Spend Or WaitSuper = True)
        GamePictures(GetPic("card1")).NextFrame.Present Me.HDC, Cards(i).Left, Cards(i).top, IIf(Fuck2 = True, 80, 255)
        GamePictures(GetPic(MCards(i).name & "Icon")).NextFrame.Present Me.HDC, Cards(i).Left + 20, Cards(i).top + 1, IIf(Fuck2 = True, 80, 255)
        GamePictures(GetPic("level" & MCards(i).Level)).NextFrame.Present Me.HDC, Cards(i).Left + Cards(i).Width - 35, Cards(i).top - 5, 180
        'DrawTextRect Cards(i).Left + 5, Cards(i).top + 1, "Lv. " & MCards(i).Level, argb(100, 242, 242, 242), StringAlignmentNear, False
        'GamePictures(GetPic("blackeffect")).NextFrame.PresentWithClip Me.HDC, Cards(i).Left, _
                            Cards(i).top, 0, 0, Cards(i).Width, Cards(i).Height
        
        If MCards(i).NowCD < MCards(i).CDTime Then
        GamePictures(GetPic("blueeffect2")).NextFrame.PresentWithClip Me.HDC, Cards(i).Left + 3, _
                            Cards(i).top + Cards(i).Height - 7, _
                            0, 0, (MCards(i).NowCD / MCards(i).CDTime) * (Cards(i).Width - 6), 3
        Else
        GamePictures(GetPic("monsterframe")).NextFrame.Present Me.HDC, Cards(i).Left - Cards(i).Width + 60, Cards(i).top, IIf(Fuck2 = True, 80, 180)
        DrawTextRect Cards(i).Left + 32.5, Cards(i).top + Cards(i).Height - 20.5, MCards(i).Spend, argb(255, 32, 32, 32), StringAlignmentCenter, False
        End If
        
    End If
    Next
    
    If WaitMonster <> -1 Then
    If WaitIndex <> -1 Then
        'If mylevel.CanNew( Friends(WaitIndex).Left, Friends(WaitIndex).top,monster2(WaitIndex).)  Then
        DrawRectangleRect Friends(0).Left, Friends(WaitIndex).top, 76 * 9 - 3, 73, argb(100, 255, 255, 255)
        DrawRectangleRect Friends(WaitIndex).Left, Friends(0).top, 73, 76 * 5 - 3, argb(100, 255, 255, 255)
        'Else
        'DrawRectangleRect Friends(0).Left, Friends(WaitIndex).top, 76 * 9 - 3, 73, argb(100, 255, 0, 0)
        'DrawRectangleRect Friends(WaitIndex).Left, Friends(0).top, 73, 76 * 5 - 3, argb(100, 255, 0, 0)
        'End If
    End If
    GamePictures(GetPic(MCards(WaitMonster).name & "0")).NextFrame.Present Me.HDC, p.X - 36, p.Y - 36, 180
    End If
    
    If WaitSuper = True Then
    GamePictures(GetPic("SuperFunny")).NextFrame.Present Me.HDC, p.X - 24, p.Y - 24, 180
    End If
    
    If DelMode = True Then
    GamePictures(GetPic("dustbin")).NextFrame.Present Me.HDC, p.X - 24, p.Y - 24, 180
    End If
    
    If GetTickCount - Val(Speaking.Tag) <= 10500 Then
        If GetTickCount - Val(Speaking.Tag) >= 10000 Then
        GamePictures(GetPic("blackeffect")).NextFrame.PresentWithClip Me.HDC, Speaking.Left, Speaking.top, 0, 0, _
                            Speaking.Width, Speaking.Height, 255 - (GetTickCount - Val(Speaking.Tag) - 10000) / 500 * 255
        ElseIf GetTickCount - Val(Speaking.Tag) <= 500 Then
            GamePictures(GetPic("blackeffect")).NextFrame.PresentWithClip Me.HDC, Speaking.Left, Speaking.top, 0, 0, _
                                        Speaking.Width, Speaking.Height, (GetTickCount - Val(Speaking.Tag)) / 500 * 255
            Else
            GamePictures(GetPic("blackeffect")).NextFrame.PresentWithClip Me.HDC, Speaking.Left, Speaking.top, 0, 0, _
                                                                                                                                Speaking.Width, Speaking.Height
        End If
        DrawTextRect Speaking.Left + Speaking.Width / 2, Speaking.top + 7.5, Speaking.ToolTipText, argb(255, 255, 255, 255), StringAlignmentCenter, False
        Else
        Speaking.Visible = False
    End If
    
    DrawTextControl FunnyCount, FunnyCounts, argb(255, 32, 32, 32), StringAlignmentCenter
    DrawTextControl MoneyText, format(Money, "0.00"), argb(255, 255, 255, 255), StringAlignmentCenter
    If NowProgress = 0 Then
        If Round((FirstDuring - (GetTickCount - LastProgress)) / 1000) > 0 Then
        DrawTextControl LevelText, WorldName(NowWorld) & " - 第" & NowLevel(NowWorld) + 1 & "天  距离下一波还有" & Round((FirstDuring - (GetTickCount - LastProgress)) / 1000) & "秒", argb(128, 255, 255, 255), StringAlignmentFar
        Else
        DrawTextControl LevelText, WorldName(NowWorld) & " - 第" & NowLevel(NowWorld) + 1 & "天  魔兽正在到来...", argb(128, 255, 255, 255), StringAlignmentFar
        End If
        
    ElseIf NowProgress < MaxProgress Then
        If Round((ProgressDuring - (GetTickCount - LastProgress)) / 1000) > 0 Then
        DrawTextControl LevelText, WorldName(NowWorld) & " - 第" & NowLevel(NowWorld) + 1 & "天  距离下一波还有" & Round((ProgressDuring - (GetTickCount - LastProgress)) / 1000) & "秒", argb(128, 255, 255, 255), StringAlignmentFar
        Else
        DrawTextControl LevelText, WorldName(NowWorld) & " - 第" & NowLevel(NowWorld) + 1 & "天  魔兽正在到来...", argb(128, 255, 255, 255), StringAlignmentFar
        End If
    
    Else
    DrawTextControl LevelText, WorldName(NowWorld) & " - 第" & NowLevel(NowWorld) + 1 & "天  结束啦，可是还有" & UBound(EMonster) & "只魔兽没有被打败", argb(128, 255, 255, 255), StringAlignmentFar
    End If
    If BigProgress <> 0 Then
    For i = 1 To Int(MaxProgress / BigProgress)
    GamePictures(GetPic("flag")).NextFrame.Present Me.HDC, LevelFrame.Left + i * BigProgress / MaxProgress * LevelProgress.Width, LevelFrame.top
    Next
    If NowProgress Mod BigProgress = 0 And GetTickCount - LastProgress <= 5000 And NowProgress <> 0 Then
    GamePictures(GetPic("bigfight")).NextFrame.Present Me.HDC, Me.ScaleWidth / 2 - 215, Me.ScaleHeight / 2 - 40
    End If
    End If
    DrawTextControl InfoText, "MonsterCount: " & UBound(EMonster) & " , FireCount:" & UBound(Fires) & ", ErrorCount:" & ErrCount, argb(128, 255, 255, 255), StringAlignmentNear
    DrawTextRect SuperFunnyFrame.Left + SuperFunnyFrame.Width - 15, SuperFunnyFrame.top + SuperFunnyFrame.Height - 15, SuperCounts, argb(128, 255, 255, 255), StringAlignmentNear

If FirstDanger = True And (GetTickCount - FirstDangerTime) Mod 500 < 300 And Sets(6) = False Then
GamePictures(GetPic("redeffect")).NextFrame.PresentWithClip Me.HDC, 0, 0, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 50
End If


If BigSpeak.Visible = True Then
GamePictures(GetPic("blackeffect")).NextFrame.PresentWithClip Me.HDC, BigSpeak.Left, BigSpeak.top, 0, 0, BigSpeak.Width, BigSpeak.Height
GamePictures(GetPic("blackeffect")).NextFrame.PresentWithClip Me.HDC, BigSpeak.Left, BigSpeak.top, 0, 0, BigSpeak.Width, BigSpeak.Height
GamePictures(GetPic("blackeffect")).NextFrame.PresentWithClip Me.HDC, BigSpeak.Left, BigSpeak.top, 0, 0, BigSpeak.Width, BigSpeak.Height
GamePictures(GetPic(FaceName & "face")).NextFrame.Present Me.HDC, BigSpeak.Left, BigSpeak.top
DrawTextRect BigSpeak.Left + 218, BigSpeak.top + 10, FaceName, argb(255, 0, 176, 240), StringAlignmentNear
LastPenColor = argb(200, 255, 255, 255)
GdipSetSolidFillColor Brush1, argb(200, 255, 255, 255)
GdipDrawString UI, StrPtr(Speach), -1, curFont, NewRectF(BigSpeak.Left + 218, BigSpeak.top + 35, BigSpeak.Width - BigSpeak.Left - 218, BigSpeak.Height - 70), strformat, Brush1
DrawTextRect BigSpeak.Left + BigSpeak.Width - 80, BigSpeak.top + BigSpeak.Height - 35, "单击继续", argb(255, 0, 176, 240), StringAlignmentNear
End If

If pauseframe.Visible = True Then
LastProgress = GetTickCount - PauseProgress
End If

GetCursorPos p
p.X = p.X - Dad.Left / 15: p.Y = p.Y - Dad.top / 15 - 16
'==================================================================================
' UI ++
On Error Resume Next
For Each hey In Me.Controls
Err.Clear
If hey.name <> Me.name Then
    If p.X >= hey.Left And p.X <= hey.Left + hey.Width And p.Y >= hey.top And p.Y <= hey.top + hey.Height And hey.Visible = True And hey.name <> "Speaking" Then
        If Err.Number = 0 Then
        If lastobj <> hey.name Then lastobj = hey.name
        GamePictures(GetPic("clickframe")).NextFrame.PresentWithClip Me.HDC, hey.Left, hey.top, 0, 0, hey.Width, hey.Height, 20
        DrawTextRect hey.Left + hey.Width / 2, hey.top + hey.Height + 2, hey.Tag, argb(185, 255, 255, 255), StringAlignmentCenter
        Exit For
        End If
    End If
End If
Next
If hey.name <> lastobj Then lastobj = ""
'==================================================================================

If Sets(7) = False Then Call DrawEffect(UI, Me.name)

If MusicMode = True Then
Dim Fuck() As Single, MusicHot As Single
Fuck = BGM.GetMusicBar
For i = 0 To 46
GamePictures(GetPic("blueeffect")).NextFrame.PresentWithClip Me.HDC, i * 20, Me.ScaleHeight - Fuck(i) / 7, 0, 0, 15, Fuck(i) / 7
MusicHot = MusicHot + Fuck(i)
Next
MusicHot = MusicHot / 2500
If UBound(FMonster) > 0 Then
    For i = 1 To UBound(FMonster)
    If i > UBound(FMonster) Then Exit For
    FMonster(i).Speed = FMonster(i).OranSpeed * MusicHot
    FMonster(i).During = FMonster(i).OranDuring / MusicHot
    Next
End If
If UBound(EMonster) > 0 Then
    For i = 1 To UBound(EMonster)
    If i > UBound(EMonster) Then Exit For
    EMonster(i).Speed = EMonster(i).OranSpeed * MusicHot
    Next
End If
End If

FPS = FPS + 1
Me.Refresh

GdipDeleteGraphics UI
ShowErr:
If Err.Number <> 0 Then Call ShowError("DrawTimer_Timer")
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 4 And Sets(3) = True Then
If VBA.Environ("Error404_Key") = "1FCB6793" Then testWindow.Show
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    OnLine.close
    OnLineMode = 0
    MainBGM = False
    DrawTimer.Enabled = False
    GdipDeleteGraphics UI
End Sub
Sub NewMsg(msg As String)
Speaking.Tag = GetTickCount
Speaking.ToolTipText = msg
Speaking.Visible = True
End Sub
Sub ShowError(SubName As String)
If Sets(4) = False Or Sets(3) = True Then Exit Sub
GdipCreateFromHDC Me.HDC, UI
GamePictures(GetPic("blackeffect")).NextFrame.PresentWithClip Me.HDC, 0, 0, 0, 0, Me.ScaleWidth, Me.ScaleHeight
DrawTextRect Me.ScaleWidth / 2, Me.ScaleHeight / 2, "An error occured in FightWindow:" & SubName, argb(255, 0, 176, 240), StringAlignmentCenter, True
DrawTextRect Me.ScaleWidth / 2, Me.ScaleHeight / 2 + 50, Err.Description, argb(180, 255, 255, 255), StringAlignmentCenter, False
Ring "error"
Call PauseButton_Click
DrawTimer.Enabled = False
FPSPrinter.Enabled = False
Dad.Caption = "魔兽混战2 - 出错 :("
Me.Refresh
GdipDeleteGraphics UI
End Sub
Private Sub FPSPrinter_Timer()
If Sets(5) = True And RFPS <= 24 Then
    DrawTimer.Interval = 20 - (32 - RFPS) / 20 * 10
    ElseIf Sets(5) = True And RFPS >= 35 Then
    DrawTimer.Interval = 20 - (32 - RFPS) / 20 * 10
End If
End Sub

Sub NewFire(MonsterIndex As Integer, Optional X As Single = -1, Optional Y As Single = -1)
If Sets(3) = False Then
    If Sets(4) = False Then
    On Error Resume Next
    Else
    On Error GoTo ShowErr
    End If
End If
If MonsterIndex > UBound(FMonster) Then Exit Sub
ReDim Preserve Fires(UBound(Fires) + 1)
Fires(UBound(Fires)).Speed = FMonster(MonsterIndex).Speed
Fires(UBound(Fires)).Attack = FMonster(MonsterIndex).Attack
Fires(UBound(Fires)).MonsterType = FMonster(MonsterIndex).MonsterType
Fires(UBound(Fires)).MonsterName = FMonster(MonsterIndex).MonsterName
If FMonster(MonsterIndex).SuperMode = True Then Fires(UBound(Fires)).Attack = Fires(UBound(Fires)).Attack * 3
If X = -1 Then
Fires(UBound(Fires)).X = FMonster(MonsterIndex).X
Else
Fires(UBound(Fires)).X = X
End If
If Y = -1 Then
Fires(UBound(Fires)).Y = FMonster(MonsterIndex).Y
Else
Fires(UBound(Fires)).Y = Y
End If

Fires(UBound(Fires)).MonsterIndex = MonsterIndex

ShowErr:
If Err.Number <> 0 Then Call ShowError("NewFire")
End Sub

Private Sub Friends_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Sets(3) = False Then
    If Sets(4) = False Then
    On Error Resume Next
    Else
    On Error GoTo ShowErr
    End If
End If

WaitIndex = Index
tempMonster = GetDoubleMonsterIndex(0, Friends(Index).Left, Friends(Index).top)
If tempMonster <> 0 Then FMonster(tempMonster).UpdateMove

ShowErr:
If Err.Number <> 0 Then Call ShowError("Friends_MouseMove")
End Sub
Sub SuperMon(ByVal Index As Integer)
If FMonster(Index).During < 900000 Then
NewEffect Me.HDC, Me.name, FMonster(Index).X, FMonster(Index).Y, TaketurnsPic, "Super", 1
Ring "super"
FMonster(Index).SuperMode = True
End If
End Sub
Sub SuperMon2(ByVal Index As Integer)
If OnLineMode <> 0 Then PostData "su " & Index
If FMonster(Index).During < 900000 Then
WaitSuper = False
NewEffect Me.HDC, Me.name, FMonster(Index).X, FMonster(Index).Y, TaketurnsPic, "Super", 1
Ring "super"
FMonster(Index).SuperMode = True
End If
End Sub

Private Sub Friends_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Sets(3) = False Then
    If Sets(4) = False Then
    On Error Resume Next
    Else
    On Error GoTo ShowErr
    End If
End If

Dim NowSuperMonster As Integer, MTY As Monster2_Types, NewType As New FriendlyMonster


NowSuperMonster = GetDoubleMonsterIndex(0, Friends(Index).Left, Friends(Index).top)
MTY = FMonster(NowSuperMonster).MonsterType

If DelMode = True And NowSuperMonster <> 0 Then
If OnLineMode <> 0 Then PostData "dm " & NowSuperMonster
MakeFMonsterDead NowSuperMonster
End If
 
If NowSuperMonster <> 0 And WaitSuper = True Then
If FMonster(NowSuperMonster).SuperMode = True Then Ring "warning": NewMsg "这个魔兽已经有滑稽能量了哦！": Exit Sub
    If MTY = BOOM Or MTY = boom2 Or MTY = boom3 Or MTY = cloud Or MTY = PurpleFunny Or MTY = BurnSuper Or _
        MTY = DarkNess Or MTY = GoBack Or MTY = FireAll Or MTY = IceAll Then
        Ring "warning"
        NewMsg "这个魔兽不喜欢滑稽能量..."
        Exit Sub
    End If
SuperMon2 NowSuperMonster
Exit Sub
End If

If WaitMonster <> -1 Then
    NewType.LoadMonster MCards(WaitMonster).name, 0
    If NewType.MonsterType = Water Then
    If NowSuperMonster <> 0 And FMonster(NowSuperMonster).NeedWater = True Then
        FMonster(NowSuperMonster).NeedWater = False
        FMonster(NowSuperMonster).Speed = FMonster(NowSuperMonster).Speed * 2
        FMonster(NowSuperMonster).During = FMonster(NowSuperMonster).During / 2
        FunnyCounts = FunnyCounts - MCards(WaitMonster).Spend
        MCards(WaitMonster).NowCD = 0
        WaitMonster = -1
        Ring "put"
        NewEffect Me.HDC, Me.name, Friends(Index).Left, Friends(Index).top, TaketurnsPic, "Put", 1
        Exit Sub
    Else
        Ring "warning"
        NewMsg "不能浪费水哇..."
        Exit Sub
    End If
    End If
    
    If MyLevel.CanNew(Friends(Index).Left, Friends(Index).top, NewType.MonsterType) = True Then
    NewFMonster Friends(Index).Left, Friends(Index).top, MCards(WaitMonster).name, MCards(WaitMonster).Level
    FunnyCounts = FunnyCounts - MCards(WaitMonster).Spend
    MCards(WaitMonster).NowCD = 0
    WaitMonster = -1
    Ring "put"
    NewEffect Me.HDC, Me.name, Friends(Index).Left, Friends(Index).top, TaketurnsPic, "Put", 1
    Exit Sub
    Else
    Ring "warning"
    NewMsg MyLevel.NewMsg
    Exit Sub
    End If
End If

If NowSuperMonster <= UBound(FMonster) Then FMonster(NowSuperMonster).UpdateClick

ShowErr:
If Err.Number <> 0 Then Call ShowError("Friends_MouseUp")
End Sub

Private Sub LostTimer_Timer()

End Sub

Private Sub FunnyTimer_Timer()
If Sets(3) = False Then
    If Sets(4) = False Then
    On Error Resume Next
    Else
    On Error GoTo ShowErr
    End If
End If
ReDim Preserve Funnys(UBound(Funnys) + 1)
Funnys(UBound(Funnys)).Y = -30
Funnys(UBound(Funnys)).X = Cards(0).Left + Cards(0).Width + Int(Rnd * (Me.ScaleWidth - Cards(0).Left - Cards(0).Width - 73))
ShowErr:
If Err.Number <> 0 Then Call ShowError("FunnyTimer_Timer")
End Sub

Private Sub hpview_Click()
IsHPView = Not IsHPView
End Sub

Sub NewFaceMsg(Face As String, msg As String)
FaceName = Face
Speach = msg
BigSpeak.Visible = True
MoveTimer.Enabled = False
EffectTimer.Enabled = False
FunnyTimer.Enabled = False
FireTimer.Enabled = False

Do While BigSpeak.Visible = True
Sleep 10: DoEvents
Loop

MoveTimer.Enabled = True
EffectTimer.Enabled = True
FunnyTimer.Enabled = True
FireTimer.Enabled = True
End Sub
Sub WinGame()
If CheatTime <> 0 Then Exit Sub
GdipCreateFromHDC Me.HDC, UI

Dim temp() As String, i As Integer, PresentString As String, Present2 As String, LastBitmap As New BitmapBuffer
Dim LastTop As Single
LastBitmap.Create Me.HDC, Me.ScaleWidth, Me.ScaleHeight
MoveTimer.Enabled = False
FireTimer.Enabled = False
EffectTimer.Enabled = False
FunnyTimer.Enabled = False
DrawTimer.Enabled = False
BitBlt LastBitmap.CompatibleDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, vbSrcCopy
If NowLevel(NowWorld) = MyLevel.Level And MyLevel.Level <> Val(ReadINI("Main", "Max", App.Path & "\level\world" & NowWorld & ".ini")) Then NowLevel(NowWorld) = NowLevel(NowWorld) + 1

OnLineMode = 0

For i = 1 To 25
LastBitmap.Present Me.HDC, 0, 0
GamePictures(GetPic("Back" & NowWorld + 1 & "blur")).NextFrame.Present Me.HDC, 0, 0, i * 10
Me.Refresh
Wait 20
Next

Ring "win"
For i = 1 To 25
GamePictures(GetPic("Back" & NowWorld + 1 & "blur")).NextFrame.Present Me.HDC, 0, 0
GamePictures(GetPic("wintext")).NextFrame.Present Me.HDC, Me.ScaleWidth / 2 - 400 / 2, 50, i / 25 * 255
GamePictures(GetPic("winframe")).NextFrame.PresentWithClip Me.HDC, Me.ScaleWidth / 2 - i / 25 * 358 / 2, 200, 0, 0, i / 25 * 385, 99, i / 25 * 255
Me.Refresh
Wait 20
Next
GamePictures(GetPic("winframe")).NextFrame.PresentWithClip Me.HDC, Me.ScaleWidth / 2 - 358 / 2, 200, 0, 0, 385, 99

If DebugMode = False Then
temp = Split(ProgressPresent, ";")

For i = 0 To UBound(temp)
PresentString = GetPresent(temp(i))
Present2 = PresentIcon(temp(i))
    BitBlt LastBitmap.CompatibleDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, vbSrcCopy
    Ring "get"
    For s = 1 To 25
    LastBitmap.Present Me.HDC, 0, 0
    'GamePictures(GetPic("winframe")).NextFrame.PresentWithClip Me.hDC, Me.ScaleWidth / 2 - 358 / 2, 200, 0, 0, 385, 99
    'GamePictures(GetPic("wintext")).NextFrame.Present Me.hDC, Me.ScaleWidth / 2 - 400 / 2, 50
    GamePictures(GetPic(Present2)).NextFrame.PresentWithClip Me.HDC, Me.ScaleWidth / 2 - 385 / 2, 200 + 45 * i + (i + 1) * 10 + 20, 0, 0, 45, 45, s / 25 * 255
    DrawTextRect Me.ScaleWidth / 2 - 358 / 2 + 50, 200 + 45 * i + (i + 1) * 10 + 33, PresentString, argb(s / 25 * 255, 255, 255, 255), StringAlignmentNear
    LastTop = i
    Me.Refresh
    Wait 10
    Next
Next

If UBound(Presents) > 0 Then
For i = 1 To UBound(Presents)
PresentString = GetPresent(Presents(i).Code)
    BitBlt LastBitmap.CompatibleDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, vbSrcCopy
    Ring "get"
    For s = 1 To 25
    LastBitmap.Present Me.HDC, 0, 0
    'GamePictures(GetPic("winframe")).NextFrame.PresentWithClip Me.hDC, Me.ScaleWidth / 2 - 358 / 2, 200, 0, 0, 385, 99
    'GamePictures(GetPic("wintext")).NextFrame.Present Me.hDC, Me.ScaleWidth / 2 - 400 / 2, 50
    GamePictures(GetPic(Presents(i).Icon)).NextFrame.PresentWithClip Me.HDC, Me.ScaleWidth / 2 - 385 / 2, 200 + 45 * (i + LastTop) + (i + 1 + LastTop) * 10 + 20, 0, 0, 45, 45, s / 25 * 255
    DrawTextRect Me.ScaleWidth / 2 - 358 / 2 + 50, 200 + 45 * (i + LastTop) + (i + 1 + LastTop) * 10 + 33, PresentString, argb(s / 25 * 255, 255, 255, 255), StringAlignmentNear
    Me.Refresh
    Wait 10
    Next
Next
End If

End If

    GamePictures(GetPic("nextbutton")).NextFrame.Present Me.HDC, Me.ScaleWidth - 80, Me.ScaleHeight - 80
    Me.Refresh

Call WriteSave
winFrame.ZOrder
winFrame.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
winFrame.Visible = True

GdipDeleteGraphics UI

End Sub
Private Sub MonsterBorn()
If Sets(3) = False Then
    If Sets(4) = False Then
    On Error Resume Next
    Else
    On Error GoTo ShowErr
    End If
End If
Dim CanBorn As Boolean
If NowProgress = 0 Then
If GetTickCount - LastProgress >= FirstDuring Then CanBorn = True
Else
If GetTickCount - LastProgress >= ProgressDuring Then CanBorn = True
End If

If NowProgress >= MaxProgress + 1 And UBound(EMonster) = 0 Then WinGame

If NowProgress <= MaxProgress And CanBorn = True Then NowProgress = NowProgress + 1
MyLevel.Update CanBorn

If CanBorn = True And NowProgress <= MaxProgress Then
LastProgress = GetTickCount
If NowProgress = MaxProgress + 1 Then Exit Sub
temp = Split(ProgressMonsters(Int(NowProgress / MaxProgress * 3)), ";")
Randomize
For i = 1 To ProgressMonster
NewEMonster temp(Int(Rnd * UBound(temp))), Me.ScaleWidth + 36, Friends(Int(Rnd * 5) * 9).top
Next


If BigProgress <> 0 Then
If NowProgress Mod BigProgress = 0 Then
Ring "big"
If BGM.Rate = 44100 Then BGM.SetPlayRate 1.5
For i = 1 To ProgressMonster * 5
NewEMonster temp(Int(Rnd * UBound(temp))), Me.ScaleWidth + 36, Friends(Int(Rnd * 5) * 9).top
Next
Else
If BGM.Rate = 66150 Then BGM.SetPlayRate 1
End If
End If

End If

    If UBound(EMonster) = 0 Then
    quickbutton.Visible = True
    Else
    quickbutton.Visible = False
    End If

ShowErr:
If Err.Number <> 0 Then Call ShowError("MonsterBorn_Timer")
End Sub

Private Sub MoveTimer_Timer()
If Sets(3) = False Then
    If Sets(4) = False Then
    On Error Resume Next
    Else
    On Error GoTo ShowErr
    End If
End If
Dim RunOnce As Boolean
restart:
Dim i As Integer, s As Integer, temp2 As Boolean

If UBound(EMonster) > 0 Then
For i = 1 To UBound(EMonster)
EMonster(i).UpdateMove
yee = (EMonster(i).Y - Friends(0).top) / 76
If EMonster(i).X <= Friends(0).Left And Dogs(yee).Visible = True Then
NewFire 0, Dogs(yee).Left + 36, Dogs(yee).top
Dogs(yee).Visible = False
End If
If EMonster(i).X < 0 Then 'GameOver
Call GameOver
Exit Sub
End If
If EMonster(i).X < Friends(0).Left + 73 * 2 Then
temp2 = True
If FirstDanger = False Then FirstDangerTime = GetTickCount: Ring "aoh"
End If
Next
End If

FirstDanger = temp2

For i = 0 To 4
FMonsterOnLine(i) = ""
EMonsterOnLine(i) = ""
Next

If UBound(FMonster) > 0 Then
For i = 1 To UBound(FMonster)
'If FMonster(i).MonsterName <> "" Then
yee = (FMonster(i).Y - Friends(0).top) / 76
FMonsterOnLine(yee) = FMonsterOnLine(yee) & i & ";"
'End If
Next
End If

If UBound(EMonster) > 0 Then
    For i = 1 To UBound(EMonster)
    yee = (EMonster(i).Y - Friends(0).top) / 76
    EMonsterOnLine(yee) = EMonsterOnLine(yee) & i & ";"
    Next
End If

If UBound(Fires) > 0 Then
For i = 1 To UBound(Fires)
If i > UBound(Fires) Then Exit For
temp = Split(EMonsterOnLine((Fires(i).Y - Friends(0).top) / 76), ";")
For s = 0 To UBound(temp)
    If Val(temp(s)) <> 0 And temp(s) <= UBound(EMonster) Then
    If Abs(Fires(i).X - EMonster(temp(s)).X) < 60 Then '进球了
        Fires(i).UpdateFire (Val(temp(s)))
            If Fires(i).MonsterType <> OhAngle And Fires(i).MonsterType <> boom2 And Fires(i).MonsterType <> ghost And Fires(i).MonsterType <> FireGhost And Fires(i).MonsterType <> ThreadGhost Then
                Set Fires(i) = Fires(UBound(Fires))
                ReDim Preserve Fires(UBound(Fires) - 1)
                i = i - 1
              GoTo NextNextOne
            End If
    End If
    End If
Next
NextNextOne:
Next
End If

If BGM.Rate = 88200 And RunOnce = False Then RunOnce = True: GoTo restart

ShowErr:
If Err.Number <> 0 Then Call ShowError("MoveTimer_Timer")
End Sub

Private Sub musicbutton_Click()
MusicMode = Not MusicMode
If MusicMode = False Then
If UBound(FMonster) > 0 Then
    For i = 1 To UBound(FMonster)
    If i > UBound(FMonster) Then Exit For
    FMonster(i).Speed = FMonster(i).OranSpeed
    FMonster(i).During = FMonster(i).OranDuring
    Next
End If
If UBound(EMonster) > 0 Then
    For i = 1 To UBound(EMonster)
    If i > UBound(EMonster) Then Exit For
    EMonster(i).Speed = EMonster(i).OranSpeed
    Next
End If
End If
End Sub

Private Sub PauseButton_Click()
If OnLineMode <> 0 Then Ring "warning": NewMsg "多人游戏不能暂停哟。": Exit Sub
PauseProgress = GetTickCount - LastProgress
MoveTimer.Enabled = False
FireTimer.Enabled = False
EffectTimer.Enabled = False
FunnyTimer.Enabled = False

'pauseframe.Visible = True
'pauseframe.ZOrder
'pauseframe.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
ShowToolWindow PauseWindow, False

End Sub

Public Sub pauseframe_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    PauseProgress = GetTickCount - LastProgress
    For i = 0 To UBound(FMonster)
    FMonster(i).LastFireTime = GetTickCount - 1500
    Next
    MoveTimer.Enabled = True
    FireTimer.Enabled = True
    EffectTimer.Enabled = True
    FunnyTimer.Enabled = True
    pauseframe.Visible = False
ElseIf Button = 2 Then
    Me.Hide
    CreateAChild MainWindow
    Unload Me
Else

    Text1.Visible = False
    BigSpeak.ZOrder
    Set MyLevel = Nothing
    MyLevel.Level = NowLevel(NowWorld)
    MyLevel.World = NowWorld
    NowProgress = 0
    MusicMode = False
    BGM.SetPlayRate 1
    MaxProgress = Val(ReadINI("Level" & NowLevel(NowWorld), "AttackCount", App.Path & "\level\world" & NowWorld & ".ini"))
    BigProgress = Val(ReadINI("Level" & NowLevel(NowWorld), "AttackExEvery", App.Path & "\level\world" & NowWorld & ".ini"))
    ProgressMonster = Val(ReadINI("Level" & NowLevel(NowWorld), "MonsterCount", App.Path & "\level\world" & NowWorld & ".ini"))
    ProgressDuring = Val(ReadINI("Level" & NowLevel(NowWorld), "AttackDuring", App.Path & "\level\world" & NowWorld & ".ini"))
    FirstDuring = Val(ReadINI("Level" & NowLevel(NowWorld), "FirstAttack", App.Path & "\level\world" & NowWorld & ".ini"))
    ProgressPresent = ReadINI("Level" & NowLevel(NowWorld), "Present", App.Path & "\level\world" & NowWorld & ".ini")
    For i = 0 To 3
    ProgressMonsters(i) = ReadINI("Level" & NowLevel(NowWorld), "Monster" & i + 1, App.Path & "\level\world" & NowWorld & ".ini")
    Next
    LastProgress = GetTickCount
    '===========================BaBaBaBaBa前方无异常===========================
    ReDim Fires(0): ReDim EMonster(0): ReDim Funnys(0): ReDim Presents(0): FunnyCounts = AtFirstFunny: WaitMonster = -1: SuperCounts = 1
    ReDim FMonster(0)
    RfWidth = Me.ScaleWidth: RfHeight = Me.ScaleHeight
    DelMode = False: IsHPView = False: WaitSuper = False
    FirstDanger = False: FirstDangerTime = 0: MoveTimer.Enabled = True: EffectTimer.Enabled = True: FunnyTimer.Enabled = True: FireTimer.Enabled = True
    '====================================================================
    If DebugMode = True Then FunnyCounts = 9999: SuperCounts = 9999
    MainBGM = True
    For i = 0 To Dogs.UBound
    Dogs(i).Visible = True
    Next
    FMonster(0).MonsterName = "Angle": FMonster(0).Speed = 10: FMonster(0).MonsterType = OhAngle '设置我们的小天使【噗】
    FPSPrinter.Enabled = True

End If
End Sub

Private Sub quickbutton_Click()
quickbutton.Visible = False
Ring "big"
If NowProgress > 0 Then
LastProgress = LastProgress - ProgressDuring
Else
LastProgress = LastProgress - FirstDuring
End If
End Sub

Private Sub ratebutton_Click()
If BGM.Rate = 88200 Then
BGM.SetPlayRate 1
FirstDuring = FirstDuring * 2
ProgressDuring = ProgressDuring * 2
Else
BGM.SetPlayRate 2
FirstDuring = FirstDuring / 2
ProgressDuring = ProgressDuring / 2
End If
End Sub

Private Sub Speaking_Click()
Speaking.Tag = GetTickCount - 10000
End Sub
Sub GameOver()
If CheatTime <> 0 Then Exit Sub
If Sets(3) = False Then
    If Sets(4) = False Then
    On Error Resume Next
    Else
    On Error GoTo ShowErr
    End If
End If
    OnLineMode = 0
    BGM.Pause
    Ring "lost"
    DrawTimer.Enabled = False
    GdipCreateFromHDC Me.HDC, UI
    EffectTimer.Enabled = False
    For s = 1 To 20
    DrawRectangleRect 0, 0, Me.ScaleWidth, Me.ScaleHeight, argb(25, 0, 0, 0)
    Sleep 50: DoEvents
    Me.Refresh
    Next
    GamePictures(GetPic("lost")).NextFrame.Present Me.HDC, Me.ScaleWidth / 2 - 516 / 2 - 50, Me.ScaleHeight / 2 - 250 / 2 + 20
    Me.Refresh
    For s = 1 To 100
    Sleep 50: DoEvents
    Next
    MoveTimer.Enabled = False
    Me.Hide
    CreateAChild MainWindow
    OnLine.close
    Unload Me
ShowErr:
If Err.Number <> 0 Then Call ShowError("GameOver")
End Sub
Sub DrawPicNameControl(Control As Object, PicName As String)
If Control.Visible = False Then Exit Sub
GamePictures(GetPic(PicName)).NextFrame.Present Me.HDC, Control.Left, Control.top
End Sub
Sub DrawPicNameMonster(Monster As FriendlyMonster, PicName As String)
On Error Resume Next
GamePictures(GetPic(PicName)).NextFrame.Present Me.HDC, Monster.X, Monster.Y
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

Private Sub SuperFunnyFrame_Click()
If Sets(3) = False Then On Error GoTo ShowErr
If DelMode = True Or WaitMonster <> -1 Then Exit Sub
If WaitSuper = True Then WaitSuper = False: SuperCounts = SuperCounts + 1: Exit Sub
If SuperCounts > 0 Then
SuperCounts = SuperCounts - 1
WaitSuper = True
Else
Ring "warning"
NewMsg ReadINI("Game", "NO_SUPER", App.Path & "\monster\game.ini")
End If
ShowErr:
If Err.Number <> 0 Then Call ShowError("SuperFunnyFrame_Click")
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Ring "warning"
NewMsg PlayerName & ":" & Replace(Text1.Text, " ", "_")
PostData "t " & Replace(Text1.Text, " ", "_")
Text1.Text = ""
End If
End Sub

Private Sub winFrame_Click()
On Error Resume Next
If NowProgress > MaxProgress Then
DrawTimer.Enabled = False
Me.Hide
CreateAChild MainWindow
OnLine.close
Unload Me
End If
End Sub
