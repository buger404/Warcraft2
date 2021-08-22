VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form testWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "测试窗口"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6945
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "一键测试关卡"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   4
      Top             =   4680
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4485
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   7911
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"testWindow.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "释放所有资源"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1860
      TabIndex        =   2
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重新加载资源"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3570
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "检查资源缺漏"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5280
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
End
Attribute VB_Name = "testWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo sth
Dim ErrorCount As Long
AddLog "=========================="
AddLog "开始检查资源文件", 2
AddLog "--检查友好类型魔兽--", 3
a = Dir(App.Path & "\monster\fmonster\")
Do While a <> ""
b = Split(a, ".")
c = c & b(0) & vbCrLf
a = Dir()
DoEvents
Loop
d = Split(c, vbCrLf)
AddLog "*发现友好魔兽数量：" & UBound(d), 2
For i = 0 To UBound(d) - 1
b(0) = d(i)
If Dir(App.Path & "\assets\" & b(0) & "0.*") = "" Then AddLog "缺少：" & b(0) & "0.png", 1
If GetPic(b(0) & "0") = 0 Then AddLog b(0) & "0.png" & "没有被加载，需要重新加载资源文件。", 3
If Dir(App.Path & "\assets\" & b(0) & "1.*") = "" Then AddLog "缺少：" & b(0) & "1.png", 1
If GetPic(b(0) & "1") = 0 Then AddLog b(0) & "1.png" & "没有被加载，需要重新加载资源文件。", 3
If Dir(App.Path & "\assets\" & b(0) & "Attack.*") = "" Then AddLog "缺少：" & b(0) & "Attack.png", 1
If GetPic(b(0) & "Attack") = 0 Then AddLog b(0) & "Attack.png" & "没有被加载，需要重新加载资源文件。", 3
If Dir(App.Path & "\assets\" & b(0) & "Icon.*") = "" Then AddLog "缺少：" & b(0) & "Icon.png", 1
If GetPic(b(0) & "Icon") = 0 Then AddLog b(0) & "Icon.png" & "没有被加载，需要重新加载资源文件。", 3
If Dir(App.Path & "\assets\" & b(0) & "Fire.*") = "" Then AddLog "缺少：" & b(0) & "Fire.png", 1
If GetPic(b(0) & "Fire") = 0 Then AddLog b(0) & "Fire.png" & "没有被加载，需要重新加载资源文件。", 3
If Dir(App.Path & "\sounds\" & b(0) & ".wav") = "" Then AddLog "缺少：" & b(0) & ".wav", 1
Ring b(0)
Next
AddLog "--友好类型魔兽资源检查完成--", 3
AddLog "--检查敌对类型魔兽--", 3
a = Dir(App.Path & "\monster\emonster\")
c = ""
Do While a <> ""
b = Split(a, ".")
c = c & b(0) & vbCrLf
Ring b(0)
a = Dir()
DoEvents
Loop
d = Split(c, vbCrLf)
AddLog "发现敌对魔兽数量：" & UBound(d), 2
For i = 0 To UBound(d) - 1
b(0) = d(i)
If Dir(App.Path & "\assets\" & b(0) & "0.*") = "" Then AddLog "缺少：" & b(0) & "0.png", 1
If GetPic(b(0) & "0") = 0 Then AddLog b(0) & "0.png" & "没有被加载，需要重新加载资源文件。", 3
If Dir(App.Path & "\assets\" & b(0) & "1.*") = "" Then AddLog "缺少：" & b(0) & "1.png", 1
If GetPic(b(0) & "1") = 0 Then AddLog b(0) & "1.png" & "没有被加载，需要重新加载资源文件。", 3
If Dir(App.Path & "\assets\" & b(0) & "Attack0.*") = "" Then AddLog "缺少：" & b(0) & "Attack0.png", 1
If GetPic(b(0) & "Attack0") = 0 Then AddLog b(0) & "Attack0.png" & "没有被加载，需要重新加载资源文件。", 3
If Dir(App.Path & "\assets\" & b(0) & "Attack1.*") = "" Then AddLog "缺少：" & b(0) & "Attack1.png", 1
If GetPic(b(0) & "Attack1") = 0 Then AddLog b(0) & "Attack1.png" & "没有被加载，需要重新加载资源文件。", 3
If Dir(App.Path & "\sounds\" & b(0) & ".wav") = "" Then AddLog "缺少：" & b(0) & ".wav", 1
Next
AddLog "--敌对类型魔兽资源检查完成--", 3
AddLog "所有检查完成   错误的操作：" & ErrorCount & "个", 2
AddLog "=========================="
sth:
If Err.Number <> 0 Then Err.Clear: ErrorCount = ErrorCount + 1: Resume Next
End Sub

Private Sub Command2_Click()
On Error GoTo sth
Dim ErrorCount As Long
AddLog "=========================="
AddLog "重新加载资源文件", 2
Set ActiveWindow = Dad
Unload MainWindow
AddLog "卸载主窗口", 1
For i = 0 To UBound(GamePictures)
GamePictures(i).NextFrame.Delete
Next
BASS_Free
AddLog "BASS被卸载", 1
ReDim GamePictures(0)
ReDim GameSounds(0)
AddLog "释放所有资源", 3
BASS_Init -1, 44100, BASS_DEVICE_3D, Dad.hWnd, 0
AddLog "BASS被初始化", 2
LoadAllAssets
LoadAllSounds '加载所有资源
AddLog "加载所有资源", 2
CreateAChild MainWindow
AddLog "加载主窗口", 2
AddLog "完成   错误的操作：" & ErrorCount & "个", 2
AddLog "=========================="
sth:
If Err.Number <> 0 Then Err.Clear: ErrorCount = ErrorCount + 1: Resume Next
End Sub
Sub AddLog(msg As String, Optional flag As Integer = 0)
'Text1.Text = Text1.Text & Now & " " & msg & vbCrLf
'Text1.SelLength = Len(msg)
'Text1.SelStart = Len(Text1.Text) - Len(msg)
Text1.SelColor = RGB(90, 185, 60)
Text1.SelText = Now & " "
Select Case flag
Case Is = 0
Text1.SelColor = RGB(36, 36, 36)
Case Is = 1
Text1.SelColor = RGB(255, 30, 0)
Case Is = 2
Text1.SelColor = RGB(0, 176, 240)
Case Is = 3
Text1.SelColor = RGB(128, 128, 128)
End Select
Text1.SelText = msg & vbCrLf
Text1.SelLength = 0
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Command3_Click()
On Error GoTo sth
Dim ErrorCount As Long
AddLog "=========================="
Set ActiveWindow = Dad
AddLog "重新加载资源文件", 2
Unload MainWindow
AddLog "卸载主窗口", 1
For i = 0 To UBound(GamePictures)
GamePictures(i).NextFrame.Delete
Next
BASS_Free
AddLog "BASS被卸载", 1
ReDim GamePictures(0)
ReDim GameSounds(0)
AddLog "释放所有资源", 3
AddLog "完成   错误的操作：" & ErrorCount & "个", 2
AddLog "=========================="
sth:
If Err.Number <> 0 Then Err.Clear: ErrorCount = ErrorCount + 1: Resume Next
End Sub

Private Sub Command4_Click()
On Error GoTo sth
Dim ErrorCount As Long
AddLog "=========================="
AddLog "设置滑稽果数量", 2
FunnyCounts = 99999
AddLog "设置滑稽能量数量", 2
FightWindow.SuperCounts = 99999
AddLog "取消关卡间隔", 2
FightWindow.ProgressDuring = 0
AddLog "取消卡片CD", 2
For i = 0 To UBound(MCards)
MCards(i).CDTime = 0
Next
AddLog "完成   错误的操作：" & ErrorCount & "个", 2
AddLog "=========================="
sth:
If Err.Number <> 0 Then Err.Clear: ErrorCount = ErrorCount + 1: Resume Next
End Sub
