VERSION 5.00
Begin VB.Form SetWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "SetWindow"
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox setbuttons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "显示FPS"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   8
      Left            =   120
      TabIndex        =   11
      Top             =   3450
      Width           =   6015
   End
   Begin VB.CheckBox setbuttons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "禁用特效文本（提高性能，降低体验）"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   7
      Left            =   120
      TabIndex        =   10
      Top             =   3060
      Width           =   6015
   End
   Begin VB.CheckBox setbuttons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "禁用魔兽警戒闪烁提示"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   2670
      Width           =   6015
   End
   Begin VB.CheckBox setbuttons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "当我的帧数太低时耗费CPU占用以提高帧数"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   6015
   End
   Begin VB.CheckBox setbuttons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "取消忽略错误"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1890
      Width           =   2235
   End
   Begin VB.CheckBox setbuttons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "开发者模式"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   4020
      Width           =   2235
   End
   Begin VB.CheckBox setbuttons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "柔和式"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1470
      Width           =   2235
   End
   Begin VB.CheckBox setbuttons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "关闭背景音乐"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2235
   End
   Begin VB.CheckBox setbuttons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "关闭音效"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   690
      Width           =   2235
   End
   Begin 魔兽混战2.UIButton UIButton2 
      Height          =   315
      Left            =   5040
      TabIndex        =   6
      Top             =   4500
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BackColor       =   15773696
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   "确定"
      BackColor2      =   15785984
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设置"
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
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F0B000&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6165
   End
End
Attribute VB_Name = "SetWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dad.SetFocus
For i = 0 To UBound(Sets)
If Sets(i) = True Then
setbuttons(i).value = 1
Else
setbuttons(i).value = 0
End If
Next
End Sub

Private Sub setbuttons_Click(Index As Integer)
If setbuttons(Index).value = 1 Then
Sets(Index) = True
Else
Sets(Index) = False
End If
SaveSet
End Sub

Private Sub UIButton2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Transitions AeroWindow.ToolFrame, True, upstairs, 5
Unload Me
Unload AeroWindow
End Sub
