VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Square war II server"
   ClientHeight    =   8100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin �����ħ�޻�ս2.UIButton UIButton3 
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   0
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BackColor       =   16777215
      ForeColor       =   1907997
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   "�ر�"
      BackColor3      =   -2147483633
   End
   Begin VB.Timer DrawTimer 
      Interval        =   1000
      Left            =   5550
      Top             =   150
   End
   Begin VB.PictureBox countframe 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00F0B000&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   531
      TabIndex        =   8
      Top             =   6900
      Width           =   7965
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100 / s"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   750
         Width           =   525
      End
   End
   Begin �����ħ�޻�ս2.UIButton UIButton2 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   556
      BackColor       =   16777215
      ForeColor       =   1907997
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   "��ҳ"
      BackColor3      =   -2147483633
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   5550
      Top             =   750
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin �����ħ�޻�ս2.UIButton UIButton5 
      Height          =   315
      Left            =   840
      TabIndex        =   18
      Top             =   0
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BackColor       =   16777215
      ForeColor       =   1907997
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   "Socks"
      BackColor3      =   -2147483633
   End
   Begin �����ħ�޻�ս2.UIButton UIButton6 
      Height          =   315
      Left            =   1770
      TabIndex        =   21
      Top             =   0
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BackColor       =   16777215
      ForeColor       =   12632064
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   "����"
      BackColor3      =   -2147483633
   End
   Begin VB.Frame MsgFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Form1"
      Height          =   3015
      Left            =   600
      TabIndex        =   2
      Top             =   2250
      Visible         =   0   'False
      Width           =   4815
      Begin �����ħ�޻�ս2.UIButton msg_ok 
         Height          =   315
         Left            =   2700
         TabIndex        =   6
         Top             =   2550
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   15773696
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Caption         =   "Ŷ"
         BackColor2      =   15785984
         BackColor3      =   15773696
      End
      Begin �����ħ�޻�ս2.UIButton msg_no 
         Height          =   315
         Left            =   3750
         TabIndex        =   7
         Top             =   2550
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   4120574
         ForeColor       =   1907997
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Caption         =   "��Ҫ"
         BackColor2      =   4128766
         BackColor3      =   4120574
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00F2F2F2&
         X1              =   150
         X2              =   4650
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label msg_text 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "All connections will be disconnected !"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D1D1D&
         Height          =   1515
         Left            =   150
         TabIndex        =   5
         Top             =   750
         Width           =   4590
      End
      Begin VB.Label msg_title 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         TabIndex        =   3
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F0B000&
         Height          =   615
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   6615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00F0B000&
         Height          =   315
         Left            =   1350
         Top             =   150
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   150
      TabIndex        =   10
      Top             =   450
      Width           =   7365
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2F2F2&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "΢���ź� Light"
            Size            =   9
            Charset         =   0
            Weight          =   290
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Text            =   "command..."
         Top             =   5100
         Width           =   5565
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   3915
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   900
         Width           =   5565
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��־"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F0B000&
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Top             =   600
         Width           =   480
      End
      Begin VB.Label newuser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���û� : "
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   2550
         TabIndex        =   13
         Top             =   300
         Width           =   750
      End
      Begin VB.Label onlineuser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������� : 0"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   0
         TabIndex        =   12
         Top             =   300
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ϣ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F0B000&
         Height          =   315
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6315
      Left            =   150
      TabIndex        =   19
      Top             =   450
      Width           =   5715
      Begin �����ħ�޻�ս2.UIButton UIButton7 
         Height          =   315
         Left            =   2250
         TabIndex        =   22
         Top             =   600
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         BackColor       =   16777215
         ForeColor       =   1907997
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Caption         =   "ˢ��"
         BackColor3      =   -2147483633
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2025
         Left            =   150
         TabIndex        =   20
         Top             =   150
         Width           =   2715
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   6165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgState As Integer
Dim MessageCount As Single, LastCount As Single
Dim MsgSize As Single, LastSize As Single
Dim OnLineCount As Long
Dim NowSock As Integer
Dim SockName(999999) As String
Dim SockWorld(999999) As Integer
Private Sub DrawTimer_Timer()
On Error Resume Next
Dim b As StdPicture
Set b = countframe.Picture
countframe.Cls
Set countframe.Picture = Nothing
countframe.PaintPicture b, -10, 0
countframe.Refresh
countframe.ForeColor = RGB(0, 176, 240)
countframe.Line (countframe.Width / 30 - 10, countframe.Height / 30 - LastCount / 300)-(countframe.Width / 30, countframe.Height / 30 - MessageCount / 300)
countframe.ForeColor = RGB(0, 204, 106)
countframe.Line (countframe.Width / 30 - 10, countframe.Height / 30 - LastSize / 300)-(countframe.Width / 30, countframe.Height / 30 - MsgSize / 300)
countframe.Refresh
Set countframe.Picture = countframe.Image
Label2.Caption = MessageCount & " / s , " & MsgSize & " kb / s ."
LastCount = MessageCount
MessageCount = 0
LastSize = MsgSize
MsgSize = 0


End Sub

Sub AddLog(Msg As String)
Text1.text = Now & " " & Msg & vbCrLf & Text1.text
If Len(Text1.text) > 5000 Then
Open App.Path & "\" & Year(Now) & "." & Month(Now) & "." & Day(Now) & ".." & Hour(Now) & ".." & Minute(Now) & ".." & Second(Now) & ".log" For Output As #1
Print #1, Text1.text
Text1.text = ""
AddLog "��־̫���ˣ��Ѿ��������沢���������Ļ��"
Close #1
End If
End Sub
Private Sub Form_Load()
AddLog "�����������"
Frame1.ZOrder

Load Winsock2(1)
Winsock2(1).LocalPort = 6604
Winsock2(1).Listen
List2.AddItem "Winsock0 State User IP"
List2.AddItem "Winsock1 State User IP"
NowSock = 1

Shape1.Move 0, 0, MsgFrame.Width, MsgFrame.Height
End Sub
Function FreeWinsock()
On Error GoTo sth
For i = 1 To Winsock2.UBound + 1
b = Winsock2(i).Index
Next
sth:
If Err.Number <> 0 Then FreeWinsock = i
End Function
Function Msgbox(text As String, title As String)
MsgState = -1
MsgFrame.Visible = True
MsgFrame.Move Me.Width / 2 - MsgFrame.Width / 2, Me.Height / 2 - MsgFrame.Height / 2
MsgFrame.ZOrder
msg_title.Caption = title
msg_text.Caption = text
Do While MsgState = -1
DoEvents
Sleep 10
Loop
Msgbox = MsgState
End Function

Private Sub Form_Resize()
MsgFrame.Move Me.Width / 2 - MsgFrame.Width / 2, Me.Height / 2 - MsgFrame.Height / 2
countframe.Width = Me.Width - countframe.Left * 2
countframe.Top = Me.Height - 1500
Frame1.Width = Me.Width - Frame1.Left * 2
Frame1.Height = countframe.Top - 300 - Frame1.Top
Text1.Width = Frame1.Width
Text1.Height = Frame1.Height - Text1.Top - Text2.Height
Text2.Top = Text1.Top + Text1.Height
Text2.Width = Frame1.Width
'Frame2.Width = Frame1.Width
'Frame2.Height = Frame1.Height
Frame3.Move Frame3.Left, Frame3.Top, Frame1.Width, Frame1.Height
List2.Move -15, -15, Frame1.Width + 30, Frame1.Height + 300
UIButton7.Move Frame3.Width - UIButton7.Width - 300, Frame3.Height - UIButton7.Height - 300

End Sub



Private Sub List2_Click()
On Error Resume Next
If Msgbox(List2.List(List2.ListIndex) & vbCrLf & "�Ͽ����������", "����") = 1 Then
Winsock2(List2.ListIndex).Close
Winsock2_Close List2.ListIndex
End If
End Sub


Private Sub msg_no_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgState = 2
MsgFrame.Visible = False
End Sub

Private Sub msg_ok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgState = 1
MsgFrame.Visible = False
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
Dim a() As String
a = Split(Text2.text, "&>")
Select Case a(1)
Case "s"
PostData Val(a(2)), Replace(a(0), "&Now&", Now)
AddLog "��Ϣ�ɹ����͵����" & Val(a(2)) & "��"
Case "test"
PostData Val(a(2)), "test " & GetTickCount & " " & Text1.text
AddLog "���ڲ��Կͻ���" & Val(a(2)) & "��Ϣ������������ʱ��..."
End Select
End If
End Sub

Private Sub UIButton1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame2.ZOrder
End Sub

Private Sub UIButton2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1.ZOrder
End Sub

Private Sub UIButton3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Msgbox("�������ӽ��ᶪʧ��", "����") = 1 Then
End
End If
End Sub

Private Sub UIButton4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
For i = 0 To Winsock2.UBound
If Winsock2(i).Tag = List1.ListIndex + 1 Then Winsock2(i).Close: Exit For
Next
End Sub

Private Sub UIButton5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame3.ZOrder
End Sub

Private Sub UIButton6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Select Case UIButton6.Caption
Case Is = "����"
Dim i As Integer
For i = 1 To Winsock2.UBound
Winsock2(i).Close
Winsock2_Close i
Next
List2.Clear
List2.AddItem "a"
AddLog "��������ģʽ���ܾ��������Ӳ��Ͽ��������ӡ�"
UIButton6.Caption = "���"
Case Is = "���"
Load Winsock2(1)
Winsock2(1).LocalPort = 6604
Winsock2(1).Listen
List2.AddItem "Winsock1 State User IP"
NowSock = 1
UIButton6.Caption = "����"
End Select
End Sub

Private Sub UIButton7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
For i = 0 To Winsock2.UBound
List2.List(i) = "#����"
List2.List(i) = "Winsock" & i & " ״̬��" & Winsock2(i).state & " ���ƣ�" & SockName(i) & " ���䣺" & Winsock2(i).Tag
If Err.Number <> 0 Then List2.List(i) = "#����": Err.Clear
Next
End Sub

Private Sub Winsock2_Close(Index As Integer)
            OnLineCount = OnLineCount - 1
            onlineuser.Caption = "��������� : " & OnLineCount
            
            If Winsock2(Index).Tag <> "" Then SendMessageRoom Val(Winsock2(Index).Tag), "t ϵͳ��" & SockName(Index) & "�˳�����Ϸ��"
            AddLog "���" & Index & " �˳�����Ϸ ."

            Winsock2(Index).Tag = ""
            Winsock2(Index).Close
            Do While Winsock2(Index).state <> 0
            DoEvents
            Loop
            
            On Error Resume Next
            If Index <> 0 Then
            List2.List(Index) = "#����"
            Unload Winsock2(Index)
            End If
            If Index = NowSock Then
            Winsock2(Index).Bind 6604
            Winsock2(Index).Listen
            End If
End Sub

Private Sub Winsock2_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Winsock2(Index).Close
Winsock2(Index).Accept requestID
OnLineCount = OnLineCount + 1
MessageCount = MessageCount + 1
Dim Free As Integer
Free = FreeWinsock
Load Winsock2(Free)
If Free > UBound(Socks) Then List2.AddItem "Winsock State User IP"
Winsock2(Free).Bind 6604
Winsock2(Free).Listen
NowSock = Free

AddLog "���" & Index & " ���ӵ��˷�������"

'For i = 0 To Winsock2.UBound
'If Winsock2(i).RemoteHostIP = Winsock2(Index).RemoteHostIP And i <> Index Then
'AddLog " " & Winsock2(Index).RemoteHostIP & " double connect , closed ."
'Winsock2(Index).Close
'Winsock2_Close Index
'End If
'Next
End Sub
Sub PostData(Index As Integer, data As String)
On Error GoTo sth
Winsock2(Index).SendData data & "��"

sth:
If Err.Number <> 0 Then AddLog "���" & Index & "����Ϣ����ʧ�ܣ���Ϣ��" & data

End Sub
Sub SendMessageRoom(ByVal Room As Integer, ByVal data As String)
On Error Resume Next
For Each Sock In Winsock2
    If Val(Sock.Tag) = Val(Winsock2(Room).Tag) And Sock.Index <> Room Then
    PostData Sock.Index, data
    End If
Next
End Sub
Sub DelRoom(ByVal Room As Integer, data As Variant)
On Error Resume Next
For Each Sock In Winsock2
    If Val(Sock.Tag) = Room Then Sock.Tag = ""
Next
End Sub
Sub AddRoom(ByVal Room As Integer)
On Error Resume Next
Dim RoomPlayer As Integer, RoomIndex As Integer
RoomIndex = -1
For Each Sock In Winsock2
If Sock.Index <> Room And SockWorld(Sock.Index) = SockWorld(Room) And Sock.Index <> 0 And Sock.Tag <> "" Then
    If RoomIndex = -1 Then
    RoomIndex = Val(Sock.Tag)
    Else
    RoomIndex = -1
    End If
End If
Next

If RoomIndex = -1 Then
Winsock2(Room).Tag = Room
AddLog "���" & Room & "�����˷���"
Else
Winsock2(Room).Tag = RoomIndex
PostData Room, "s " & SockWorld(RoomIndex)  '��ʼ��Ϸ
PostData RoomIndex, "s2 " & SockWorld(RoomIndex)  '������ʼ��Ϸ
AddLog "���" & Room & "�����˷���" & RoomIndex & "����ʼս����"
End If
End Sub
Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)

'On Error GoTo sth
Dim temp As String, data() As String, Run() As String

Winsock2(Index).GetData temp
data = Split(temp, "��")
For i = 0 To UBound(data)
    If data(i) <> "" Then
    MessageCount = MessageCount + 1
    MsgSize = MsgSize + (Len(data(i)) * 2 + 2) / 1024
    Run = Split(data(i), " ")

    Select Case Run(0) 'ȫ������
        Case "test"
        AddLog "�Կͻ���" & Index & "�Ĳ�����ϣ����ͽ���һ�������̣�" & Len(Run(1)) * 2 * 2 + Len(Text1.text) * 2 & "Byte����ʱ��" & GetTickCount - Val(Run(1)) & "ms"
        Case "c" ' [c world] ��������
        SockWorld(Index) = Val(Run(1))
        AddRoom Index
        Case "t" '[t msg] ����
        SendMessageRoom Index, "t " & SockName(Index) & ":" & Run(1)
        AddLog "���� " & Winsock2(Index).Tag & "��" & SockName(Index) & "˵����" & Run(1) & "��"
        Case "d" '[d]ɾ�����䣨������Ȩ��
        If Val(Winsock2(Index).Tag) = Index Then
        AddLog "���" & Index & "ɾ���˷���" & Winsock2(Index).Tag
        SendMessageRoom Index, "d"
        DelRoom Winsock2(Index).Tag, ""
        End If
        Case "n" '[n name]��������
        SockName(Index) = Run(1)
        AddLog "��ȡ���" & Index & "������Ϊ" & Run(1)
        Case "fm" ' [fm name x y]�����Ѻ�ħ��
        SendMessageRoom Index, data(i)
        'AddLog "���� " & Winsock2(Index).Tag & "��" & "���" & Index & "��" & Run(1) & "������(" & Run(2) & "," & Run(3) & ")"
        Case "em" ' [em name x y]�����ж�ħ�� ��������Ȩ��
        If Val(Winsock2(Index).Tag) = Index Then
        SendMessageRoom Index, data(i)
        'AddLog "���� " & Winsock2(Index).Tag & "��" & Run(1) & "��(" & Run(2) & "," & Run(3) & ")������"
        End If
        Case "su" '[su findex]���Ѻ�ħ��ʹ�û�������
        SendMessageRoom Index, data(i)
        'AddLog "���� " & Winsock2(Index).Tag & "��ħ��" & Run(1) & "��ʹ�û�������"
        Case "dm" '[dm findex]ɾ��ħ��
        SendMessageRoom Index, data(i)
        'AddLog "���� " & Winsock2(Index).Tag & "��ħ��" & Run(1) & "��ɾ��"
        Case "ef" '[ef effect eindex]��Чת��
        'effect ice �� fire �� che ��
        SendMessageRoom Index, data(i)
        'AddLog "���� " & Winsock2(Index).Tag & "���ж�ħ��" & Run(2) & "������Buff " & Run(1)
    End Select
    
    End If '���жϣ�ǰ����If Data(i) <> "" Then
    
Next
'sth:
'If Err.Number <> 0 Then
'Msgbox Err.Description, "Error"
'End If
End Sub

