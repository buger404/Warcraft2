VERSION 5.00
Begin VB.UserControl UIButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0B000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Height          =   15
      Left            =   1500
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   15
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label2 
      Height          =   15
      Left            =   750
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hello Label"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   1650
      Width           =   975
   End
End
Attribute VB_Name = "UIButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
                                                                                '事件声明:
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Event Click()                                                                   'MappingInfo=Label1,Label1,-1,Click
Attribute Click.VB_Description = "按钮被按下"
Event DblClick()                                                                'MappingInfo=Label1,Label1,-1,DblClick
Attribute DblClick.VB_Description = "按钮被双击"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)  'MappingInfo=Label1,Label1,-1,MouseDown
Attribute MouseDown.VB_Description = "按钮被按下"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)  'MappingInfo=Label1,Label1,-1,MouseMove
Attribute MouseMove.VB_Description = "鼠标经过按钮"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)    'MappingInfo=Label1,Label1,-1,MouseUp
Attribute MouseUp.VB_Description = "鼠标在按钮上抬起"
Event MouseExit()
Attribute MouseExit.VB_Description = "鼠标从按钮移开"
Dim Captureing As Boolean
                                                                                '注意！不要删除或修改下列被注释的行！
                                                                                'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "按钮正常状态下的颜色"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Label3.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

                                                                                '注意！不要删除或修改下列被注释的行！
                                                                                'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "按钮的文本颜色"
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

                                                                                '注意！不要删除或修改下列被注释的行！
                                                                                'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "是否让按钮可以响应用户的行为"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Icon() As StdPicture
Attribute Icon.VB_Description = "在按钮上显示一个图标"
'    Set Icon = Image1.Picture
End Property

Public Property Set Icon(ByVal New_Icon As StdPicture)
'    Set Image1.Picture = New_Icon
  '  If Not (New_Icon Is Nothing) Then Image1.Visible = True
    PropertyChanged "Icon"
End Property

                                                                                '注意！不要删除或修改下列被注释的行！
                                                                                'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

                                                                                '注意！不要删除或修改下列被注释的行！
                                                                                'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "指出 Label 或 Shape 的背景样式是透明的还是不透明的。"
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Private Sub Image1_Click()

End Sub

Private Sub Label1_Click()
    RaiseEvent Click
End Sub

Private Sub Label1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.BackColor = Label4.BackColor
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, Label1.Left + X, Label1.top + Y)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    UserControl.BackColor = Label2.BackColor
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Captureing = True Then ReleaseCapture: Captureing = False
    UserControl.BackColor = Label3.BackColor
    
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.BackColor = Label4.BackColor
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then UserControl.BackColor = Label2.BackColor
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If X >= 0 And X <= UserControl.Width And Y >= 0 And Y <= UserControl.Height Then
        
        If Captureing = False Then
            SetCapture UserControl.hWnd
            Captureing = True
        End If
        
    Else
        If Captureing = True Then ReleaseCapture: Captureing = False
        UserControl.BackColor = Label3.BackColor
        RaiseEvent MouseExit
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UserControl.BackColor = Label2.BackColor
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Captureing = True Then ReleaseCapture: Captureing = False
    UserControl.BackColor = Label3.BackColor
    
End Sub

                                                                                '注意！不要删除或修改下列被注释的行！
                                                                                'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "按钮显示的文本"
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
    UserControl_Resize
End Property

Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_Description = "鼠标经过按钮时所呈现的颜色"
    BackColor2 = Label2.BackColor
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
    Label2.BackColor() = New_BackColor2
    PropertyChanged "BackColor2"
End Property

Public Property Get BackColor3() As OLE_COLOR
Attribute BackColor3.VB_Description = "按钮被按下时所呈现的颜色"
    BackColor3 = Label4.BackColor
End Property

Public Property Let BackColor3(ByVal New_BackColor3 As OLE_COLOR)
    Label4.BackColor() = New_BackColor3
    PropertyChanged "BackColor3"
End Property

                                                                                '注意！不要删除或修改下列被注释的行！
                                                                                'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "设置一个自定义鼠标图标。"
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

                                                                                '注意！不要删除或修改下列被注释的行！
                                                                                'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "返回/设置当鼠标经过对象某一部分时鼠标的指针类型。"
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
Sub Refresh()
Attribute Refresh.VB_Description = "刷新控件"
UserControl.Refresh
End Sub
Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub
Sub OnLeft()
Attribute OnLeft.VB_Description = "使按钮的文本位于左侧"
    Label1.Left = 500
End Sub
Function GetHwnd()
Attribute GetHwnd.VB_Description = "获得按钮的句柄"
    GetHwnd = UserControl.hWnd
End Function
                                                                                '从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label3.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label2.BackColor = PropBag.ReadProperty("BackColor2", &H8000000F)
    Label4.BackColor = PropBag.ReadProperty("BackColor3", &H8000000D)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 0)
    Label1.Caption = PropBag.ReadProperty("Caption", "Hello Label")
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Icon = PropBag.ReadProperty("Icon", Nothing)
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    Label1.Move UserControl.Width / 2 - Label1.Width / 2, UserControl.Height / 2 - Label1.Height / 2
                                                                                'Label5.Move 0, 0, UserControl.Width, UserControl.Height
'    Image1.Move 300, UserControl.Height / 2 - Image1.Height / 2
    
End Sub

                                                                                '将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 0)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Hello Label")
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("BackColor2", Label2.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackColor3", Label4.BackColor, &H8000000D)
    Call PropBag.WriteProperty("Icon", Icon, Nothing)
    UserControl_Resize
End Sub

