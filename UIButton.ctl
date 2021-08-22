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
                                                                                '�¼�����:
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Event Click()                                                                   'MappingInfo=Label1,Label1,-1,Click
Attribute Click.VB_Description = "��ť������"
Event DblClick()                                                                'MappingInfo=Label1,Label1,-1,DblClick
Attribute DblClick.VB_Description = "��ť��˫��"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)  'MappingInfo=Label1,Label1,-1,MouseDown
Attribute MouseDown.VB_Description = "��ť������"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)  'MappingInfo=Label1,Label1,-1,MouseMove
Attribute MouseMove.VB_Description = "��꾭����ť"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)    'MappingInfo=Label1,Label1,-1,MouseUp
Attribute MouseUp.VB_Description = "����ڰ�ť��̧��"
Event MouseExit()
Attribute MouseExit.VB_Description = "���Ӱ�ť�ƿ�"
Dim Captureing As Boolean
                                                                                'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
                                                                                'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "��ť����״̬�µ���ɫ"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Label3.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

                                                                                'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
                                                                                'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "��ť���ı���ɫ"
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

                                                                                'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
                                                                                'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "�Ƿ��ð�ť������Ӧ�û�����Ϊ"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Icon() As StdPicture
Attribute Icon.VB_Description = "�ڰ�ť����ʾһ��ͼ��"
'    Set Icon = Image1.Picture
End Property

Public Property Set Icon(ByVal New_Icon As StdPicture)
'    Set Image1.Picture = New_Icon
  '  If Not (New_Icon Is Nothing) Then Image1.Visible = True
    PropertyChanged "Icon"
End Property

                                                                                'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
                                                                                'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "����һ�� Font ����"
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

                                                                                'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
                                                                                'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "ָ�� Label �� Shape �ı�����ʽ��͸���Ļ��ǲ�͸���ġ�"
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

                                                                                'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
                                                                                'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "��ť��ʾ���ı�"
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
    UserControl_Resize
End Property

Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_Description = "��꾭����ťʱ�����ֵ���ɫ"
    BackColor2 = Label2.BackColor
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
    Label2.BackColor() = New_BackColor2
    PropertyChanged "BackColor2"
End Property

Public Property Get BackColor3() As OLE_COLOR
Attribute BackColor3.VB_Description = "��ť������ʱ�����ֵ���ɫ"
    BackColor3 = Label4.BackColor
End Property

Public Property Let BackColor3(ByVal New_BackColor3 As OLE_COLOR)
    Label4.BackColor() = New_BackColor3
    PropertyChanged "BackColor3"
End Property

                                                                                'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
                                                                                'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "����һ���Զ������ͼ�ꡣ"
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

                                                                                'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
                                                                                'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "����/���õ���꾭������ĳһ����ʱ����ָ�����͡�"
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
Attribute Refresh.VB_Description = "ˢ�¿ؼ�"
UserControl.Refresh
End Sub
Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub
Sub OnLeft()
Attribute OnLeft.VB_Description = "ʹ��ť���ı�λ�����"
    Label1.Left = 500
End Sub
Function GetHwnd()
Attribute GetHwnd.VB_Description = "��ð�ť�ľ��"
    GetHwnd = UserControl.hWnd
End Function
                                                                                '�Ӵ������м�������ֵ
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

                                                                                '������ֵд���洢��
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

