Attribute VB_Name = "rs_Transitions"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   控件过场动画模块-Redstone Supremacy
'   Copyright 2016-2017 Redstone Supremacy . All rights reserved.
'   Error 404(QQ 1361778219)
'   Version : 1.0.0
'   未经本人同意请勿修改本模块
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Enum Transitions_Value
    LabelValue = 1
    ImageXValue = 2
    OtherControls = 3
End Enum
Public Enum Transitions_Mode
    Fromlefttoright = 1
    Fromrighttoleft = 2
    upstairs = 3
    downStairs = 4
End Enum
    

Sub TextFromZero(Control As Object, ByVal NewMath As Double, Time As Long)
Control.Caption = 0
Dim a As Double, b As Double
a = NewMath / Time
For i = 1 To Time
b = b + a
Control.Caption = Int(b)
Sleep 10: DoEvents
Next
Control.Caption = NewMath
End Sub
Sub TextShowFromCenter(Control As Object, Text As String)
Dim a As Long, b As Long
a = Int(Len(Text) / 2)
b = a + 1
Control.Caption = Mid(Text, a, 1) & Mid(Text, b, 1)
Do While a > 0 Or b < Len(Text) + 1
a = a - 1
b = b + 1
If a > 0 Then Control.Caption = Mid(Text, a, 1) & Control.Caption
If b < Len(Text) + 1 Then Control.Caption = Control.Caption & Mid(Text, b, 1)
Sleep 50: DoEvents
Loop
End Sub
Sub Magic(Control, newx, newy, newW, newH, Optional Duration As Integer = 10)
    a = (newx - Control.Left) / Duration
    b = (newy - Control.Top) / Duration
    c = (newW - Control.Width) / Duration
    d = (newH - Control.Height) / Duration
    For i = 1 To Duration
        Control.Move Control.Left + a, Control.Top + b, Control.Width + c, Control.Height + d
        Sleep 10: DoEvents
    Next
    
End Sub
Sub 居中(Control, Parents)
    Control.Left = Parents.Left + Parents.Width / 2 - Control.Width / 2
    Control.Top = Parents.Top + Parents.Height / 2 - Control.Height / 2
End Sub
Sub FlyTo(Control, newx, newy, Optional Duration As Integer = 10)
    a = (newx - Control.Left) / Duration
    b = (newy - Control.Top) / Duration
    For i = 1 To Duration
        Control.Move Control.Left + a, Control.Top + b
        Sleep 10: DoEvents
    Next
End Sub
Sub FRGBTo(Control, oldr, oldG, oldB, NewR, NewG, NewB, Optional Duration As Integer = 10)
    a = (NewR - oldr) / Duration
    b = (NewG - oldG) / Duration
    c = (NewB - oldB) / Duration
    For i = 1 To Duration
        Control.ForeColor = RGB(oldr + a * i, oldG + b * i, oldB + c * i)
        Sleep 10: DoEvents
    Next
End Sub
Sub BRGBTo(Control, oldr, oldG, oldB, NewR, NewG, NewB, Optional Duration As Integer = 10)
    a = (NewR - oldr) / Duration
    b = (NewG - oldG) / Duration
    c = (NewB - oldB) / Duration
    For i = 1 To Duration
        Control.BackColor = RGB(oldr + a * i, oldG + b * i, oldB + c * i)
        Sleep 10: DoEvents
    Next
End Sub
Sub 抖动(Control, Optional Counts As Integer = 10, Optional Duration As Integer = 10)
    For i = 1 To Counts
        Control.Left = Control.Left + 60
        Control.Top = Control.Top + 60
        Sleep Duration: DoEvents
        Control.Left = Control.Left - 120
        Control.Top = Control.Top - 120
        Sleep Duration: DoEvents
        Control.Left = Control.Left - 60
        Control.Top = Control.Top - 60
        Sleep Duration: DoEvents
        Control.Left = Control.Left + 120
        Control.Top = Control.Top + 120
        Sleep Duration: DoEvents
    Next
End Sub
Sub 大小抖动(Control, Optional Counts As Integer = 10, Optional Duration As Integer = 10)
    For i = 1 To Counts
        Control.Width = Control.Width + (Counts - i)
        Control.Height = Control.Height + (Counts - i)
        Control.Left = Control.Left - (Counts - i) / 2
        Control.Top = Control.Top - (Counts - i) / 2
        Sleep Duration: DoEvents
        Control.Width = Control.Width - (Counts - i)
        Control.Height = Control.Height - (Counts - i)
        Control.Left = Control.Left + (Counts - i) / 2
        Control.Top = Control.Top + (Counts - i) / 2
        Sleep Duration: DoEvents
    Next
End Sub
Sub FrameTransitions(Frame1 As Frame, Frame2 As Frame, Optional CMode As Transitions_Mode = Fromrighttoleft, Optional Duration As Integer = 1)
'Duration = Duration * 10
ox = Frame1.Left
oy = Frame1.Top
ox2 = Frame2.Left
oy2 = Frame2.Top
If CMode = Fromlefttoright Or CMode = Fromrighttoleft Then
skill = (Frame1.parent.Width) / Duration
Else
skill = (Frame1.parent.Height) / Duration
End If

            Select Case CMode
                Case Is = Transitions_Mode.Fromlefttoright
                    Frame2.Left = -Frame2.Width
                Case Is = Transitions_Mode.Fromrighttoleft
                    Frame2.Left = Frame2.parent.Width
                Case Is = Transitions_Mode.upstairs
                    Frame2.Top = -Frame2.Height
                Case Is = Transitions_Mode.downStairs
                    Frame2.Top = Frame2.parent.Height
            End Select

    Frame1.Visible = True
    Frame2.Visible = True

        For i = 1 To Duration
            DoEvents
            Select Case CMode
                Case Is = Transitions_Mode.Fromlefttoright
                    Frame1.Left = Frame1.Left + skill
                    Frame2.Left = Frame2.Left + skill
                Case Is = Transitions_Mode.Fromrighttoleft
                    Frame1.Left = Frame1.Left - skill
                    Frame2.Left = Frame2.Left - skill
                Case Is = Transitions_Mode.upstairs
                    Frame1.Top = Frame1.Top + skill
                    Frame2.Top = Frame2.Top + skill
                Case Is = Transitions_Mode.downStairs
                    Frame1.Top = Frame1.Top - skill
                    Frame2.Top = Frame2.Top - skill
            End Select
            Sleep 10
        Next
        Frame2.Left = ox2
        Frame2.Top = oy2
        Frame1.Visible = False
        Frame1.Left = ox
        Frame1.Top = oy
End Sub
Sub Transitions(Control As Object, Optional mode As Boolean = False, Optional CMode As Transitions_Mode = upstairs, Optional Duration As Integer = 1, Optional CType As Transitions_Value = OtherControls)
    On Error Resume Next
    '''''''''''关于CMode'''''''''''''''''
    '  FromLeftToRight  从左到右
    '  FromRightToLeft  从右到左
    '  UpStairs 向上
    '  DownStairs 向下
    '''''''''''关于Mode'''''''''''''''''''
    '''''True：出场
    '''''False：入场
    '''''''''''关于Duration''''''''''''''
    '''''持续时间（单位：毫秒）
    '''''''''''关于CType''''''''''''''''''
    '''''控件类型
    '''''''''''''''''''''''''''''''''''''''''''''''''

    Control.Visible = True
    If mode = False And CType = OtherControls Then
        
        If CMode = 1 Then Control.Left = Control.Left - 500
        If CMode = 2 Then Control.Left = Control.Left + 500
        If CMode = 3 Then Control.Top = Control.Top + 500
        If CMode = 4 Then Control.Top = Control.Top - 500
        
        For i = 1 To Duration                                                   '1 To 5
            DoEvents
            If CMode = 1 Then Control.Left = Control.Left + 500 / Duration      '- 100
            If CMode = 2 Then Control.Left = Control.Left - 500 / Duration      '- 100
            If CMode = 3 Then Control.Top = Control.Top - 500 / Duration        '- 100
            If CMode = 4 Then Control.Top = Control.Top + 500 / Duration        '- 100
            Sleep 10                                                             'Int(Duration / 5)
        Next
    End If
    
    If mode = True And CType = OtherControls Then
        b = Control.Left
        c = Control.Top
        For i = 1 To Duration                                                   '1 To 5
            DoEvents
            If CMode = 1 Then Control.Left = Control.Left - 500 / Duration      '- 100
            If CMode = 2 Then Control.Left = Control.Left + 500 / Duration      '- 100
            If CMode = 3 Then Control.Top = Control.Top + 500 / Duration        '- 100
            If CMode = 4 Then Control.Top = Control.Top - 500 / Duration        '- 100
            Sleep 10                                                             'Int(Duration / 5)
        Next
        Control.Visible = False
        Control.Left = b
        Control.Top = c
    End If
    
    If mode = False And CType = LabelValue Then
        b = Control.Left
        c = Control.Top
        d = Control.Width
        e = Control.Height
        a = Control.FontSize
        Control.FontSize = Int(Control.FontSize * 0.7)
        Control.Left = b + d / 2 - Control.Width / 2
        Control.Top = c + e / 2 - Control.Height / 2
        a = a - Control.FontSize
        For i = 1 To Duration                                                   '1 To 5
            DoEvents
            Control.FontSize = Control.FontSize + a / Duration                  'a / 5
            Control.Left = b + d / 2 - Control.Width / 2
            Control.Top = c + e / 2 - Control.Height / 2
            Sleep 10                                                             'Int(Duration / 5)
        Next
    End If
    
    If mode = True And CType = LabelValue Then
        b = Control.Left
        c = Control.Top
        d = Control.Width
        e = Control.Height
        a = Control.FontSize
        c = Control.FontSize
        Control.FontSize = Int(Control.FontSize * 0.7)
        Control.Left = b + d / 2 - Control.Width / 2
        Control.Top = c + e / 2 - Control.Height / 2
        a = a - Control.FontSize
        Control.FontSize = c
        For i = 1 To Duration                                                   '1 To 5
            DoEvents
            Control.FontSize = Control.FontSize - a / Duration                  'a / 5
            Control.Left = b + d / 2 - Control.Width / 2
            Control.Top = c + e / 2 - Control.Height / 2
            Sleep 10                                                             'Int(Duration / 5)
        Next
        Control.FontSize = c
        Control.Visible = False
    End If
    
    If mode = False And CType = ImageXValue Then
        a = Control.Opacity
        Control.Opacity = 0
        If CMode = 1 Then Control.Left = Control.Left - 500
        If CMode = 2 Then Control.Left = Control.Left + 500
        If CMode = 3 Then Control.Top = Control.Top + 500
        If CMode = 4 Then Control.Top = Control.Top - 500
        For i = 1 To Duration                                                   '1 To 5
            DoEvents
            If CMode = 1 Then Control.Left = Control.Left + 500 / Duration      '- 100
            If CMode = 2 Then Control.Left = Control.Left - 500 / Duration      '- 100
            If CMode = 3 Then Control.Top = Control.Top - 500 / Duration        '- 100
            If CMode = 4 Then Control.Top = Control.Top + 500 / Duration        '- 100
            Control.Opacity = Control.Opacity + a / Duration                    'a / 5
            Sleep 10                                                             'Int(Duration / 5)
        Next
    End If
    
    If mode = True And CType = ImageXValue Then
        a = Control.Opacity
        b = Control.Left
        c = Control.Top
        For i = 1 To Duration                                                   '1 To 5
            DoEvents
            If CMode = 1 Then Control.Left = Control.Left - 500 / Duration      '- 100
            If CMode = 2 Then Control.Left = Control.Left + 500 / Duration      '- 100
            If CMode = 3 Then Control.Top = Control.Top + 500 / Duration        '- 100
            If CMode = 4 Then Control.Top = Control.Top - 500 / Duration        '- 100
            Control.Opacity = Control.Opacity - a / Duration                    'a / 5
            Sleep 10                                                             'Int(Duration / 5)
        Next
        Control.Visible = False
        Control.Opacity = a
        Control.Left = b
        Control.Top = c
    End If
    
End Sub
Sub Download(Url As String, Path As String)
        Dim XmlHttp, b() As Byte
        Set XmlHttp = CreateObject("Msxml2.ServerXMLHTTP.3.0")
        XmlHttp.Open "Get", Url, True
        
        XmlHttp.Send
        temp = Timer
        Do While XmlHttp.ReadyState <> 4
        DoEvents
        Select Case XmlHttp.ReadyState
        Case Is = 1
        Form1.Label19.Caption = "正在连接服务器 (" & Int(Timer - temp) & "s)"
        Case Is = 2
        Form1.Label19.Caption = "正在等待服务器响应 (" & Int(Timer - temp) & "s)"
        Case Is = 3
        Form1.Label19.Caption = "正在下载 (" & Int(Timer - temp) & "s)"
        End Select
        If Timer - temp > 10 And XmlHttp.ReadyState = 1 Then
        Form1.Say "无法连接到服务器，安装程序无法进行。", "运行时错误"
        Exit Do
        End If
        Sleep 1
        Loop
        
        If XmlHttp.ReadyState <> 4 Then Exit Sub
        
            b() = XmlHttp.ResponseBody
            Open Path For Binary As #1
            Put #1, , b(): DoEvents
            Close #1
        
        Set XmlHttp = Nothing
End Sub
