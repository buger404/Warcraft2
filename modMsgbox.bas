Attribute VB_Name = "modMsgbox"
Public CloseInfo As Boolean
Sub ShowInfo(Text As String, Pic As String, Optional Wave As String = "get")
        Dim lngDesktopHwnd As Long
        Dim lngDesktopDC As Long
        Dim Window As New InfoWindow

        Ring Wave
        SetParent AeroWindow.hWnd, Dad.hWnd
        AeroWindow.Move 0, 0, Dad.Width, Dad.Height
        
        lngDesktopHwnd = Dad.hWnd
        lngDesktopDC = GetDC(lngDesktopHwnd)

        Set AeroWindow.Picture = Nothing

        BitBlt AeroWindow.hDC, 0, 0, Dad.ScaleWidth, Dad.ScaleHeight, lngDesktopDC, 0, 0, SRCCOPY

        Call ReleaseDC(lngDesktopHwnd, lngDesktopDC)
        
        BlurPic AeroWindow, 100
        
        Window.infotext = Text
        Window.InfoPic = Pic
        Window.Draw
        Window.Show
        SetParent Window.hWnd, AeroWindow.ToolFrame.hWnd
        
        Window.Move 0, 0
        AeroWindow.ToolFrame.Move AeroWindow.Width / 15 / 2 - Window.Width / 15 / 2, AeroWindow.Height / 15 / 2 - Window.Height / 15 / 2, Window.Width / 15, Window.Height / 15
        
        AeroWindow.Show
        AeroWindow.ZOrder
        
        Wait 3000
        Unload Window
        AeroWindow.Hide
End Sub
Sub ShowToolWindow(Window As Form, Optional UseTran As Boolean = True)
        SetParent AeroWindow.hWnd, Dad.hWnd
        AeroWindow.Move 0, 0, Dad.Width, Dad.Height
        Dim lngDesktopHwnd As Long
        Dim lngDesktopDC As Long

        lngDesktopHwnd = Dad.hWnd
        lngDesktopDC = GetDC(lngDesktopHwnd)

        Set AeroWindow.Picture = Nothing

        BitBlt AeroWindow.hDC, 0, 0, Dad.ScaleWidth, Dad.ScaleHeight, lngDesktopDC, 0, 0, SRCCOPY

        Call ReleaseDC(lngDesktopHwnd, lngDesktopDC)
        
        BlurPic AeroWindow, 50
        
        Window.Show
        SetParent Window.hWnd, AeroWindow.ToolFrame.hWnd
        
        Window.Move 0, 0
        AeroWindow.ToolFrame.Move AeroWindow.Width / 15 / 2 - Window.Width / 15 / 2, AeroWindow.Height / 15 / 2 - Window.Height / 15 / 2, Window.Width / 15, Window.Height / 15
        
        AeroWindow.Show
        AeroWindow.ZOrder
        
        If UseTran = True Then Transitions AeroWindow.ToolFrame, , upstairs, 5
End Sub
Function Inputbox(ByVal Text As String, Optional ByVal Title As String = "", Optional ByVal NormalText As String = "请输入...")
        Dim Window As New MsgWindow
        Window.Move 0, 0, Dad.ScaleWidth * 15, Dad.ScaleHeight * 15
        
        Dim lngDesktopHwnd As Long
        Dim lngDesktopDC As Long

        lngDesktopHwnd = Dad.hWnd
        lngDesktopDC = GetDC(lngDesktopHwnd)

        Set Window.Picture = Nothing

        BitBlt Window.hDC, 0, 0, Dad.ScaleWidth, Dad.ScaleHeight, lngDesktopDC, 0, 0, SRCCOPY

        Call ReleaseDC(lngDesktopHwnd, lngDesktopDC)
        
        BlurPic Window, 50
        
        Window.Label2.Caption = Title
        Window.Label3.Caption = Text
        SetParent Window.hWnd, Dad.hWnd
        
        Window.Text1.Text = NormalText
        Window.Text1.Visible = True
        
        Window.Show
        Window.SetScroll
        Transitions Window.Frame1, , upstairs, 5
        
        Window.Text1.SetFocus
        Window.Text1.SelStart = 0
        Window.Text1.SelLength = Len(Window.Text1.Text)
        
        Do While Window.ClickButton = 0
        DoEvents
        Sleep 10
        Loop
        
        Transitions Window.Frame1, True, upstairs, 5
        
        If Window.ClickButton = 1 Then
        Inputbox = ""
        Else
        Inputbox = Window.Text1.Text
        End If
        
        Unload Window
End Function
Sub InfoBox(ByVal Text As String, Optional ByVal 没用的参数 As Integer = 0, Optional ByVal Title As String = "")
        CloseInfo = False
        Dim Window As New MsgWindow
        Window.Move 0, 0, Dad.ScaleWidth * 15, Dad.ScaleHeight * 15
        
        Dim lngDesktopHwnd As Long
        Dim lngDesktopDC As Long

        lngDesktopHwnd = Dad.hWnd
        lngDesktopDC = GetDC(lngDesktopHwnd)

        Set Window.Picture = Nothing

        BitBlt Window.hDC, 0, 0, Dad.ScaleWidth, Dad.ScaleHeight, lngDesktopDC, 0, 0, SRCCOPY

        Call ReleaseDC(lngDesktopHwnd, lngDesktopDC)
        
        BlurPic Window, 50
        
        Window.Label2.Caption = Title
        Window.Label3.Caption = Text
        Window.UIButton2.Visible = False
        SetParent Window.hWnd, Dad.hWnd
        
        Window.Show
        Window.SetScroll
        Transitions Window.Frame1, , upstairs, 5
        
        
        Window.Refresh
        Do While Window.ClickButton = 0 And CloseInfo = False
        DoEvents
        Sleep 10
        Loop
        
        Transitions Window.Frame1, True, upstairs, 5
        
        Unload Window
        
End Sub
Function Msgbox(ByVal Text As String, Optional ByVal 没用的参数 As Integer = 0, Optional ByVal Title As String = "")
        Dim Window As New MsgWindow
        Window.Move 0, 0, Dad.ScaleWidth * 15, Dad.ScaleHeight * 15
        
        Dim lngDesktopHwnd As Long
        Dim lngDesktopDC As Long

        lngDesktopHwnd = Dad.hWnd
        lngDesktopDC = GetDC(lngDesktopHwnd)

        Set Window.Picture = Nothing

        BitBlt Window.hDC, 0, 0, Dad.ScaleWidth, Dad.ScaleHeight, lngDesktopDC, 0, 0, SRCCOPY

        Call ReleaseDC(lngDesktopHwnd, lngDesktopDC)
        
        BlurPic Window, 50
        
        Window.Label2.Caption = Title
        Window.Label3.Caption = Text
        SetParent Window.hWnd, Dad.hWnd
        
        Window.Show
        Window.SetScroll
        Transitions Window.Frame1, , upstairs, 5
        
        
        Window.Refresh
        Do While Window.ClickButton = 0
        DoEvents
        Sleep 10
        Loop
        
        Transitions Window.Frame1, True, upstairs, 5
        
        Msgbox = Window.ClickButton
        Unload Window
        
End Function
Sub BlurPic(Pic As Object, Optional Radius As Long = 40, Optional Shadow As Integer = 0)
On Error Resume Next
Dim Fattri As Long, p As BlurParams, graphics2 As Long
Dim bitmap As Long, Effect As Long, bitmapoutput As Long
GdipCreateEffect2 GdipEffectType.Blur, Effect
p.Radius = Radius
p.expandEdge = Shadow

Set Pic.Picture = Pic.Image
Pic.Refresh
GdipCreateFromHDC Pic.hDC, graphics2
GdipCreateBitmapFromHBITMAP Pic.Image.handle, 0, bitmap

GdipSetEffectParameters Effect, p, LenB(p)
GdipBitmapCreateApplyEffect bitmap, 1, Effect, NewRectL(0, 0, Pic.ScaleWidth, Pic.ScaleHeight), NewRectL(0, 0, 0, 0), bitmapoutput, 0, 0, 0
GdipDrawImageRect graphics2, bitmapoutput, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight

Pic.Refresh
GdipDisposeImage bitmap
GdipDisposeImage bitmapoutput

GdipDeleteEffect Effect

End Sub
