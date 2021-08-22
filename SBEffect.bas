Attribute VB_Name = "modWindowEffect"
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Sub FadeIn(Time As Long, mb As Form)
    mb.Visible = True
    a = 255 / Time
    b = 0
    For i = 0 To Time
        
        Dim sty As Long
        
        sty = GetWindowLong(mb.hWnd, GWL_EXSTYLE)
        sty = sty Or WS_EX_LAYERED
        SetWindowLong mb.hWnd, GWL_EXSTYLE, sty
        
        SetLayeredWindowAttributes mb.hWnd, mb.BackColor, b, LWA_ALPHA
        b = b + a
        Sleep 10: DoEvents
    Next
    
End Sub
Public Sub FadeOut(Time As Long, mb As Form)
    a = 255 / Time
    For i = 1 To Time
        b = b + a
        Dim sty As Long
        sty = GetWindowLong(mb.hWnd, GWL_EXSTYLE)
        sty = sty Or WS_EX_LAYERED
        SetWindowLong mb.hWnd, GWL_EXSTYLE, sty
        SetLayeredWindowAttributes mb.hWnd, mb.BackColor, 255 - b, LWA_ALPHA
        Sleep 10: DoEvents
    Next
    
    mb.Visible = False
End Sub
