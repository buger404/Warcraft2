Attribute VB_Name = "modScrollbar"
Option Explicit
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public PrevWndProc As Long
Private Booking As Boolean
Public Function WndProcForBookWindow(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Booking = True
    If uMsg = WM_MOUSEWHEEL Then
    If BookWindow.lastobj <> "ListFrame" And BookWindow.lastobj <> "scrollbutton" And BookWindow.lastobj <> "scrollbar" Then GoTo Last
    With BookWindow
        If wParam < 0 Then '¡ý
        .scrollbutton.Left = .scrollbutton.Left + 60 * Abs(wParam / 7864320)
        Else
        .scrollbutton.Left = .scrollbutton.Left - 60 * Abs(wParam / 7864320)
        End If
        If .scrollbutton.Left < .scrollbar.Left Then .scrollbutton.Left = .scrollbar.Left
        If .scrollbutton.Left > .scrollbar.Left + .scrollbar.Width - .scrollbutton.Width Then .scrollbutton.Left = .scrollbar.Left + .scrollbar.Width - .scrollbutton.Width
    End With
    End If
Last:
    WndProcForBookWindow = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)
End Function

