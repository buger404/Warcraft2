Attribute VB_Name = "modNoCheat"
Public Type EVENTMSG
 vKey As Long
 sKey As Long
 flag As Long
 time As Long
 End Type
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public mymsg As EVENTMSG
Public Const WH_KEYBOARD_LL = 13
Public Const WM_KEYDOWN = &H100
Public hHook&, i%, ggg$, s1$, s2$, pos1$(), pos2$(), Scxx
Sub ints()
s1 = "96 97 98 99 100 101 102 103 104 105 106 107 109 110 111 13 " + _
"144 65 66 67 68 69 70 71 72 73 74 75 76 77 78 79 80 81 82 83 84 " + _
"85 86 87 88 89 90 48 49 50 51 52 53 54 55 56 57 192 189 187 220 8 " + _
"44 45 46 145 36 35 19 33 34 38 40 37 39 27 112 113 114 115 116 117 " + _
"118 119 120 121 122 123 9 20 160 162 91 13 161 92 93"
s2 = "小0 小1 小2 小3 小4 小5 小6 小7 小8 小9 小* 小+ 小- 小. 小/ " + _
"小Enter 小NumLock A B C D E F G H I G K L M N O P Q R S T U V W X Y Z " + _
"0 1 2 3 4 5 6 7 8 9 ` - = \ BackSpace " + _
"PrintScreen Insert Delete ScrollLock Home End PauseBreak PageUp PageDown " + _
"上 下 左 右 ESC F1 F2 F3 F4 F5 F6 F7 F8 F9 F10 F11 F12 " + _
"TAB CapsLock 左Shift 左Ctrl 左Win Enter 右Shift 右Win 右List 右Ctrl"
pos1 = Split(s1, " "): pos2 = Split(s2, " ")
End Sub
Public Function MyKBHook(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If ncode = 0 Then
    If wParam = WM_KEYDOWN Then
        CopyMemory mymsg, ByVal lParam, Len(mymsg)
         'For i = 0 To UBound(pos1) - 1
         'If mymsg.vKey = Val(pos1(i)) Then
        'If pos2(i) = "F2" Then Form1.Show
         'End If
         'Next
    End If
End If
MyKBHook = CallNextHookEx(0, ncode, wParam, lParam) '就是不把消息还给你~
End Function




