VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EffectManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public EffectType As Monster2_Effect, EffectTag As String, EffectCount As Long
Public EffectWindow As String, EffectIndex As Integer, During As Long
Public X As Single, Y As Single, EHDC As Long
Dim EffectTexts() As String
Dim LastEffectTime As Long
Sub ChangeText(ByVal NewText As String)
EffectTag = NewText
ReDim EffectTexts(Len(EffectTag))
For i = 1 To Len(EffectTag)
EffectTexts(i) = LCase(Mid(EffectTag, i, 1))
Next
End Sub
Sub Update(Graphics As Long)
On Error Resume Next
Select Case EffectType
    Case Monster2_Effect.TaketurnsPic
    Call TaketurnsPicUpdate(Graphics)
    Case Monster2_Effect.MagicText
    Call MagicTextUpdate(Graphics)
    Case Monster2_Effect.FadeInPic
    Call FadeUpdate(Graphics)
End Select
End Sub
Private Sub TaketurnsPicUpdate(Graphics As Long)
If During <> 0 Then
If GetPic(EffectTag & EffectCount) <= 0 Then DeleteEffect (EffectIndex): Exit Sub
End If
If GetPic(EffectTag & EffectCount) > 0 Then
GamePictures(GetPic(EffectTag & EffectCount)).NextFrame.Present EHDC, X, Y, 255
ElseIf During = 0 Then
EffectCount = 0
End If
If GetTickCount - LastEffectTime >= 40 Then EffectCount = EffectCount + 1: LastEffectTime = GetTickCount
End Sub

Private Sub MagicTextUpdate(Graphics As Long)
Dim temp As Byte

If EffectCount = 0 Then
If UBound(EffectTexts) = 0 Then
    ReDim EffectTexts(Len(EffectTag))
    For i = 1 To Len(EffectTag)
    EffectTexts(i) = LCase(Mid(EffectTag, i, 1))
    Next
End If
End If

If EffectCount > During Then DeleteEffect (EffectIndex): Exit Sub
'GdipSetStringFormatAlign strformat, StringAlignmentNear
If GetTickCount - LastEffectTime >= 30 Then
EffectCount = EffectCount + 1
LastEffectTime = GetTickCount
End If

If EffectCount / During < 0.5 Then
    temp = EffectCount / (During / 2) * 255
    'LastPenColor = temp
    'GdipSetSolidFillColor Brush1, temp
    temp2 = EffectCount / During * 16
    For i = 1 To UBound(EffectTexts)
    'If i = 2 Then Dad.Caption = "char" & EffectTexts(i) & "," & GetPic("char" & EffectTexts(i)) & "," & EHDC & "=?" & WelcomeWindow.HDC & ";" & temp
    GamePictures(GetPic("char" & EffectTexts(i))).NextFrame.Present EHDC, X + (i - Len(EffectTag) / 2) * 1.5 * temp2, Y, temp
    Next

Else
    temp = 255 - (EffectCount - (During / 2)) / (During / 2) * 255
    'LastPenColor = temp
    'GdipSetSolidFillColor Brush1, temp
    temp2 = EffectCount / During * 16
    For i = 1 To Len(EffectTag)
    GamePictures(GetPic("char" & EffectTexts(i))).NextFrame.Present EHDC, X + (i - Len(EffectTag) / 2) * 1.5 * temp2, Y, temp
    Next
End If
'DrawTextRectUI Graphics, X, Y, temp, argb(255 - EffectCount / During * 255, 255, 255, 255), StringAlignmentCenter, True
End Sub
Private Sub FadeUpdate(Graphics As Long)
If EffectCount > During Then DeleteEffect (EffectIndex): Exit Sub
If EffectCount / During < 0.5 Then
GamePictures(GetPic(EffectTag)).NextFrame.Present EHDC, X, Y, EffectCount / (During / 2) * 255
Else
GamePictures(GetPic(EffectTag)).NextFrame.Present EHDC, X, Y, 255 - (EffectCount - (During / 2)) / (During / 2) * 255
End If
If GetTickCount - LastEffectTime >= 30 Then EffectCount = EffectCount + 1: LastEffectTime = GetTickCount
End Sub

Private Sub Class_Initialize()
ReDim EffectTexts(0)
End Sub
