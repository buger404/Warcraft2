Attribute VB_Name = "modEffectDraw"
Function NewEffect(HDC As Long, ByVal WindowName As String, ByVal X As Single, ByVal Y As Single, Types As Monster2_Effect, ByVal Tag As String, ByVal Durings As Long)
ReDim Preserve GameEffect(UBound(GameEffect) + 1)
With GameEffect(UBound(GameEffect))
.EffectWindow = WindowName
.EffectTag = Tag
.EffectType = Types
.X = X
.Y = Y
.During = Durings
.EffectIndex = UBound(GameEffect)
.EHDC = HDC
NewEffect = .EffectIndex
End With
End Function
Sub DeleteEffect(ByVal Index As Integer)
If Index <= UBound(GameEffect) Then
Set GameEffect(Index) = GameEffect(UBound(GameEffect))
GameEffect(Index).EffectIndex = Index
ReDim Preserve GameEffect(UBound(GameEffect) - 1)
End If
End Sub
Sub DrawEffect(UI As Long, WindowName As String)
BGMBox.DrawScreen ActiveWindow
If UBound(GameEffect) > 0 Then
For i = 1 To UBound(GameEffect)
If i > UBound(GameEffect) Then Exit Sub
If GameEffect(i).EffectWindow = WindowName Then
GameEffect(i).Update (UI)
End If
Next
End If
End Sub
Sub DrawPicWithBlur(X As Single, Y As Single, W As Long, H As Long, UIGraphics As Long, Optional Radius As Long = 50)
Dim p As BlurParams
Dim bitmap As Long, Effect As Long, bitmapoutput As Long

GdipCreateEffect2 GdipEffectType.Blur, Effect

p.Radius = Radius
p.expandEdge = 0

GdipCreateBitmapFromGraphics 556, 949, UIGraphics, bitmap
'GdipCreateBitmapFromFile App.Path & "\assets\lost.png", bitmap

GdipSetEffectParameters Effect, p, LenB(p)
'GdipBitmapApplyEffect bitmap, Effect, NewRectL(X, Y, W, H), 0, 0, 0
GdipBitmapCreateApplyEffect bitmap, 1, Effect, NewRectL(X, Y, W, H), NewRectL(0, 0, 0, 0), bitmapoutput, 0, 0, 0
GdipDrawImage UIGraphics, bitmapoutput, 0, 0

GdipDisposeImage bitmap
GdipDisposeImage bitmapoutput

GdipDeleteEffect Effect
End Sub
