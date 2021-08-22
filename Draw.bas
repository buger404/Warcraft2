Attribute VB_Name = "modDraw"
Public Brush1 As Long
Public Pen1 As Long
Public LastPenColor As Long

'===============================字体要用到的================================
Public Brush As Long
Public fontfam As Long
Public strformat As Long
Public curFont As Long
Public curFontBig As Long
Public rclayout As RECTF
'=======================================================================

Sub DrawShadowRectangle(Graphics As Long, Color1 As Long, Color2 As Long, ByVal X As Single, ByVal Y As Single, ByVal W As Single, ByVal H As Single)
    Dim MixColor(1) As Long, MixPos(1) As Single, Path As Long, Brush As Long
    GdipCreatePath FillModeAlternate, Path
    GdipAddPathRectangle Path, X, Y, W, H
    GdipCreatePathGradientFromPath Path, Brush
    MixColor(0) = Color1: MixColor(1) = Color2: MixPos(0) = 0#: MixPos(1) = 1#
    GdipSetPathGradientPresetBlend Brush, MixColor(0), MixPos(0), 2
    GdipFillPath Graphics, Brush, Path
    GdipDeleteBrush Brush
    GdipDeletePath Path
End Sub
Sub DrawPicNameControlUI(UI As Long, Control As Object, PicName As String)
If Control.Visible = False Then Exit Sub
GamePictures(GetPic(PicName)).NextFrame.Present UI, Control.Left, Control.top
End Sub
Sub DrawRectangleRectUI(UI As Long, X As Single, Y As Single, W As Single, H As Single, BackColor As Long)
'Exit Sub
If LastPenColor <> BackColor Then
LastPenColor = BackColor
GdipSetSolidFillColor Brush1, BackColor
End If
GdipFillRectangle UI, Brush1, X, Y, W, H
End Sub
Sub DrawRectangleControlUI(UI As Long, Control As Object, BackColor As Long)
'Exit Sub
If Control.Visible = False Then Exit Sub
If LastPenColor <> BackColor Then
LastPenColor = BackColor
GdipSetSolidFillColor Brush1, BackColor
End If
GdipFillRectangle UI, Brush1, Control.Left, Control.top, Control.Width, Control.Height
End Sub
Sub DrawTextControlUI(UI As Long, Control As Object, ByVal Text As String, ForeColor As Long, mode As StringAlignment, Optional BigSize As Boolean = False)
If Control.Visible = False Then Exit Sub
If LastPenColor <> ForeColor Then
LastPenColor = ForeColor
GdipSetSolidFillColor Brush1, ForeColor
End If
GdipSetStringFormatAlign strformat, mode
If BigSize = False Then
GdipDrawString UI, StrPtr(Text), -1, curFont, NewRectF(Control.Left, Control.top, Control.Width, Control.Height), strformat, Brush1
Else
GdipDrawString UI, StrPtr(Text), -1, curFontBig, NewRectF(Control.Left, Control.top, Control.Width, Control.Height), strformat, Brush1
End If
End Sub
Sub DrawTextRectUI(UI As Long, ByVal X As Single, ByVal Y As Single, ByVal Text As String, ForeColor As Long, mode As StringAlignment, Optional BigSize As Boolean = False)
If LastPenColor <> ForeColor Then
LastPenColor = ForeColor
GdipSetSolidFillColor Brush1, ForeColor
End If
GdipSetStringFormatAlign strformat, mode
If BigSize = False Then
GdipDrawString UI, StrPtr(Text), -1, curFont, NewRectF(X, Y, 0, 0), strformat, Brush1
Else
GdipDrawString UI, StrPtr(Text), -1, curFontBig, NewRectF(X, Y, 0, 0), strformat, Brush1
End If
End Sub


