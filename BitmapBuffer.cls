VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BlendFunction As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal HDC As Long, pBitmapInfo As BITMAPINFOHEADER, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

 Const GMEM_MOVEABLE As Long = &H2
 Const AC_SRC_Alpha = &H1
 Const AC_SRC_Over = &H0
 Const DIB_RGB_Colors As Long = 0


Private Type BlendFunction
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type size
 cx As Long
 cy As Long
End Type

Private Type Rect
        Left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type

Const ULW_Alpha = &H2
Const GWL_EXSTYLE As Long = -20
Const WS_EX_LAYERED = &H80000

Public CompatibleDC As Long
Public Graphics As Long
Private DIB As Long
Private mWidth As Long, mHeight As Long, BlendFunc As BlendFunction

Public Property Get Height() As Long 'DC�߶�
Height = mHeight
End Property

Public Property Get Width() As Long 'DC����
Width = mWidth
End Property
 '����һ����GDI��ͼ�õĿ���ݻ���
Public Sub BeginDraw()
    GdipCreateFromHDC CompatibleDC, Graphics
    GdipSetSmoothingMode Graphics, SmoothingModeAntiAlias '���������
End Sub
    '���ٻ���
Public Sub EndDraw()
    GdipDeleteGraphics Graphics
End Sub
 '��ʼ����Hdc���Դ��봰�ڵ�Hdc��width/heightΪ������DC�Ĵ�С
Public Sub Create(HDC As Long, Width As Long, Height As Long)
    Dim bmHeader As BITMAPINFOHEADER, BlendFunc As BlendFunction

    mWidth = Width
    mHeight = Height

    With BlendFunc
    .AlphaFormat = AC_SRC_Alpha
    .BlendFlags = 0
    .BlendOp = AC_SRC_Over
    .SourceConstantAlpha = 255
    End With

    With bmHeader
    .biSize = Len(bmHeader)
    .biBitCount = 32
    .biHeight = Height
    .biWidth = Width
    .biPlanes = 1
    .biSizeImage = .biWidth * .biHeight * 4
    End With

    CompatibleDC = CreateCompatibleDC(HDC)
    DIB = CreateDIBSection(CompatibleDC, bmHeader, DIB_RGB_Colors, ByVal 0, 0, 0)
    DeleteObject SelectObject(CompatibleDC, DIB)

End Sub

'��DC���ͼ�����ź�����Ŀ��DC
'DstDC:Ŀ��DC
'W/H:���ź�Ĵ�С
'Alpha:���ʱ������͸����
Public Sub PresentWithWH(ByVal DstDC As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional Alpha As Byte = 255)
    Dim BlendFunLong As Long

    With BlendFunc
     .BlendOp = AC_SRC_Over
     .AlphaFormat = AC_SRC_Alpha
     .BlendFlags = 0
     .SourceConstantAlpha = Alpha
    End With

    RtlMoveMemory BlendFunLong, BlendFunc, 4
    AlphaBlend DstDC, X, Y, W, H, CompatibleDC, 0, 0, mWidth, mHeight, BlendFunLong

End Sub

'��DC���ͼ����ú�����Ŀ��DC
'DstDC:Ŀ��DC
'X/Y:������Ŀ��DC������
'OffsetX/OffsetY������ƫ������
'height/width�����ô�С
'Alpha:���ʱ������͸����
Public Sub PresentWithClip(ByVal DstDC As Long, X As Long, Y As Long, offsetX As Long, offsetY As Long, Width As Long, Height As Long, Optional Alpha As Byte = 255)
    Dim SrcSize As Rect, BlendFunLong As Long

    If offsetX + Width > mWidth Then Width = mWidth - offsetX
    If offsetY + Height > mHeight Then Height = mHeight - offsetY

    SrcSize.Right = Width
    SrcSize.Bottom = Height

    With BlendFunc
     .BlendOp = AC_SRC_Over
     .AlphaFormat = AC_SRC_Alpha
     .BlendFlags = 0
     .SourceConstantAlpha = Alpha
    End With

    RtlMoveMemory BlendFunLong, BlendFunc, 4
    AlphaBlend DstDC, X, Y, SrcSize.Right, SrcSize.Bottom, CompatibleDC, offsetX, offsetY, SrcSize.Right, SrcSize.Bottom, BlendFunLong

End Sub

'��DC�������ͼ��ֱ������Ŀ��DC
'DstDC:Ŀ��DC
'X/Y:������Ŀ��DC������
'Alpha:���ʱ������͸����
Public Sub Present(ByVal DstDC As Long, ByVal X As Long, ByVal Y As Long, Optional Alpha As Byte = 255)
    Dim SrcSize As Rect, BlendFunLong As Long

    SrcSize.Right = mWidth
    SrcSize.Bottom = mHeight

    With BlendFunc
     .BlendOp = AC_SRC_Over
     .AlphaFormat = AC_SRC_Alpha
     .BlendFlags = 0
     .SourceConstantAlpha = Alpha
    End With

    RtlMoveMemory BlendFunLong, BlendFunc, 4
    AlphaBlend DstDC, X, Y, SrcSize.Right, SrcSize.Bottom, CompatibleDC, SrcSize.Left, SrcSize.top, SrcSize.Right, SrcSize.Bottom, BlendFunLong

End Sub

Sub Delete()
    If Graphics <> 0 Then GdipDeleteGraphics Graphics '��ֹ��������
    DeleteObject CompatibleDC
    DeleteObject DIB
End Sub

