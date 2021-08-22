VERSION 5.00
Begin VB.Form LOGO 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "LOGO"
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   Icon            =   "LOGO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   266
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
End
Attribute VB_Name = "LOGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WS_EX_TRANSPARENT   As Long = &H20&
Private Const HTCAPTION = 2
Private Const ArcSize = 25
Private Const ArcAngle = 90
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1


Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, ByVal pptDst As Long, ByRef psize As size, ByVal hdcSrc As Long, pptSrc As Currency, ByVal crKey As Long, ByRef pBlend As BlendFunction, ByVal dwFlags As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal HDC As Long, pBitmapInfo As BITMAPINFOHEADER, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

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

Const ULW_Alpha = &H2
Const AC_SRC_Alpha = &H1
Const AC_SRC_Over = &H0
Const GWL_EXSTYLE As Long = -20
Const WS_EX_LAYERED = &H80000
Const DIB_RGB_Colors As Long = 0
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Sub Form_Load()
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H1 Or &H2

Dim FormWidth As Long, FormHeight As Long
Dim mColor As Variant
Dim RectLBrush As Long, Graphic As Long, Path As Long, mToken As Long, mGSI As GdiplusStartupInput
Dim ArcRP(11, 1) As Single
Dim CompatibleDC As Long, DIB As Long, bmHeader As BITMAPINFOHEADER
Dim WinSz As size, BlendFunc As BlendFunction

    FormWidth = Me.ScaleWidth
    FormHeight = Me.ScaleHeight
    
    WinSz.cx = FormWidth
    WinSz.cy = FormHeight
    


    With bmHeader
    .biSize = Len(bmHeader)
    .biBitCount = 32
    .biHeight = FormHeight
    .biWidth = FormWidth
    .biPlanes = 1
    .biSizeImage = .biWidth * .biHeight * 4
    End With
 
    With BlendFunc
    .AlphaFormat = AC_SRC_Alpha
    .BlendFlags = 0
    .BlendOp = AC_SRC_Over
    .SourceConstantAlpha = 255
    End With
 
    CompatibleDC = CreateCompatibleDC(Me.HDC)
    DIB = CreateDIBSection(Me.HDC, bmHeader, DIB_RGB_Colors, ByVal 0, 0, 0)
    DeleteObject SelectObject(CompatibleDC, DIB)
    
    mGSI.GdiplusVersion = "1.0"
    GdiplusStartup mToken, mGSI
    GdipCreateFromHDC CompatibleDC, Graphic
    
    Dim Img As Long
    
    GdipLoadImageFromFile StrPtr(App.Path + "\assets\LOGO.png"), Img
    GdipDrawImage Graphic, Img, 0, 0
    GdipDisposeImage Img
    GdipDeleteGraphics Graphic
    
    GdiplusShutdown mToken

    SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED    'Or WS_EX_TRANSPARENT
    UpdateLayeredWindow Me.hWnd, 0, 0, WinSz, CompatibleDC, 0, 0, BlendFunc, ULW_Alpha

    DeleteObject CompatibleDC
    DeleteObject DIB

End Sub


