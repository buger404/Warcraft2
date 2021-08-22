VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Square war II server"
   ClientHeight    =   8565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   8565
   ScaleWidth      =   6135
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   6105
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ä§ÊÞ»ìÕ½2·þÎñÆ÷"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00262626&
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   90
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8100
      Left            =   15
      TabIndex        =   0
      Top             =   450
      Width           =   6105
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00F0B000&
      Height          =   8565
      Left            =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1
  
'Private Const GWL_STYLE = (-16)
  
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
  
Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
           Dim ReturnVal As Long
           X = ReleaseCapture()
           ReturnVal = SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub
Private Sub Form_Load()
Form1.Show
SetParent Form1.hWnd, Frame1.hWnd
Form1.Move 0, 0
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Frame2_MouseDown(Button, Shift, X, Y)
End Sub
