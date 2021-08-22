VERSION 5.00
Begin VB.Form MsgWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   567
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   887
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3765
      Left            =   3600
      TabIndex        =   0
      Top             =   2100
      Width           =   6165
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   150
         TabIndex        =   5
         Text            =   "ÇëÊäÈë..."
         Top             =   2850
         Visible         =   0   'False
         Width           =   5925
      End
      Begin Ä§ÊÞ»ìÕ½2.UIButton UIButton2 
         Height          =   315
         Left            =   4050
         TabIndex        =   3
         Top             =   3300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   15773696
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Caption         =   "È·¶¨"
         BackColor2      =   15785984
      End
      Begin Ä§ÊÞ»ìÕ½2.UIButton UIButton3 
         Height          =   315
         Left            =   5100
         TabIndex        =   4
         Top             =   3300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   6999040
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Caption         =   "È¡Ïû"
         BackColor2      =   7003136
         BackColor3      =   7003136
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   150
         TabIndex        =   6
         Top             =   660
         Width           =   5925
         Begin VB.Label ScrollButton 
            BackColor       =   &H00F0B000&
            Height          =   255
            Left            =   5820
            TabIndex        =   7
            Top             =   0
            Width           =   105
         End
         Begin VB.Label ScrollBar 
            BackColor       =   &H00F2F2F2&
            Height          =   2415
            Left            =   5820
            TabIndex        =   8
            Top             =   0
            Width           =   105
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ÕâÀïÏÔÊ¾ÎÄ±¾"
            BeginProperty Font 
               Name            =   "Î¢ÈíÑÅºÚ"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00262626&
            Height          =   285
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1170
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E9E4E9&
         X1              =   150
         X2              =   6000
         Y1              =   3150
         Y2              =   3150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "±êÌâà¶"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F0B000&
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6165
      End
   End
End
Attribute VB_Name = "MsgWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ClickButton As Integer
Public StartY As Single
Private Sub Form_Load()
On Error Resume Next
Frame1.Move Me.ScaleWidth / 2 - Frame1.Width / 2, Me.ScaleHeight / 2 - Frame1.Height / 2
Dad.SetFocus
End Sub

Private Sub ScrollBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call scrollbutton_MouseDown(1, 0, 0, 0)
End Sub

Private Sub ScrollBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call scrollbutton_MouseMove(1, 0, 0, 0)
End Sub

Private Sub UIButton2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClickButton = 2
End Sub

Private Sub UIButton3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClickButton = 1
End Sub
Sub SetScroll()
scrollbutton.Visible = True
scrollbar.Visible = True
scrollbutton.top = scrollbar.top
If Label3.Height > Frame2.Height Then
scrollcount = Round((Label3.Height - Frame2.Height) / 285)
Else
scrollcount = 0
End If

If scrollcount <= 0 Then
scrollbutton.Height = scrollbar.Height
scrollbutton.Visible = False
scrollbar.Visible = False
Else
scrollbutton.Height = scrollbar.Height / scrollcount
End If
scrollbutton.Tag = scrollbutton.Height
If scrollbutton.Height < 45 Then scrollbutton.Height = 45
End Sub
Private Sub scrollbutton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartY = Y
End Sub

Public Sub scrollbutton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Dim p As POINTAPI
GetCursorPos p
p.Y = p.Y * 15 - StartY
tmpy = p.Y - Frame1.top * 15 - Frame2.top - Dad.top - 16 * 15
If tmpy < scrollbar.top Then tmpy = scrollbar.top
If tmpy > scrollbar.top + scrollbar.Height - scrollbutton.Height Then tmpy = scrollbar.top + scrollbar.Height - scrollbutton.Height
'tmpx = scrollbar.Left + Round((tmpx - scrollbar.Left) / scrollbutton.Width) * scrollbutton.Width
scrollbutton.top = tmpy
Label3.top = -((scrollbutton.top - scrollbar.top) / Val(scrollbutton.Tag) * 285)
End If
End Sub
