VERSION 5.00
Begin VB.Form AeroWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H002D282A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   484
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   743
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame ToolFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   3000
      TabIndex        =   1
      Top             =   1650
      Width           =   4815
   End
   Begin VB.Label litter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ã«²£Á§±³¾°"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   4950
      TabIndex        =   0
      Top             =   3450
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "AeroWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


