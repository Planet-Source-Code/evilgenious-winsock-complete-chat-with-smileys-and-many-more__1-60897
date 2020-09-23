VERSION 5.00
Begin VB.Form frmNick 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   88
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   239
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MynetChat.chameleonButton cmdSetNick 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Set Nick"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmNick.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtnickright 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtnick 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.PictureBox windowborder 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      Picture         =   "frmNick.frx":001C
      ScaleHeight     =   420
      ScaleWidth      =   3615
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         ToolTipText     =   "Close"
         Top             =   90
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "chameleonButton1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmNick.frx":517A
         PICN            =   "frmNick.frx":5196
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   -1  'True
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Change your nick"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmNick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'declare constants:
Private Const HWND_TOPMOST = -1

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

'declare API:
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, _
  ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
  ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
  
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302
  
'Round the form
Dim rndfrm As New ROUND_FORM





Private Sub cmdClose_Click()
Me.Visible = False
End Sub

Private Sub cmdSetNick_Click()
Client.NewNick = txtnick & txtnickright
End Sub

Private Sub Form_Load()
'set on top
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or _
SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'round form shape
rndfrm.ROUND_FORM Me, 12, 1, 1
End Sub


Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
End Sub

Private Sub WindowBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
End Sub

