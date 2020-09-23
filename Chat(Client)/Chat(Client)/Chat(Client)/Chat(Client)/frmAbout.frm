VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MynetChat.chameleonButton cmdClose 
      Height          =   375
      Left            =   3855
      TabIndex        =   10
      Top             =   1395
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmAbout.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1200
      Picture         =   "frmAbout.frx":001C
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   9
      Top             =   7080
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox windowborder 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      Picture         =   "frmAbout.frx":2572
      ScaleHeight     =   420
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "About Mynet Chat ..."
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
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Irteza Khan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Special thanks to : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":93D0
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   105
      Left            =   360
      Picture         =   "frmAbout.frx":A212
      Top             =   1200
      Width           =   4350
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright ©"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Hackers"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Underworld"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "hat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFCBCB&
      Height          =   375
      Index           =   1
      Left            =   1860
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF9999&
      Height          =   615
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ynet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CC6666&
      Height          =   375
      Index           =   0
      Left            =   990
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00993333&
      Height          =   615
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "frmAbout"
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
Unload Me
End Sub

Private Sub Form_Load()
'set on top
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or _
SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'round form shape
rndfrm.ROUND_FORM Me, 12, 1, 1
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
