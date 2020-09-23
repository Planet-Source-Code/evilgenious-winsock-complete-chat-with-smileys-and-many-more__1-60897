VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUpdateInfo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "frmUpdateInfo"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   LinkTopic       =   "Form1"
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox WindowBorder 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      Picture         =   "frmUpdateInfo.frx":0000
      ScaleHeight     =   420
      ScaleWidth      =   1575
      TabIndex        =   1
      Top             =   0
      Width           =   1575
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Close"
         Top             =   60
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
         MICON           =   "frmUpdateInfo.frx":223E
         PICN            =   "frmUpdateInfo.frx":225A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4095
      Left            =   30
      TabIndex        =   0
      Top             =   480
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      ForeColor       =   0
      BackColorFixed  =   15463924
      ForeColorFixed  =   0
      BackColorSel    =   0
      BackColorBkg    =   15463924
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "frmUpdateInfo"
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
Me.Hide
End Sub

Private Sub Form_Load()
'set window on top
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or _
SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'round form shape
rndfrm.ROUND_FORM Me, 12, 1, 1
End Sub



Private Sub WindowBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
End Sub
