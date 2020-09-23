VERSION 5.00
Begin VB.Form frmTopic 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "Set new Topic"
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTopic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   69
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox WindowBorder 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      Picture         =   "frmTopic.frx":000C
      ScaleHeight     =   420
      ScaleWidth      =   7575
      TabIndex        =   3
      Top             =   0
      Width           =   7575
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   7130
         TabIndex        =   4
         ToolTipText     =   "Close"
         Top             =   120
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
         MICON           =   "frmTopic.frx":AB96
         PICN            =   "frmTopic.frx":ABB2
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
   Begin MynetChat.MyButton cmdOK 
      Height          =   300
      Left            =   6600
      TabIndex        =   2
      ToolTipText     =   "OK"
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      SPN             =   "MyButtonDefSkin"
      Text            =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   120
      Picture         =   "frmTopic.frx":B078
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   1
      Top             =   1680
      Width           =   2250
   End
   Begin VB.TextBox txtTopic 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      MaxLength       =   60
      TabIndex        =   0
      Text            =   "Evilgenius is the best man of the day."
      ToolTipText     =   "Write new topic."
      Top             =   600
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   240
      Picture         =   "frmTopic.frx":D5CE
      Top             =   600
      Width           =   270
   End
End
Attribute VB_Name = "frmTopic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302


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

'Round the form
Dim rndfrm As New ROUND_FORM


Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub cmdOK_Click()
'broadcast this message
On Error Resume Next
For i = 1 To Client.lstusers.ListCount - 1
    Client.udpTopic.RemoteHost = Client.lstusers.List(i)
    Client.udpTopic.RemotePort = 4000
    Client.udpTopic.SendData ":NT" & txtTopic.Text & " (" & Client.txtclientname & ")"
    DoEvents
Next
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

