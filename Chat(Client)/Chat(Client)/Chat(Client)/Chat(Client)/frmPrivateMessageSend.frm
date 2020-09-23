VERSION 5.00
Begin VB.Form frmPrivateMessageSend 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "Send Message"
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrivateMessageSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox WindowBorder 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      Picture         =   "frmPrivateMessageSend.frx":0BC2
      ScaleHeight     =   420
      ScaleWidth      =   6975
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   6480
         TabIndex        =   1
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
         MICON           =   "frmPrivateMessageSend.frx":A858
         PICN            =   "frmPrivateMessageSend.frx":A874
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Private Message ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      Picture         =   "frmPrivateMessageSend.frx":AD3A
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   3
      Top             =   4200
      Width           =   2250
   End
   Begin VB.TextBox txtprivatemessage 
      Appearance      =   0  'Flat
      Height          =   1300
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   945
      Width           =   6200
   End
   Begin MynetChat.MyButton cmdSend 
      Height          =   495
      Left            =   5230
      TabIndex        =   4
      Top             =   2475
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      SPN             =   "MyButtonDefSkin"
      Text            =   "Send"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPrivateMessageSend.frx":D290
   End
   Begin MynetChat.MyButton cmdCancel 
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   2475
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      SPN             =   "MyButtonDefSkin"
      Text            =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPrivateMessageSend.frx":D6A6
   End
   Begin MynetChat.MyButton MyButton1 
      Height          =   1575
      Left            =   240
      TabIndex        =   6
      Top             =   810
      Width           =   6445
      _ExtentX        =   11377
      _ExtentY        =   2778
      SPN             =   "MyButtonDefSkin"
      Text            =   "MyButton1"
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
   Begin VB.Label lblname 
      BackStyle       =   0  'Transparent
      Caption         =   "evilgenious"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   525
      Width           =   5535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Send to:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   240
      TabIndex        =   7
      Top             =   525
      Width           =   1335
   End
End
Attribute VB_Name = "frmPrivateMessageSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302

'Round the form
Dim rndfrm As New ROUND_FORM

Dim localclientname As String


Private Sub cmdClear_Click()
txtprivatemessage.Text = ""
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSend_Click()
On Error Resume Next
Client.clientpmsend.RemoteHost = lblname.Caption
Client.clientpmsend.RemotePort = 8000
Client.clientpmsend.SendData localclientname & ":" & txtprivatemessage.Text
DoEvents
Unload Me
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Client.lstusers.ListIndex > 0 Then
    lblname.Caption = Client.lstusers.List(Client.lstusers.ListIndex)
    localclientname = Client.txtclientname.Text
Else
    cmdSend.Enabled = False
    lblname.Caption = "You did not select any user."
End If

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
