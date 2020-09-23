VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMassMessage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "Mass Message"
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   Icon            =   "frmMassMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   40
      Picture         =   "frmMassMessage.frx":0BC2
      ScaleHeight     =   255
      ScaleWidth      =   7470
      TabIndex        =   8
      Top             =   480
      Width           =   7470
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the users from the right list."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox WindowBorder 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      Picture         =   "frmMassMessage.frx":A6FC
      ScaleHeight     =   420
      ScaleWidth      =   7575
      TabIndex        =   6
      Top             =   0
      Width           =   7575
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   7130
         TabIndex        =   7
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
         MICON           =   "frmMassMessage.frx":15286
         PICN            =   "frmMassMessage.frx":152A2
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
         Caption         =   "Send Mass Message to all users ..."
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
         TabIndex        =   10
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.TextBox txtmassmessage 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin MynetChat.MyButton cmdSend 
      Height          =   465
      Left            =   5880
      TabIndex        =   1
      Top             =   2130
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   820
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
      Picture         =   "frmMassMessage.frx":15768
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   240
      Picture         =   "frmMassMessage.frx":15B7E
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   0
      Top             =   3600
      Width           =   2250
   End
   Begin MynetChat.MyButton cmdCancel 
      Height          =   465
      Left            =   4320
      TabIndex        =   2
      Top             =   2130
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   820
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
      Picture         =   "frmMassMessage.frx":180D4
   End
   Begin MSComctlLib.ListView lstusers 
      Height          =   945
      Left            =   4680
      TabIndex        =   4
      Top             =   960
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   1667
      View            =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MynetChat.MyButton MyButton1 
      Height          =   1185
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2090
      SPN             =   "MyButtonDefSkin"
      Text            =   "MyButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisableHover    =   -1  'True
   End
End
Attribute VB_Name = "frmMassMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302

'Round the form
Dim rndfrm As New ROUND_FORM

Dim localclientname As String


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSend_Click()

'run loop to send all the clients
For i = 1 To lstusers.ListItems.Count
    If lstusers.ListItems(i).Checked = True Then
        Client.clientpmsend.RemoteHost = lstusers.ListItems(i)
        Client.clientpmsend.RemotePort = 8000
        Client.clientpmsend.SendData localclientname & ":" & txtmassmessage.Text
        DoEvents
    End If
Next
Unload Me

End Sub


Private Sub Form_Load()

localclientname = Client.txtclientname.Text

For i = 1 To Client.lstusers.ListCount - 1
    lstusers.ListItems.Add i, , Client.lstusers.List(i)
Next

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
