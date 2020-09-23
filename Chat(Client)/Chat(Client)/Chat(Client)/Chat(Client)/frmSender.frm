VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSender 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "Sender Form"
   ClientHeight    =   3060
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MynetChat.BinarySender Sender 
      Left            =   1560
      Top             =   4800
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   2415
      Left            =   120
      TabIndex        =   21
      Top             =   480
      Width           =   6735
      Begin VB.TextBox txtSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00EBF5F4&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtEventLog 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   615
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         ToolTipText     =   "Status log"
         Top             =   1680
         Width           =   6375
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EBF5F4&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Left            =   1080
         TabIndex        =   22
         Top             =   135
         Width           =   3735
         Begin VB.TextBox txtSource 
            Appearance      =   0  'Flat
            BackColor       =   &H00EBF5F4&
            BorderStyle     =   0  'None
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   0
            Width           =   3375
         End
         Begin VB.TextBox txtTotalFileSize 
            Appearance      =   0  'Flat
            BackColor       =   &H00EBF5F4&
            BorderStyle     =   0  'None
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "0"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "KB"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   225
            Index           =   7
            Left            =   1560
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
      End
      Begin MynetChat.MyButton cmdBrowse 
         Height          =   240
         Left            =   5640
         TabIndex        =   28
         ToolTipText     =   "Open the file"
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   423
         SPN             =   "MyButtonDefSkin"
         Text            =   "Browse"
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
      Begin MynetChat.MyButton cmdSend 
         Height          =   255
         Left            =   5640
         TabIndex        =   29
         ToolTipText     =   "Send the file"
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
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
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Filesize:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Filename:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "KB/sec"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   4
         Left            =   2280
         TabIndex        =   32
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Progress:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Events Log:"
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
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.PictureBox WindowBorder 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      Picture         =   "frmSender.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   6975
      TabIndex        =   18
      Top             =   0
      Width           =   6975
      Begin MynetChat.chameleonButton cmdRequest 
         Height          =   300
         Left            =   4320
         TabIndex        =   36
         ToolTipText     =   "Send the request and wait."
         Top             =   75
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         BTYPE           =   4
         TX              =   "Send file transfer request"
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
         BCOL            =   16576
         BCOLO           =   16576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSender.frx":9C96
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox txtRemoteHost 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   20
         ToolTipText     =   "Select the user"
         Top             =   60
         Width           =   1815
      End
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   6480
         TabIndex        =   37
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
         MICON           =   "frmSender.frx":9CB2
         PICN            =   "frmSender.frx":9CCE
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
         BackColor       =   &H0099007B&
         BackStyle       =   0  'Transparent
         Caption         =   "FTP Sender ..."
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
         Index           =   9
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Users:"
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
         Height          =   225
         Index           =   0
         Left            =   1800
         TabIndex        =   19
         Top             =   120
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   240
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1560
      Picture         =   "frmSender.frx":A194
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   16
      Top             =   6240
      Width           =   2250
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5040
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame famRemoteInfo 
      Caption         =   "Remote Receiver Information"
      Height          =   1215
      Left            =   7320
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtRemotePortBinary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtRemotePortInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect to remote receiver"
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Port(Binary):"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Port(Info):"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame famSendOption 
      Caption         =   "Send File Option"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   7320
      TabIndex        =   0
      Top             =   1200
      Width           =   5655
      Begin VB.TextBox txtChunkSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2400
         TabIndex        =   17
         Text            =   "4096"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Chunk Size(in bytes):"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total File Size(in bytes):"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "The file in local machine you want to send:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame famSendStatus 
      Caption         =   "Send File Status"
      Height          =   1095
      Left            =   7320
      TabIndex        =   11
      Top             =   3000
      Width           =   5655
      Begin VB.Timer tmrStatus 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   240
         Top             =   360
      End
      Begin VB.Label lblPercentage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblTotalByteSent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sending Speed(KBps):"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bytes Sent:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame famLog 
      Caption         =   "Event Log"
      Height          =   1935
      Left            =   7320
      TabIndex        =   7
      Top             =   4080
      Width           =   5655
   End
   Begin MSWinsockLib.Winsock udprequest 
      Left            =   2760
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302

'Round the form
Dim rndfrm As New ROUND_FORM


Dim mProgress As Long
Dim mProgressMax As Long
Dim mPercentage As Long

Dim FLAG As Boolean


Private Sub cmdBrowse_Click()
With cd1

    .Filter = "All Files *.*|*.*"
    .CancelError = True
    .Flags = cdlOFNFileMustExist
    
    On Error GoTo 1
    .ShowOpen
    
    txtSource.Text = .Filename
    
    cmdSend.Enabled = True
    
End With
Exit Sub
1
End Sub

Private Sub cmdClose_Click()
If pb.Value < 2 Then
    Me.Hide
    Sender.Reset
    FLAG = False
End If
End Sub

Private Sub cmdRequest_Click()
If txtRemoteHost.Text <> "" Then
    FLAG = True
    udprequest.RemoteHost = txtRemoteHost.Text
    udprequest.RemotePort = 5000
    udprequest.SendData Client.txtclientname & ":FTR"
    DoEvents
Else
    MsgBox "Please select the user from the list to whom you want to send the file.", vbInformation, "Information"
End If
End Sub

Private Sub cmdSend_Click()

cmdSend.Enabled = False

Sender.ChunkSize = CLng(txtChunkSize.Text)
Sender.Source = txtSource.Text

txtTotalFileSize.Text = Sender.CurrentFileSize

Sender.SendInfo

AddLog "File Information Sent to " & Sender.RemoteHost
AddLog "File Name= " & Sender.CurrentFileName
AddLog "File Size(in bytes)= " & Sender.CurrentFileSize

AddLog "Waiting for receiver[" & Sender.RemoteHostIP & "] ready signal..."

End Sub

Private Sub Form_Load()
Me.Show
Me.Move 0, 0

'round form shape
rndfrm.ROUND_FORM Me, 12, 1, 1

'fill combo with usernames
For i = 1 To Client.lstusers.ListCount - 1
    txtRemoteHost.AddItem Client.lstusers.List(i)
Next

End Sub


Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
End Sub

Private Sub Sender_CommandAccepted()
AddLog "Sending command accepted by [" & Sender.RemoteHostIP & "]"
AddLog "File Sending Started at " & Time
AddLog "Now sending file [" & Sender.Source & "]"
tmrStatus.Enabled = True
Sender.SendFile
End Sub

Private Sub Sender_CommandRefused()
AddLog "Sending command refused by [" & Sender.RemoteHostIP & "]"

Sender.ResetFile
mProgress = 0
mProgressMax = 0
mPercentage = 0

pb.Value = 0
cmdConnect.Enabled = False
cmdSend.Enabled = True
famSendOption.Enabled = True

End Sub

Private Sub Sender_Connect()
AddLog "<Sender Control> connected to [" & Sender.RemoteHost & "] successfully."
txtRemotePortInfo.Text = Sender.RemotePortInfo
txtRemotePortBinary.Text = Sender.RemotePortBinary

famSendOption.Enabled = True

End Sub

Private Sub Sender_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbCritical, Number
If Number = 10049 Then cmdConnect.Enabled = True
End Sub

Private Sub Sender_SendComplete()
tmrStatus.Enabled = False
AddLog "File Sending Complete at " & Time

Sender.ResetFile
mProgress = 0
mProgressMax = 0
mPercentage = 0

pb.Value = 0
lblPercentage.Caption = "100%"
cmdConnect.Enabled = False
cmdSend.Enabled = True
famSendOption.Enabled = True

End Sub

Private Sub Sender_SendProgress(ByVal Progress As Long, ByVal ProgressMax As Long)
mProgress = Progress
mProgressMax = ProgressMax
End Sub

Sub AddLog(str As String)
txtEventLog.Text = txtEventLog.Text & str & vbCrLf
txtEventLog.SelStart = Len(txtEventLog.Text)
End Sub

Private Sub Sender_SpeedRecord(ByVal Speed As Long)
txtSpeed.Text = Speed
End Sub

Private Sub tmrStatus_Timer()
pb.Max = mProgressMax
pb.Value = mProgress
lblTotalByteSent.Caption = mProgress
lblPercentage.Caption = (Int(mProgress / mProgressMax * 100) + 1) & "%"
End Sub

Private Sub udprequest_DataArrival(ByVal bytesTotal As Long)
Dim Msg As String
udprequest.GetData Msg
If Msg = "YES" And FLAG = True Then
    Me.Show
    Sender.RemoteHost = txtRemoteHost.Text
    Sender.Connect
    Frame1.Enabled = True
ElseIf Msg = "NO" Then
    MsgBox txtRemoteHost.Text & " declined to accept the file.", vbCritical, "FTP declined ..."
ElseIf Msg = "RESET" Then
    Sender.Reset
    Frame1.Enabled = False
    txtRemoteHost.Text = " "
    txtEventLog.Text = ""
    FLAG = False
End If
End Sub


Private Sub WindowBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
End Sub
