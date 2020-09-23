VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2DA548C8-DE1F-4A93-9392-7FF8380971C9}#3.0#0"; "BinaryTransferControl.ocx"
Begin VB.Form frmSender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sender Form"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
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
   ScaleHeight     =   6615
   ScaleWidth      =   5880
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5040
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin BinaryTransferControl.BinarySender Sender 
      Left            =   4440
      Top             =   6120
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin VB.Frame famRemoteInfo 
      Caption         =   "Remote Receiver Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtRemotePortBinary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtRemotePortInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtRemoteHost 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect to remote receiver"
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Port(Binary):"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Port(Info):"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame famSendOption 
      Caption         =   "Send File Option"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   5655
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send the file"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtChunkSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   18
         Text            =   "4096"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtTotalFileSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         Top             =   840
         Width           =   3135
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse..."
         Height          =   285
         Left            =   4560
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtSource 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Chunk Size(in bytes):"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total File Size(in bytes):"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "The file in local machine you want to send:"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4095
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame famSendStatus 
      Caption         =   "Send File Status"
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   5655
      Begin VB.Timer tmrStatus 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   240
         Top             =   360
      End
      Begin VB.TextBox txtSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblPercentage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage"
         Height          =   255
         Left            =   3120
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sending Speed(KBps):"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bytes Sent:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame famLog 
      Caption         =   "Event Log"
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   5655
      Begin VB.TextBox txtEventLog 
         Height          =   1575
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mProgress As Long
Dim mProgressMax As Long
Dim mPercentage As Long

Private Sub cmdBrowse_Click()
With cd1

    .Filter = "All Files *.*|*.*"
    .CancelError = True
    .Flags = cdlOFNFileMustExist
    
    On Error GoTo 1
    .ShowOpen
    
    txtSource.Text = .FileName
    
    cmdSend.Enabled = True
    
End With
Exit Sub
1
End Sub

Private Sub cmdConnect_Click()

cmdConnect.Enabled = False

Sender.RemoteHost = txtRemoteHost.Text

Sender.Connect

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

frmReceiver.Show
frmReceiver.Move frmSender.Width, 0

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
