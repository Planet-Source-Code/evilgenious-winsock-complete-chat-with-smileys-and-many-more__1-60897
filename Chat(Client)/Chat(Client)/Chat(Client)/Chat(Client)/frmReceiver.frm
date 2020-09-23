VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReceiver 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "Receiver Form"
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   178
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   960
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MynetChat.BinaryReceiver Reader 
      Left            =   4320
      Top             =   4320
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin VB.TextBox txtTarget 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3600
      Width           =   3615
   End
   Begin MSWinsockLib.Winsock udprequest 
      Left            =   4080
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1200
      TabIndex        =   25
      Top             =   480
      Width           =   4335
      Begin VB.TextBox txtTotalFileSize 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBF5F4&
         BorderStyle     =   0  'None
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0"
         Top             =   0
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
         Index           =   5
         Left            =   1560
         TabIndex        =   27
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.PictureBox WindowBorder 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      Picture         =   "frmReceiver.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   6975
      TabIndex        =   24
      Top             =   0
      Width           =   6975
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   6480
         TabIndex        =   30
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
         MICON           =   "frmReceiver.frx":A2EA
         PICN            =   "frmReceiver.frx":A306
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
         Caption         =   "FTP Receiver ..."
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
         Index           =   8
         Left            =   360
         TabIndex        =   31
         Top             =   120
         Width           =   1575
      End
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
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      ToolTipText     =   "Status log"
      Top             =   1920
      Width           =   6375
   End
   Begin VB.Frame famLog 
      Caption         =   "Event Log"
      Height          =   2175
      Left            =   8040
      TabIndex        =   20
      Top             =   3840
      Width           =   5655
   End
   Begin VB.Frame famSendStatus 
      Caption         =   "Receive File Status"
      Height          =   1095
      Left            =   8040
      TabIndex        =   16
      Top             =   2760
      Width           =   5655
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bytes Received:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblTotalByteReceived 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblPercentage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage: "
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame famReceiveOption 
      Caption         =   "Send File Option"
      Enabled         =   0   'False
      Height          =   1575
      Left            =   8040
      TabIndex        =   13
      Top             =   1200
      Width           =   5655
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total File Size(in bytes):"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "The location to save the file:"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame famRemoteInfo 
      Caption         =   "Remote Receiver Information"
      Height          =   1215
      Left            =   8040
      TabIndex        =   6
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtLocalPortBinary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtLocalPortInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtLocalIP 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Local Port(Binary):"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Local Port(Info):"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Local Host IP:"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.TextBox txtSpeed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   1320
      Width           =   1335
   End
   Begin MynetChat.MyButton cmdSetTarget 
      Height          =   255
      Left            =   5880
      TabIndex        =   1
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      SPN             =   "MyButtonDefSkin"
      Text            =   "Save to"
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
      Left            =   1320
      Picture         =   "frmReceiver.frx":A7CC
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   0
      Top             =   5520
      Width           =   2250
   End
   Begin VB.Timer tmrStatus 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   2760
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   13080
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Save to:"
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
      Index           =   0
      Left            =   1200
      TabIndex        =   29
      Top             =   3600
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
      Index           =   7
      Left            =   240
      TabIndex        =   23
      Top             =   1680
      Width           =   1095
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
      Index           =   6
      Left            =   240
      TabIndex        =   21
      Top             =   720
      Width           =   855
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
      Left            =   240
      TabIndex        =   5
      Top             =   1320
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
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
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
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmReceiver"
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

Private Sub cmdClose_Click()
If pb.Value < 2 Then
    Me.Hide
    Reader.Reset
    Reader.Listen
    FLAG = False
    udprequest.SendData "RESET"
    DoEvents
End If
End Sub

Private Sub cmdSetTarget_Click()
With cd1
    
    .CancelError = True
    .Filter = "All File *.*|*.*"
    .Flags = cdlOFNOverwritePrompt
    .Filename = Reader.CurrentFileName
    
    On Error GoTo OpenError
    .ShowSave
    
    txtTarget.Text = .Filename
    
    tmrStatus.Enabled = True
    Reader.AcceptSendRequest txtTarget.Text
    
    AddLog "Accepted Send Request at " & Time
    AddLog "Saving file to [" & txtTarget.Text & "]"
    
    cmdSetTarget.Enabled = False
    
End With

Exit Sub
OpenError:
cmdSetTarget.Enabled = True
End Sub

Private Sub Form_Load()

'Set the port for the two data protocol
Reader.LocalPortBinary = 3000
Reader.LocalPortInfo = 1700

Reader.Listen

txtLocalIP.Text = Reader.LocalIP
txtLocalPortInfo.Text = Reader.LocalPortInfo
txtLocalPortBinary.Text = Reader.LocalPortBinary

'round form shape
rndfrm.ROUND_FORM Me, 12, 1, 1

'start udpclose listening for close message
udprequest.Bind 5000

End Sub

Private Sub Reader_ConnectionRequest()
AddLog "Connection Request from [" & Reader.TheWinsock.RemoteHostIP & "]"
End Sub

Private Sub Reader_ReceiveComplete()
tmrStatus.Enabled = False
AddLog "File Receive Complete at " & Time
pb.Value = 0
Reader.ResetFile
mProgress = 0
mProgressMax = 0
mPercentage = 0
lblPercentage.Caption = "100%"
End Sub

Private Sub Reader_ReceiveProgress(ByVal Progress As Long, ByVal ProgressMax As Long)

mProgress = Progress
mProgressMax = ProgressMax

End Sub

Private Sub Reader_SendRequest()

AddLog "Send Request from [" & Reader.TheWinsock.RemoteHostIP & "]"
AddLog "File Name= " & Reader.CurrentFileName
AddLog "File Size(in bytes)= " & Reader.CurrentFileSize

Dim m As Integer
m = MsgBox("A file sending request from " & Reader.TheWinsock.RemoteHostIP & " is made. Do you want to accept it?", vbInformation + vbYesNo)

If m = vbYes Then
    
    txtTotalFileSize.Text = Reader.CurrentFileSize
    
    famReceiveOption.Enabled = True
    
    With cd1
        
        .CancelError = True
        .Filter = "All File *.*|*.*"
        .Flags = cdlOFNOverwritePrompt
        .Filename = Reader.CurrentFileName
        
        On Error GoTo OpenError
        .ShowSave
        
        txtTarget.Text = .Filename
        
        tmrStatus.Enabled = True
        Reader.AcceptSendRequest txtTarget.Text
        
        AddLog "Accepted Send Request at " & Time
        AddLog "Saving file to [" & txtTarget.Text & "]"
        
        cmdSetTarget.Enabled = False
        
    End With

Else
    
    Reader.RefuseSendRequest
    Reader.ResetFile
    
End If

Exit Sub
OpenError:
cmdSetTarget.Enabled = True
End Sub

Sub AddLog(str As String)
txtEventLog.Text = txtEventLog.Text & str & vbCrLf
txtEventLog.SelStart = Len(txtEventLog.Text)
End Sub

Private Sub Reader_SpeedRecord(ByVal Speed As Long)
txtSpeed.Text = Speed
End Sub

Private Sub tmrStatus_Timer()
On Error Resume Next
pb.Max = mProgressMax
pb.Value = mProgress
lblTotalByteReceived.Caption = mProgress
lblPercentage.Caption = (Int(mProgress / mProgressMax * 100) + 1) & "%"
End Sub


Private Sub udprequest_DataArrival(ByVal bytesTotal As Long)
Dim Msg As String
udprequest.GetData Msg
If Right(Msg, 4) = ":FTR" And FLAG = False Then
    MESSAGE = MsgBox(Left(Msg, Len(Msg) - 4) & " wants to send a file to you", vbYesNo, "File sending request ...")
    If MESSAGE = vbYes Then
        FLAG = True
        frmReceiver.Visible = True
        frmReceiver.txtEventLog.Text = ""
        udprequest.SendData "YES"
        DoEvents
    ElseIf MESSAGE = vbNo Then
        udprequest.SendData "NO"
        DoEvents
    End If
ElseIf FLAG = True Then
    udprequest.SendData "NO"
    DoEvents
End If
End Sub


Private Sub WindowBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
End Sub
