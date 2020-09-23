VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2DA548C8-DE1F-4A93-9392-7FF8380971C9}#3.0#0"; "BinaryTransferControl.ocx"
Begin VB.Form frmReceiver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receiver Form"
   ClientHeight    =   6585
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   5880
   Begin VB.Timer tmrStatus 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   360
      Top             =   3120
   End
   Begin VB.Frame famRemoteInfo 
      Caption         =   "Remote Receiver Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtLocalIP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtLocalPortInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtLocalPortBinary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Local Host IP:"
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Local Port(Info):"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Local Port(Binary):"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5160
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin BinaryTransferControl.BinaryReceiver Reader 
      Left            =   4680
      Top             =   6120
      _ExtentX        =   794
      _ExtentY        =   794
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
   Begin VB.Frame famReceiveOption 
      Caption         =   "Send File Option"
      Enabled         =   0   'False
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   5655
      Begin VB.CommandButton cmdSetTarget 
         Caption         =   "Set Save &Target"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtTarget 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txtTotalFileSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "The location to save the file:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total File Size(in bytes):"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2295
      End
   End
   Begin VB.Frame famSendStatus 
      Caption         =   "Receive File Status"
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   5655
      Begin VB.TextBox txtSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblPercentage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage: "
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   480
         Width           =   1815
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
         TabIndex        =   20
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bytes Received:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Download Speed(KBps):"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Frame famLog 
      Caption         =   "Event Log"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   5655
      Begin VB.TextBox txtEventLog 
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mProgress As Long
Dim mProgressMax As Long
Dim mPercentage As Long

Private Sub cmdSetTarget_Click()
With cd1
    
    .CancelError = True
    .Filter = "All File *.*|*.*"
    .Flags = cdlOFNOverwritePrompt
    .FileName = Reader.CurrentFileName
    
    On Error GoTo OpenError
    .ShowSave
    
    txtTarget.Text = .FileName
    
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
        .FileName = Reader.CurrentFileName
        
        On Error GoTo OpenError
        .ShowSave
        
        txtTarget.Text = .FileName
        
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
