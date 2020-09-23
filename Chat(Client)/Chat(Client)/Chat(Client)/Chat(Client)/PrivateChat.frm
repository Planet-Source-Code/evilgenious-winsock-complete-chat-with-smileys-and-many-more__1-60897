VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PrivateChat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "PRIVATE CHAT"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   Icon            =   "PrivateChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   90
      Picture         =   "PrivateChat.frx":08CA
      ScaleHeight     =   255
      ScaleWidth      =   6690
      TabIndex        =   19
      Top             =   435
      Width           =   6690
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Private Chat"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   2160
   End
   Begin VB.PictureBox namecolor 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   5760
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   12
      ToolTipText     =   "chat name color"
      Top             =   3420
      Width           =   200
   End
   Begin VB.PictureBox backgroundcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   5400
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   11
      ToolTipText     =   "Chat background color"
      Top             =   3420
      Width           =   200
   End
   Begin VB.PictureBox hypercolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   6480
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   10
      ToolTipText     =   "URL color"
      Top             =   3420
      Width           =   200
   End
   Begin VB.PictureBox messagecolor 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   6120
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   9
      ToolTipText     =   "chat message color"
      Top             =   3420
      Width           =   200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   2160
   End
   Begin MynetChat.MyButton cmdsend 
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   3780
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   661
      SPN             =   "MyButtonDefSkin"
      Text            =   "Send"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Left            =   960
      Picture         =   "PrivateChat.frx":7140
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   5
      Top             =   6120
      Width           =   2250
   End
   Begin VB.PictureBox WindowBorder 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      Picture         =   "PrivateChat.frx":9696
      ScaleHeight     =   420
      ScaleWidth      =   6975
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   6480
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
         MICON           =   "PrivateChat.frx":1332C
         PICN            =   "PrivateChat.frx":13348
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   -1  'True
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   120
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
         BackColor       =   &H00EBF5F4&
         BackStyle       =   0  'Transparent
         Caption         =   "Smdanyal"
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
         Height          =   315
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Chatter name"
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1680
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock privateclient 
      Left            =   2160
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin RichTextLib.RichTextBox txtchat 
      Height          =   2595
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4577
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"PrivateChat.frx":1380E
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
   Begin RichTextLib.RichTextBox txtmessage 
      Height          =   330
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Type your message here"
      Top             =   3795
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"PrivateChat.frx":13889
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MynetChat.chameleonButton cmdPrivateChat 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Nudge"
      Top             =   3375
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   ""
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   15463924
      BCOLO           =   15463924
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "PrivateChat.frx":13906
      PICN            =   "PrivateChat.frx":13922
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton cmdFont 
      Height          =   330
      Left            =   3960
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Change fonts"
      Top             =   3375
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      BTYPE           =   8
      TX              =   ""
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   15463924
      BCOLO           =   15463924
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "PrivateChat.frx":13D8A
      PICN            =   "PrivateChat.frx":13DA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton cmdClear 
      Height          =   330
      Left            =   4320
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Clear chat"
      Top             =   3375
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      BTYPE           =   8
      TX              =   ""
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   15463924
      BCOLO           =   15463924
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "PrivateChat.frx":142A8
      PICN            =   "PrivateChat.frx":142C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton cmdGroup 
      Height          =   330
      Left            =   4680
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Hackers Group Info"
      Top             =   3375
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      BTYPE           =   8
      TX              =   ""
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   15463924
      BCOLO           =   15463924
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "PrivateChat.frx":1474E
      PICN            =   "PrivateChat.frx":1476A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.MyButton cmdBold 
      Height          =   345
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Bold"
      Top             =   3420
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   609
      SPN             =   "MyButtonDefSkin"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PrivateChat.frx":14A0A
      PicturePos      =   4
   End
   Begin MynetChat.MyButton cmdItalic 
      Height          =   345
      Left            =   600
      TabIndex        =   17
      ToolTipText     =   "Italic"
      Top             =   3420
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   609
      SPN             =   "MyButtonDefSkin"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PrivateChat.frx":14E38
      PicturePos      =   4
   End
   Begin MynetChat.MyButton cmdUnderline 
      Height          =   345
      Left            =   960
      TabIndex        =   18
      ToolTipText     =   "Underline"
      Top             =   3420
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   609
      SPN             =   "MyButtonDefSkin"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PrivateChat.frx":14ED2
      PicturePos      =   4
   End
   Begin MynetChat.chameleonButton cmdemoticons 
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Smileys"
      Top             =   3375
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   ""
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   15463924
      BCOLO           =   15463924
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "PrivateChat.frx":14F74
      PICN            =   "PrivateChat.frx":14F90
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton cmdFtp 
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "File Transfer"
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Send File"
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   15463924
      BCOLO           =   15463924
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "PrivateChat.frx":15456
      PICN            =   "PrivateChat.frx":15472
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00EBF5F4&
      Index           =   1
      X1              =   96
      X2              =   6
      Y1              =   252
      Y2              =   252
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   450
      X2              =   258
      Y1              =   223
      Y2              =   223
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00BAB6B3&
      Height          =   360
      Left            =   105
      Top             =   3780
      Width           =   5715
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00BAB6B3&
      Height          =   2655
      Left            =   105
      Top             =   705
      Width           =   6675
   End
   Begin VB.Label lblbytestransferred 
      Caption         =   "Label3"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00BAB6B3&
      BorderColor     =   &H00BAB6B3&
      Height          =   390
      Index           =   0
      Left            =   3885
      Top             =   3345
      Width           =   2895
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00BAB6B3&
      BorderColor     =   &H00BAB6B3&
      Height          =   405
      Index           =   1
      Left            =   105
      Top             =   3390
      Width           =   1335
   End
End
Attribute VB_Name = "PrivateChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302

'Round the form
Dim rndfrm As New ROUND_FORM

'color variables
Dim name_color, message_color, hyper_color, back_color As Ole_Color

Dim start As Integer
Dim newstart As Integer

Dim MESSAGE As String
Dim nickis As String
Dim countnudge As Integer

Dim i As Integer
Dim temp As Integer
Dim emotfind As Integer
Dim G_startfrom  As Long

Dim x_fontname As String
Dim x_IsItalic As Boolean
Dim x_IsBold As Boolean

Private Sub cmdBold_Click()
txtmessage.Font.Bold = Not txtmessage.Font.Bold
End Sub

Private Sub cmdClear_Click()
txtchat.Text = ""
newstart = 0
G_startfrom = 0
End Sub

Private Sub cmdClose_Click()
'CPC = CLOSE PRVATE CHAT
privateclient.SendData "CPC:"
DoEvents
End Sub

Private Sub cmdemoticons_Click()
frmEmot.Show
End Sub

Private Sub cmdFont_Click()
frmFonts.Show
frmFonts.FormObjectHandle = Me
End Sub

Private Sub cmdFtp_Click()
frmSender.Visible = True
frmSender.txtRemoteHost.Clear
'fill combo with usernames
Dim i As Integer
For i = 1 To Client.lstusers.ListCount - 1
    frmSender.txtRemoteHost.AddItem Client.lstusers.List(i)
Next
frmSender.txtRemoteHost.AddItem " "
End Sub

Private Sub cmdGroup_Click()
GROUP.Show 1
End Sub

Private Sub cmdItalic_Click()
txtmessage.Font.Italic = Not txtmessage.Font.Italic
End Sub

Private Sub cmdPrivateChat_Click()
privateclient.SendData ":NUDGE"
DoEvents
End Sub

Private Sub cmdUnderline_Click()
txtmessage.Font.Underline = Not txtmessage.Font.Underline
End Sub

Private Sub Form_Load()
txtmessage.SelStart = Len(txtmessage.Text)
'round form shape
rndfrm.ROUND_FORM Me, 12, 1, 1

G_startfrom = 1
txtchat.SelColor = vbBlack
name_color = namecolor.BackColor
message_color = messagecolor.BackColor
hyper_color = hypercolor.BackColor

End Sub

Private Sub Form_Unload(Cancel As Integer)
'CPC = CLOSE PRVATE CHAT
privateclient.SendData "CPC:"
DoEvents
End Sub

Private Sub privateclient_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
'SAVE MESSAGE BUFFER IN MESSAGE
privateclient.GetData MESSAGE

If MESSAGE = "CPC:" Then
    Unload Me
    Exit Sub
End If
If MESSAGE = ":NUDGE" Then
    PLAY_SOUND "nudge"
    NUDGE_ME
    Timer1.Enabled = True
    Exit Sub
'IF IT IS A MESSAGE FROM SERVER
Else
    txtchat.SelFontName = FontNameIs
    txtchat.SelItalic = IsItalic
    txtchat.SelBold = IsBold
    txtchat.SelUnderline = False
    txtchat.SelStart = Len(txtchat.Text)
    newstart = Len(txtchat.Text)
    txtchat.SelText = MESSAGE
    COLORTEXT
    'set the selstart
    txtchat.SelStart = Len(txtchat.Text)
    txtchat.SelText = vbCrLf
    
    newstart = Len(txtchat.Text)
End If

MESSAGE = ""

End Sub

Private Sub cmdSend_Click()
On Error Resume Next
If Len(txtmessage.Text) < Len(nick & " >> ") + 1 Then
    Exit Sub
End If
'This is because the text in RichTextBox  is in RTF format and you can't write txtchat.Text = txtmessage.Text, as it will only copy the text from the source not the graphics.
'The RichTextBox format is originally in TextRTF format so whenever you are sending RichTextBox contents to other RichTextBox then always write
'RichTextBox1.TextRTF = RichTextBox2.TextRTF


'Disable Timer
Timer1.Enabled = False

'Set SelStart = 0 to copy the text from start
txtmessage.SelStart = 0
'Set lenght upto the length of txtmessage
txtmessage.SelLength = Len(txtmessage.Text)

'Set SelStart = length of txtchat
txtchat.SelStart = Len(txtchat.Text)
txtchat.SelUnderline = False

'copy the contents to txtchat
newstart = Len(txtchat.Text)
txtchat.SelStart = newstart

txtchat.SelFontName = FontNameIs
txtchat.SelItalic = IsItalic
txtchat.SelBold = IsBold
txtchat.SelText = txtmessage.Text & vbCrLf

'send data immediately to server
privateclient.SendData txtmessage.Text
DoEvents

'color the text in txtchat
COLORTEXT
'clear and set the start of typing in txtmessage
txtmessage.Text = ""
txtmessage.SelStart = Len(txtmessage.Text)

'restore factory defaults
txtmessage.SelColor = vbBlack
txtmessage.SelUnderline = False
End Sub

Private Sub Timer1_Timer()
NUDGE_ME
Timer1.Enabled = False
End Sub

Private Sub Timer4_Timer()
CONVERT_INTO_SMILEYS G_startfrom
End Sub

Private Sub txtchat_Change()
Timer4.Enabled = True
End Sub

Private Sub txtmessage_Change()
If Len(txtmessage.Text) < Len(nick & " >> ") Then
    txtmessage.Text = nick & " >> "
    txtmessage.SelStart = Len(txtmessage.Text)
    Exit Sub
End If
End Sub

Private Sub txtmessage_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSend_Click
End If
End Sub

Public Property Get nick() As String
nick = nickis
End Property

Public Property Let nick(ByVal newnick As String)
nickis = newnick
End Property

Private Sub WindowBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
End Sub


Private Sub hypercolor_Click()
cd.ShowColor
hypercolor.BackColor = cd.Color
hyper_color = hypercolor.BackColor
End Sub

Private Sub messagecolor_Click()
cd.ShowColor
messagecolor.BackColor = cd.Color
message_color = messagecolor.BackColor
End Sub

Private Sub namecolor_Click()
cd.ShowColor
namecolor.BackColor = cd.Color
name_color = namecolor.BackColor
End Sub

Private Sub backgroundcolor_Click()
cd.ShowColor
txtchat.BackColor = cd.Color
backgroundcolor.BackColor = cd.Color
End Sub

































Public Function COLORTEXT()

On Error Resume Next

Dim colorupto As Integer
Dim newstart2 As Integer

newstart2 = newstart

'color the whole message first
txtchat.SelStart = newstart
txtchat.SelLength = Len(txtchat.Text) - newstart
txtchat.SelColor = message_color

'color the name
colorupto = txtchat.Find(">>", newstart, Len(txtchat.Text))
If colorupto > -1 Then
    txtchat.SelStart = newstart
    txtchat.SelLength = colorupto - newstart + 2
    txtchat.SelColor = name_color

    'color the message
    newstart = colorupto + 2
    colorupto = Len(txtchat.Text)
    txtchat.SelStart = newstart
    txtchat.SelLength = colorupto - newstart
    txtchat.SelColor = message_color
    txtchat.SelStart = Len(txtchat.Text)
End If

'color the hyperlink
Dim colorstart, colorend As Integer

If txtchat.Find("http", newstart2, Len(txtchat.Text)) > -1 Then
    colorstart = txtchat.Find("http", newstart2, Len(txtchat.Text))
ElseIf txtchat.Find("www.", newstart2, Len(txtchat.Text)) > -1 Then
    colorstart = txtchat.Find("www.", newstart2, Len(txtchat.Text))
End If

If colorstart = 0 Then
    Exit Function
End If

newstart2 = colorstart
colorend = txtchat.Find(" ", newstart2, Len(txtchat.Text))
If colorstart > -1 Then
    txtchat.SelStart = colorstart
    If colorend = -1 Then
        txtchat.SelLength = Len(txtchat.Text) - colorstart
        txtchat.SelColor = hyper_color
        txtchat.SelUnderline = True
    Else
        txtchat.SelLength = colorend - colorstart
        txtchat.SelColor = hyper_color
        txtchat.SelUnderline = True
    End If
End If


End Function


Public Function CONVERT_INTO_SMILEYS(STARTFROM As Long)

Dim x_length As Long
Dim x_foundat As Long
Dim x_locset As Long
Dim x_EMOT As String
Dim x_i As Integer
Dim x_j As Integer

x_length = Len(txtchat.Text)

'run the loop upto the length of richtextbox
For x_i = STARTFROM To x_length
    x_foundat = txtchat.Find(":", x_i, x_length)
    x_locset = x_foundat
    
    If x_foundat = -1 Then
        x_i = x_length
        Timer4.Enabled = False
        Exit Function
    ElseIf x_foundat > -1 And IsNumeric(Mid(txtchat.Text, x_foundat + 2, 1)) = True Then
        x_foundat = x_foundat + 2
        'loop for calculating x_EMOT number
        For x_j = 1 To 4
            If IsNumeric(Mid(txtchat.Text, x_foundat, 1)) = True Then
                x_EMOT = x_EMOT & Mid(txtchat.Text, x_foundat, 1)
                x_foundat = x_foundat + 1
            Else
                Exit For
            End If
        Next
        'convert number into smiley
        If x_EMOT <> "" Then
            SET_PICTURE txtchat, x_locset, CInt(x_EMOT), Len(x_EMOT) + 1
            x_i = x_foundat - Len(x_EMOT) - 2
        Else
            x_i = x_foundat - 1
        End If
        
        x_EMOT = ""
    End If
    
    G_startfrom = x_length
    txtchat.SelStart = x_length
    Timer4.Enabled = False
Next

End Function


Public Function SET_PICTURE(rt_box As RichTextBox, pos As Long, EMOTICON As Integer, length As Integer)

'This function will paste the picture  in the Richtextbox at position = pos
'emoticon defines the picture number to be pasted
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetData LoadPicture(App.Path & "\emoticons\" & EMOTICON & ".gif"), vbCFBitmap
    rt_box.SelStart = pos
   'Replace the text   :) or :( or :| with empty string
    rt_box.SelLength = length
    rt_box.SelText = ""
    ' Paste the picture into the RichTextBox.
    SendMessage rt_box.hWnd, WM_PASTE, 0, 0
    
End Function



Public Function PLAY_SOUND(Filename As String)
sndPlaySound App.Path & "\" & Filename, SND_ASYNC Or SND_NODEFAULT
End Function

Public Sub NUDGE_ME()
Me.Top = Me.Top + 100
Me.Left = Me.Left + 200
Me.Top = Me.Top - 200
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left + 200
Me.Top = Me.Top - 200
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left + 200
Me.Top = Me.Top - 200
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left + 200
Me.Top = Me.Top - 200
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left + 200
Me.Top = Me.Top - 200
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left + 200
Me.Top = Me.Top - 200
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left + 200
Me.Top = Me.Top - 200
Me.Left = Me.Left - 100
Me.Top = Me.Top + 100
Me.Left = Me.Left - 100


End Sub

Public Property Get FontNameIs() As String
FontNameIs = x_fontname
End Property

Public Property Let FontNameIs(ByVal vNewValue As String)
x_fontname = vNewValue
End Property

Public Property Get IsItalic() As Boolean
IsItalic = x_IsItalic
End Property

Public Property Let IsItalic(ByVal vNewValue As Boolean)
x_IsItalic = vNewValue
End Property

Public Property Get IsBold() As Boolean
IsBold = x_IsBold
End Property

Public Property Let IsBold(ByVal vNewValue As Boolean)
x_IsBold = vNewValue
End Property
