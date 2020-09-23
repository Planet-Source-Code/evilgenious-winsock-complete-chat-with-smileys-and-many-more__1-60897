VERSION 5.00
Begin VB.Form frmEmot 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "Emot Check"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   Icon            =   "frmEmot.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   127
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtemotno 
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
      Height          =   285
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Enter any number"
      Top             =   1530
      Width           =   1230
   End
   Begin VB.PictureBox imageemot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   1200
      TabIndex        =   1
      ToolTipText     =   "Watch smiley here"
      Top             =   525
      Width           =   1230
   End
   Begin VB.PictureBox WindowBorder 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      Picture         =   "frmEmot.frx":000C
      ScaleHeight     =   420
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   0
      Width           =   1500
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   1140
         TabIndex        =   3
         ToolTipText     =   "Close"
         Top             =   90
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
         MICON           =   "frmEmot.frx":224A
         PICN            =   "frmEmot.frx":2266
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
         Caption         =   "Smileys ..."
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
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmEmot"
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
  
  

Public Function DISPLAY_PICTURE(EMOTICON As Integer)

'This function will paste the picture  in the Richtextbox at position = pos
'emoticon defines the picture number to be pasted
imageemot.Picture = LoadPicture(App.Path & "\emoticons\" & EMOTICON & ".gif")
    
End Function

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

Private Sub txtemotno_Change()
On Error GoTo ERRDESCRIPTION
If txtemotno.Text <> "" Then
    DISPLAY_PICTURE CInt(txtemotno.Text)
    Exit Sub
End If

ERRDESCRIPTION:
imageemot.Cls
End Sub

Private Sub txtemotno_KeyPress(KeyAscii As Integer)
If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    Client.txtmessage.SelText = ":" & txtemotno.Text
End If
End Sub

Private Sub WindowBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
End Sub
