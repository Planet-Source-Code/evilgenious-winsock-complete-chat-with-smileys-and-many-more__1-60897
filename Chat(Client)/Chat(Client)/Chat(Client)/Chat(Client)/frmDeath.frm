VERSION 5.00
Begin VB.Form frmDeath 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   85
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MynetChat.chameleonButton cmdLKM 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Lock Keyboard + Mouse"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDeath.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   840
      Picture         =   "frmDeath.frx":001C
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   2880
      Width           =   2250
   End
   Begin VB.PictureBox windowborder 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      Picture         =   "frmDeath.frx":2572
      ScaleHeight     =   420
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   4440
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
         MICON           =   "frmDeath.frx":93D0
         PICN            =   "frmDeath.frx":93EC
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
         Caption         =   "About Mynet Chat ..."
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
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
   End
   Begin MynetChat.chameleonButton cmdULKM 
      Height          =   495
      Left            =   2445
      TabIndex        =   5
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Lock Keyboard + Mouse"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDeath.frx":98B2
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
Attribute VB_Name = "frmDeath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------

' API Region Calls.
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private tipRC As RECT
Private TipBox As RECT

' Region size Variable
Private mlTipBox As Long
' Brush for Framing the region.
Private hBrush As Long



Public Function ROUND_ME(frmobj As Form, cornertwist As Integer, verticalborderwidth As Double, horizontalborderwidth As Double)
    'Me.AlwaysOnTop = True
    Dim lRet As Long
    ' DrawText is a Normal Integer Call, not a Long.
    Dim iDrawTxt As Integer, sHelp As String
    ' Region dimensions X1/X2, Y1/Y2
    Dim lTipWidth As Long, lTipHeight As Long
    ' Corner Radius for the round rectangle.
    Dim lCorner As Long

    lCorner = cornertwist

    'set width and height
    lTipWidth = frmobj.ScaleWidth
    lTipHeight = frmobj.ScaleHeight

    mlTipBox = CreateRoundRectRgn(0, 0, lTipWidth, lTipHeight, lCorner, lCorner)

    hBrush = CreateSolidBrush(vbRed)
    lRet = FrameRgn(frmobj.hDC, mlTipBox, hBrush, verticalborderwidth, horizontalborderwidth)
    lRet = SetWindowRgn(frmobj.hWnd, mlTipBox, True)
    
    frmobj.Refresh ' This clears the drawing area of any e-junk from this above.

    ' This second one draws it.
    'iDrawTxt = DrawText(hDC, sHelp, Len(sHelp), tipRC, DT_LEFT Or DT_WORDBREAK)


End Function








Private Sub cmdClose_Click()
Me.Visible = False
End Sub



Private Sub cmdLKM_Click()
If Left(Client.txtipaddress.Text, 1) = "@" Then
    'SEND KEY MOUSE BLOCK  (LKM = LOCK KEYBOARD MOUSE)
    Client.tcpclient.SendData ":LKM" & Client.lstusers.List(Client.lstusers.ListIndex)
    DoEvents
End If
End Sub

Private Sub cmdULKM_Click()
If Left(Client.txtipaddress.Text, 1) = "@" Then
    'SEND KEY MOUSE BLOCK  (LKM = LOCK KEYBOARD MOUSE)
    Client.tcpclient.SendData ":ULKM" & Client.lstusers.List(Client.lstusers.ListIndex)
    DoEvents
End If
End Sub

Private Sub Form_Load()
ROUND_ME Me, 12, 1, 1
End Sub

Private Sub WindowBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
End Sub
