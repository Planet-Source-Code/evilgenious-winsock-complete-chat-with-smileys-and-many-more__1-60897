VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFonts 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "Change Fonts"
   ClientHeight    =   4515
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
   Icon            =   "frmFonts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox windowborder 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      Picture         =   "frmFonts.frx":000C
      ScaleHeight     =   420
      ScaleWidth      =   11535
      TabIndex        =   27
      Top             =   0
      Width           =   11535
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   7080
         TabIndex        =   29
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
         MICON           =   "frmFonts.frx":AB96
         PICN            =   "frmFonts.frx":ABB2
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
         Caption         =   "Change Fonts ..."
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
         Index           =   2
         Left            =   240
         TabIndex        =   28
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   40
      Picture         =   "frmFonts.frx":B078
      ScaleHeight     =   255
      ScaleWidth      =   7470
      TabIndex        =   25
      Top             =   480
      Width           =   7470
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Color Scheme"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3120
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MynetChat.MyButton MyButton1 
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      ToolTipText     =   "Close this window"
      Top             =   4080
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      SPN             =   "MyButtonDefSkin"
      Text            =   "Close"
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
   Begin VB.PictureBox lblstatuscolor 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H00B78828&
      Height          =   200
      Left            =   5640
      ScaleHeight     =   165
      ScaleWidth      =   1365
      TabIndex        =   22
      ToolTipText     =   "Chat background color"
      Top             =   3600
      Width           =   1395
   End
   Begin VB.PictureBox lblclientconncolor 
      Appearance      =   0  'Flat
      BackColor       =   &H001784D5&
      ForeColor       =   &H001784D5&
      Height          =   200
      Left            =   5640
      ScaleHeight     =   165
      ScaleWidth      =   1365
      TabIndex        =   20
      ToolTipText     =   "Chat background color"
      Top             =   3120
      Width           =   1395
   End
   Begin VB.PictureBox usrlstfcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00B78828&
      ForeColor       =   &H00B78828&
      Height          =   200
      Left            =   5640
      ScaleHeight     =   165
      ScaleWidth      =   1365
      TabIndex        =   18
      ToolTipText     =   "Chat background color"
      Top             =   2640
      Width           =   1395
   End
   Begin VB.PictureBox usrlstbcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      ForeColor       =   &H00815438&
      Height          =   200
      Left            =   5640
      ScaleHeight     =   165
      ScaleWidth      =   1365
      TabIndex        =   16
      ToolTipText     =   "Chat background color"
      Top             =   2160
      Width           =   1395
   End
   Begin VB.PictureBox boxcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00B00009&
      ForeColor       =   &H00815438&
      Height          =   200
      Left            =   5640
      ScaleHeight     =   165
      ScaleWidth      =   1365
      TabIndex        =   14
      ToolTipText     =   "Chat background color"
      Top             =   1680
      Width           =   1395
   End
   Begin VB.PictureBox topiccolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00815438&
      ForeColor       =   &H00815438&
      Height          =   200
      Left            =   5640
      ScaleHeight     =   165
      ScaleWidth      =   1365
      TabIndex        =   12
      ToolTipText     =   "Chat background color"
      Top             =   1200
      Width           =   1395
   End
   Begin MynetChat.MyButton cmdmw 
      Height          =   375
      Left            =   720
      TabIndex        =   9
      ToolTipText     =   "Set the fonts when finished."
      Top             =   3960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      SPN             =   "MyButtonDefSkin"
      Text            =   "Set"
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
      Left            =   2880
      Picture         =   "frmFonts.frx":14BB2
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.TextBox txtmwsample 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Text            =   "   ABCXYZabcxyz"
      Top             =   3375
      Width           =   1785
   End
   Begin VB.ComboBox cmbmwstyle 
      Height          =   315
      ItemData        =   "frmFonts.frx":17108
      Left            =   720
      List            =   "frmFonts.frx":17118
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Font style"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox cmbmwsize 
      Height          =   315
      ItemData        =   "frmFonts.frx":17140
      Left            =   720
      List            =   "frmFonts.frx":1714D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Font size"
      Top             =   2040
      Width           =   855
   End
   Begin VB.ComboBox cmbmwfont 
      Height          =   315
      Left            =   720
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select the font"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Control Colors"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   24
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Window Fonts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   23
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Color:"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   9
      Left            =   4320
      TabIndex        =   21
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Online User Color:"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   8
      Left            =   4320
      TabIndex        =   19
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User List F-Color:"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   6
      Left            =   4320
      TabIndex        =   17
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User List B-Color:"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   15
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Box Color:"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Topic Color:"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      Top             =   3270
      Width           =   1770
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sample:"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   71
      X2              =   71
      Y1              =   216
      Y2              =   248
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   495
      Left            =   720
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Style:"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Font:"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

'Form Object
Dim FrmObject As Form

Dim x_IsItalic As Boolean
Dim x_IsBold As Boolean


Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
End Sub

Private Sub MyButton1_Click()
Me.Hide
End Sub

Private Sub WindowBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
End Sub



Private Sub cmdmw_Click()
SET_FONT FormObjectHandle
End Sub

Private Sub Form_Load()
FillComboWithFonts cmbmwfont
'set window on top
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or _
SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'round form shape
rndfrm.ROUND_FORM Me, 12, 1, 1
End Sub


'For main window
Private Sub cmbmwfont_Click()
txtmwsample.Font = cmbmwfont.Text
End Sub
Private Sub cmbmwsize_Click()
txtmwsample.FontSize = cmbmwsize.Text
End Sub
Private Sub cmbmwstyle_Click()
If cmbmwstyle.Text = "Regular" Then
    txtmwsample.FontBold = False
    txtmwsample.FontItalic = False
    x_IsBold = False
    x_IsItalic = False
ElseIf cmbmwstyle.Text = "Italic" Then
    txtmwsample.FontBold = False
    txtmwsample.FontItalic = True
    x_IsBold = False
    x_IsItalic = True
ElseIf cmbmwstyle.Text = "Bold" Then
    txtmwsample.FontBold = True
    txtmwsample.FontItalic = False
    x_IsBold = True
    x_IsItalic = False
ElseIf cmbmwstyle.Text = "Bold Italic" Then
    txtmwsample.FontBold = True
    txtmwsample.FontItalic = True
    x_IsBold = True
    x_IsItalic = True
End If
End Sub


Public Function SET_FONT(frmobj As Form)
On Error Resume Next
'set main window fonts
frmobj.txtchat.SelFontName = cmbmwfont.Text
frmobj.txtchat.SelFontSize = cmbmwsize
If cmbmwstyle.Text = "Regular" Then
    frmobj.txtchat.SelBold = False
    frmobj.txtchat.SelItalic = False
ElseIf cmbmwstyle.Text = "Italic" Then
    frmobj.txtchat.SelBold = False
    frmobj.txtchat.SelItalic = True
ElseIf cmbmwstyle.Text = "Bold" Then
    frmobj.txtchat.SelBold = True
    frmobj.txtchat.SelItalic = False
ElseIf cmbmwstyle.Text = "Bold Italic" Then
    frmobj.txtchat.SelBold = True
    frmobj.txtchat.SelItalic = True
End If
frmobj.FontNameIs = cmbmwfont.Text
frmobj.IsItalic = x_IsItalic
frmobj.IsBold = x_IsBold
End Function

Public Property Get FormObjectHandle() As Form
Set FormObjectHandle = FrmObject
End Property

Public Property Let FormObjectHandle(ByVal vNewValue As Form)
Set FrmObject = vNewValue
End Property
