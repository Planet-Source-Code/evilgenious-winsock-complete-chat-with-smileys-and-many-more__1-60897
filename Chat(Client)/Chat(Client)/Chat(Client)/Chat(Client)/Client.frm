VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Client 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "EVIL CHAT VER.(X)"
   ClientHeight    =   8835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Client.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   589
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MynetChat.MyButton cmdDeath 
      Height          =   480
      Left            =   9285
      TabIndex        =   67
      Top             =   8160
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   847
      SPN             =   "MyButtonDefSkin"
      Text            =   "Death"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Client.frx":1E72
   End
   Begin MynetChat.chameleonButton cmdNick 
      Height          =   330
      Left            =   330
      TabIndex        =   66
      Top             =   7785
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "Nick"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Client.frx":328C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3000
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock udpNeutralNetworkListen 
      Left            =   960
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock udpNeutralNetworkSend 
      Left            =   1560
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock udpTopic 
      Left            =   4200
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MynetChat.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   4920
      TabIndex        =   62
      ToolTipText     =   "Update info."
      Top             =   7740
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "..."
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Client.frx":32A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton cmdUpdate 
      Height          =   375
      Left            =   2160
      TabIndex        =   61
      ToolTipText     =   "Get Updated"
      Top             =   7740
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Press this button to be Updated"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Client.frx":32C4
      PICN            =   "Client.frx":32E0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   15
      Picture         =   "Client.frx":33CB
      ScaleHeight     =   255
      ScaleWidth      =   10440
      TabIndex        =   59
      Top             =   450
      Width           =   10440
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "About ?"
         Height          =   160
         Left            =   9720
         TabIndex        =   65
         ToolTipText     =   "About Mynet Chat"
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Mynet Chat"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         ToolTipText     =   "This is Mynet Chat"
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox windowborder 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      Picture         =   "Client.frx":CF05
      ScaleHeight     =   420
      ScaleWidth      =   11535
      TabIndex        =   53
      Top             =   0
      Width           =   11535
      Begin MynetChat.chameleonButton cmdMinimize 
         Height          =   255
         Left            =   9480
         TabIndex        =   54
         ToolTipText     =   "Send to tray"
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
         MICON           =   "Client.frx":1B4F7
         PICN            =   "Client.frx":1B513
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   -1  'True
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MynetChat.chameleonButton chameleonButton4 
         Height          =   255
         Left            =   9780
         TabIndex        =   55
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
         MICON           =   "Client.frx":1B9D9
         PICN            =   "Client.frx":1B9F5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   -1  'True
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   10080
         TabIndex        =   56
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
         MICON           =   "Client.frx":1BEBB
         PICN            =   "Client.frx":1BED7
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
         BackStyle       =   0  'Transparent
         Caption         =   "Mynet Chat ... ver x-01        Admin: Waqas / Author: Evilgenius"
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
         Left            =   840
         TabIndex        =   57
         Top             =   120
         Width           =   5415
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   240
         Picture         =   "Client.frx":1C39D
         Top             =   0
         Width           =   480
      End
   End
   Begin MynetChat.MyButton cmdReset 
      Height          =   285
      Left            =   4200
      TabIndex        =   49
      Top             =   960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      SPN             =   "MyButtonDefSkin"
      Text            =   "Reset"
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
   Begin VB.TextBox txtport 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B00009&
      Height          =   250
      IMEMode         =   3  'DISABLE
      Left            =   3375
      TabIndex        =   52
      Text            =   "10000"
      ToolTipText     =   "Connection Port (Hidden)"
      Top             =   975
      Width           =   720
   End
   Begin VB.TextBox txtclientname2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B00009&
      Height          =   250
      Left            =   3375
      Locked          =   -1  'True
      TabIndex        =   51
      ToolTipText     =   "Nick (Currently Disabled)"
      Top             =   1335
      Width           =   1695
   End
   Begin VB.TextBox txtclientname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B00009&
      Height          =   250
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   50
      ToolTipText     =   "Your name (Static)"
      Top             =   1335
      Width           =   1695
   End
   Begin MynetChat.TrayArea TrayArea 
      Left            =   6000
      Top             =   4320
      _ExtentX        =   900
      _ExtentY        =   900
      ToolTip         =   "Mynet Chat"
   End
   Begin MynetChat.chameleonButton kck 
      Height          =   405
      Left            =   9930
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Kick (If you are Op)"
      Top             =   7620
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Client.frx":1D1DF
      PICN            =   "Client.frx":1D1FB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton pvtmsg 
      Height          =   405
      Left            =   9465
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   "Private Message"
      Top             =   7620
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Client.frx":1D6C8
      PICN            =   "Client.frx":1D6E4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton pvtchat 
      Height          =   405
      Left            =   9030
      TabIndex        =   46
      TabStop         =   0   'False
      ToolTipText     =   "Private Chat"
      Top             =   7620
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Client.frx":1DB76
      PICN            =   "Client.frx":1DB92
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton wrn 
      Height          =   405
      Left            =   8595
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "Warn User"
      Top             =   7620
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Client.frx":1E003
      PICN            =   "Client.frx":1E01F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton cmdupdclientlist 
      Height          =   405
      Left            =   8160
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "Update User List"
      Top             =   7620
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Client.frx":1E426
      PICN            =   "Client.frx":1E442
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.MyButton cmdBold 
      Height          =   345
      Left            =   1020
      TabIndex        =   41
      ToolTipText     =   "Bold"
      Top             =   7785
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
      Picture         =   "Client.frx":1E8FF
      PicturePos      =   4
   End
   Begin MynetChat.chameleonButton cmdPrivateChat 
      Height          =   375
      Left            =   240
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Private Chat"
      Top             =   1905
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
      MICON           =   "Client.frx":1ED2D
      PICN            =   "Client.frx":1ED49
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox messagecolor 
      Appearance      =   0  'Flat
      BackColor       =   &H0069A451&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   7440
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   22
      ToolTipText     =   "chat message color"
      Top             =   7800
      Width           =   200
   End
   Begin VB.PictureBox hypercolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   7800
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   21
      ToolTipText     =   "URL color"
      Top             =   7800
      Width           =   200
   End
   Begin VB.PictureBox backgroundcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   6720
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   20
      ToolTipText     =   "Chat background color"
      Top             =   7800
      Width           =   200
   End
   Begin VB.PictureBox namecolor 
      Appearance      =   0  'Flat
      BackColor       =   &H009A7B34&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   7080
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   19
      ToolTipText     =   "chat name color"
      Top             =   7800
      Width           =   200
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   4200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1080
      Top             =   4725
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   4725
   End
   Begin MSWinsockLib.Winsock udpupdate 
      Left            =   3720
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock clientpmsend 
      Left            =   4200
      Top             =   5085
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock clientpmlisten 
      Left            =   3720
      Top             =   5085
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock udpclientsend 
      Left            =   4680
      Top             =   5595
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock udpclientlisten 
      Left            =   4200
      Top             =   5595
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock tcpclient 
      Left            =   3720
      Top             =   5595
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtchat 
      Height          =   4620
      Left            =   270
      TabIndex        =   16
      Top             =   3045
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   8149
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Client.frx":1F75B
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
   Begin VB.TextBox txthostname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B00009&
      Height          =   250
      Left            =   1560
      TabIndex        =   7
      Text            =   "Localhost"
      ToolTipText     =   "Server Name"
      Top             =   960
      Width           =   1695
   End
   Begin VB.PictureBox Standard 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   360
      Picture         =   "Client.frx":1F7D8
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   2
      Top             =   9720
      Width           =   1200
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   360
      Picture         =   "Client.frx":2071A
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   1
      Top             =   9360
      Width           =   2250
   End
   Begin VB.PictureBox Skin1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   360
      Picture         =   "Client.frx":22C70
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   0
      Top             =   10020
      Width           =   2250
   End
   Begin MynetChat.MyButton cmddisconnect 
      Height          =   615
      Left            =   7800
      TabIndex        =   3
      ToolTipText     =   "Disconnect from server"
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      SPN             =   "MyButtonDefSkin"
      Text            =   "Disconnect"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Client.frx":231CD
   End
   Begin MynetChat.MyButton cmdconnect 
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      ToolTipText     =   "Connect to server"
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      SPN             =   "MyButtonDefSkin"
      Text            =   "Connect"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Client.frx":2597F
   End
   Begin MynetChat.MyButton cmdSend 
      Height          =   480
      Left            =   8160
      TabIndex        =   17
      ToolTipText     =   "Send the message"
      Top             =   8160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   847
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
      Picture         =   "Client.frx":27801
   End
   Begin RichTextLib.RichTextBox txtmessage 
      Height          =   450
      Left            =   255
      TabIndex        =   18
      ToolTipText     =   "Type the message here"
      Top             =   8190
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   794
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MaxLength       =   100
      Appearance      =   0
      TextRTF         =   $"Client.frx":29803
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
   Begin MynetChat.chameleonButton cmdemoticons 
      Height          =   375
      Left            =   1575
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Smileys"
      Top             =   1905
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
      MICON           =   "Client.frx":29880
      PICN            =   "Client.frx":2989C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton cmdmassmsg 
      Height          =   375
      Left            =   2280
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Mass Message"
      Top             =   1905
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
      MICON           =   "Client.frx":29D62
      PICN            =   "Client.frx":29D7E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton cmdtopic 
      Height          =   375
      Left            =   3000
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Change Topic"
      Top             =   1905
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
      MICON           =   "Client.frx":2A244
      PICN            =   "Client.frx":2A260
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton cmdSettings 
      Height          =   375
      Left            =   3720
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Settings"
      Top             =   1905
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
      MICON           =   "Client.frx":2A63E
      PICN            =   "Client.frx":2A65A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton cmdTransparent 
      Height          =   375
      Left            =   4440
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Make Transparent"
      Top             =   1905
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
      MICON           =   "Client.frx":2AAE4
      PICN            =   "Client.frx":2AB00
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton cmdOpaque 
      Height          =   375
      Left            =   4920
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Make Opaque"
      Top             =   1905
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
      MICON           =   "Client.frx":2AFB8
      PICN            =   "Client.frx":2AFD4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.chameleonButton cmdAbout 
      Height          =   375
      Left            =   5880
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "About Evil"
      Top             =   1905
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
      MICON           =   "Client.frx":2B49A
      PICN            =   "Client.frx":2B4B6
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
      Left            =   5760
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Change fonts"
      Top             =   7755
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
      MICON           =   "Client.frx":2B9B8
      PICN            =   "Client.frx":2B9D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MynetChat.MyButton cmdItalic 
      Height          =   345
      Left            =   1380
      TabIndex        =   42
      ToolTipText     =   "Italic"
      Top             =   7785
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
      Picture         =   "Client.frx":2BED6
      PicturePos      =   4
   End
   Begin MynetChat.MyButton cmdUnderline 
      Height          =   345
      Left            =   1740
      TabIndex        =   43
      ToolTipText     =   "Underline"
      Top             =   7785
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
      Picture         =   "Client.frx":2BF70
      PicturePos      =   4
   End
   Begin MynetChat.chameleonButton cmdClear 
      Height          =   330
      Left            =   6120
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "Clear chat"
      Top             =   7755
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
      MICON           =   "Client.frx":2C012
      PICN            =   "Client.frx":2C02E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox moto 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   800
      Left            =   360
      Picture         =   "Client.frx":2C4B8
      ScaleHeight     =   795
      ScaleWidth      =   975
      TabIndex        =   5
      ToolTipText     =   "Status Emoticon"
      Top             =   795
      Width           =   975
   End
   Begin VB.PictureBox moto2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   360
      Picture         =   "Client.frx":2ED4A
      ScaleHeight     =   870
      ScaleWidth      =   975
      TabIndex        =   6
      Top             =   810
      Width           =   975
   End
   Begin VB.ListBox lstusers 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B78828&
      Height          =   4515
      Left            =   8160
      TabIndex        =   58
      ToolTipText     =   "Online Users"
      Top             =   3000
      Width           =   2205
   End
   Begin VB.TextBox txtipaddress 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AF6580&
      Height          =   315
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   15
      ToolTipText     =   "Your IP-Address"
      Top             =   2760
      Width           =   1650
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0099007B&
      Height          =   315
      Left            =   8100
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "IP :"
      Top             =   2760
      Width           =   375
   End
   Begin MynetChat.chameleonButton cmdFtp 
      Height          =   375
      Left            =   840
      TabIndex        =   63
      TabStop         =   0   'False
      ToolTipText     =   "File Transfer"
      Top             =   1905
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
      MICON           =   "Client.frx":31484
      PICN            =   "Client.frx":314A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ftp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   960
      TabIndex        =   64
      Top             =   2295
      Width           =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   140
      X2              =   18
      Y1              =   544
      Y2              =   544
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   536
      X2              =   360
      Y1              =   512
      Y2              =   512
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00BAB6B3&
      X1              =   16
      X2              =   688
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00BAB6B3&
      X1              =   16
      X2              =   688
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00BAB6B3&
      X1              =   440
      X2              =   440
      Y1              =   128
      Y2              =   160
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C0C0C0&
      Height          =   285
      Left            =   3360
      Top             =   960
      Width           =   765
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C0C0C0&
      Height          =   285
      Left            =   3360
      Top             =   1320
      Width           =   1725
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      Height          =   285
      Left            =   1560
      Top             =   1320
      Width           =   1725
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      Height          =   285
      Left            =   1545
      Top             =   945
      Width           =   1725
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "About Me"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5880
      TabIndex        =   38
      Top             =   2295
      Width           =   720
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Opaq"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5010
      TabIndex        =   36
      Top             =   2295
      Width           =   360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Trans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4500
      TabIndex        =   34
      Top             =   2295
      Width           =   345
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3735
      TabIndex        =   32
      Top             =   2295
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3075
      TabIndex        =   30
      Top             =   2295
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mass"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2355
      TabIndex        =   28
      Top             =   2295
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Smiley"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1635
      TabIndex        =   26
      Top             =   2295
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Chat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   25
      Top             =   2295
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   495
      Left            =   240
      Top             =   8160
      Width           =   7815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   4695
      Left            =   240
      Top             =   3000
      Width           =   7815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Topic :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0099007B&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lbltopic 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "You are the best for giving me 5 globes."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   12
      ToolTipText     =   "Topic"
      Top             =   2760
      Width           =   7215
   End
   Begin VB.Label lblclientsconn 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001784D5&
      Height          =   255
      Left            =   10080
      TabIndex        =   11
      Top             =   2070
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Online :"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9360
      TabIndex        =   10
      Top             =   2070
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   9000
      Picture         =   "Client.frx":31BD6
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Status :"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   9
      Top             =   2070
      Width           =   735
   End
   Begin VB.Label lblstatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Not Connected"
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
      Height          =   255
      Left            =   7680
      TabIndex        =   8
      Top             =   2070
      Width           =   1335
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00BAB6B3&
      BorderColor     =   &H00BAB6B3&
      Height          =   420
      Left            =   5400
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00BAB6B3&
      Height          =   495
      Left            =   240
      Top             =   7725
      Width           =   1875
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For play sound
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2

'lock key mouse function
Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long

Option Explicit

'For clipboard
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302

Dim start As Integer
Dim newstart As Integer

'color variables
Dim name_color, message_color, hyper_color, back_color As Ole_Color

'put your nick
Dim nick As String
'flag for neutral broadcast
Dim neutralflag As Boolean

'               USERLIST TAG = §
'               NAME TAG     = Æ

'spark warn message
Dim sparkstart As Long
Dim sparkend As Long
Dim sparkcount As Integer

'DECLARATIONS
Dim PORTNO As Long          'CONNECTION PORT TO SERVER
Dim PRIVATEPORTNO As Long

'GET MESSAGE FROM USERS
Dim MESSAGE As String
Dim MESSAGEX As String

'FOR COLORING TEXT
Dim POS_START As Integer
Dim FIND_POS As Integer

Dim i As Integer
Dim temp As Integer
Dim emotfind As Integer
Dim G_startfrom  As Long

'Round the form
Dim rndfrm As New ROUND_FORM

Dim x_fontname As String
Dim x_IsItalic As Boolean
Dim x_IsBold As Boolean





Private Sub chameleonButton1_Click()
frmUpdateInfo.Show
End Sub

Private Sub clientpmlisten_DataArrival(ByVal bytesTotal As Long)

On Error Resume Next
Dim MESSAGEPM As String
clientpmlisten.GetData MESSAGEPM
Dim findcolonpos As Integer
findcolonpos = InStr(1, MESSAGEPM, ":")

'create private message form instance
Dim pmfrm As New frmPrivateMessageReply

pmfrm.lblnameby.Caption = Mid(MESSAGEPM, 1, findcolonpos - 1)
pmfrm.txtprivatemessage.Text = Mid(MESSAGEPM, findcolonpos + 1, Len(MESSAGEPM) - findcolonpos)
pmfrm.Show

PLAY_SOUND "chimes"

End Sub


Private Sub cmdabout_Click()
frmAboutme.Show 1
End Sub

Private Sub cmdBold_Click()
txtmessage.Font.Bold = Not txtmessage.Font.Bold
End Sub

Private Sub cmdClear_Click()
txtchat.Text = ""
newstart = 0
G_startfrom = 0
End Sub

Private Sub cmdClose_Click()
If lblstatus.Caption = "Connected" Then
    MsgBox "Disconnect first.", vbCritical
    Exit Sub
End If
TrayArea.Visible = False
End
End Sub

Private Sub cmdConnect_Click()

On Error Resume Next
PORTNO = CLng(txtport.Text)
lblstatus.Caption = "Connecting ..."
tcpclient.RemoteHost = txthostname.Text
tcpclient.RemotePort = PORTNO
tcpclient.Connect
txtport.Text = PORTNO
End Sub


Private Sub cmdDeath_Click()
If Left(txtipaddress, 1) = "@" Then
    frmDeath.Visible = True
Else
    MsgBox "You do not have sufficient privillages.Contact your administrator", vbOKOnly, "Non sufficient privillages"
End If
End Sub

Private Sub cmddisconnect_Click()
On Error Resume Next
'SEND CLOSE MESSAGE (CMS = CLOSE MY SOCKET)
tcpclient.SendData ":CMS" & txtclientname
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
frmSender.txtEventLog.Text = ""
frmSender.Frame1.Enabled = False
End Sub

Private Sub cmdItalic_Click()
txtmessage.Font.Italic = Not txtmessage.Font.Italic
End Sub

Private Sub cmdmassmsg_Click()
frmMassMessage.Show
End Sub

Private Sub cmdMinimize_Click()
Me.Hide
TrayArea.Visible = True
Set TrayArea.Icon = LoadPicture(App.Path & "\me.ico")
End Sub

Private Sub cmdNick_Click()
frmNick.Visible = True
frmNick.txtnick.Text = txtclientname
End Sub

Private Sub cmdOpaque_Click()
MakeOpaque Me.hWnd
End Sub

Private Sub cmdPrivateChat_Click()
pvtchat_Click
End Sub

Private Sub cmdreset_Click()
txtport.Text = 10000
PORTNO = 10000
lblstatus.Caption = "Not Connected"
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

'send data immediately to server
tcpclient.SendData txtmessage.Text
DoEvents

'when user is not found on network
'udpNeutralNetworkSend.RemoteHost = txthostname
'udpNeutralNetworkSend.RemotePort = 15000
'udpNeutralNetworkSend.SendData txtmessage.Text
'txtmessage.Text = ""
'DoEvents

'broadcast data to all clients
'For i = 1 To lstusers.ListCount - 1
'    udpclientsend.RemoteHost = lstusers.List(i)
'    udpclientsend.RemotePort = 9000
'    On Error GoTo 1 'if client is on other network
'    udpclientsend.SendData txtmessage.Text
'    DoEvents
'Next

'color the text in txtchat
COLORTEXT
'detect hyperlink in txtchat
'DETECT_HYPERLINK
'clear and set the start of typing in txtmessage
txtmessage.Text = ""
txtmessage.SelStart = Len(txtmessage.Text)

'restore factory defaults
txtmessage.SelColor = vbBlack
txtmessage.SelUnderline = False
Exit Sub


End Sub

Private Sub cmdSettings_Click()
MsgBox "Currently disabled working.", vbOKOnly, "MSG-X090"
End Sub

Private Sub cmdtopic_Click()
frmTopic.Show
End Sub

Private Sub cmdTransparent_Click()
MakeTransparent Me.hWnd, 220
End Sub

Private Sub cmdUnderline_Click()
txtmessage.Font.Underline = Not txtmessage.Font.Underline
End Sub

Private Sub cmdupdclientlist_Click()
On Error Resume Next
'SEND CLOSE MESSAGE (SMUL = SEND ME USER LIST)
tcpclient.SendData "SMUL:"
DoEvents
End Sub

Private Sub Form_Load()
                
'ADD YOUR NAME TO USER LIST
txtclientname.Text = GetIPHostName

'set UDP topic port
udpTopic.Bind 4000
'set UDP FTP port
udpupdate.Bind 6000
'Set private chat port
PRIVATEPORTNO = 30000
'set listen UDP Private Message port
clientpmlisten.Bind 8000
'set listen UDP port
udpclientlisten.Bind 9000
'Broadcast neutral network port
udpNeutralNetworkListen.Bind 20000


newstart = 1
G_startfrom = 1
txtchat.SelColor = vbBlack
name_color = namecolor.BackColor
message_color = messagecolor.BackColor
hyper_color = hypercolor.BackColor
txtmessage.SelStart = Len(txtmessage.Text)

'apply greyascale when not connected
COLOR_CONTROLS False
'set tray icon
TrayArea.Visible = True
Set TrayArea.Icon = LoadPicture(App.Path & "\me.ico")
'round form shape
rndfrm.ROUND_FORM Me, 12, 1, 1
'show update form form
frmUpdateInfo.Show
frmUpdateInfo.Hide
frmSender.Visible = True
frmSender.Visible = False
frmReceiver.Visible = True
frmReceiver.Visible = False

'App.TaskVisible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
temp = 0
DoEvents
'request private chat ports from all users
For i = 1 To Client.lstusers.ListCount - 1
    udpupdate.RemoteHost = Client.lstusers.List(i)
    udpupdate.RemotePort = 6000
    'GMPPN = GIVE ME PRIVATE PORT NO
    udpupdate.SendData ":GMPPN"
    DoEvents
Next
End Sub

Private Sub hypercolor_Click()
cd.ShowColor
hypercolor.BackColor = cd.Color
hyper_color = hypercolor.BackColor
End Sub

Private Sub Label7_Click()
frmAbout.Show
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = vbWhite
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

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = vbBlack
End Sub

Private Sub pvtchat_Click()
On Error Resume Next
If lstusers.Selected(0) = False Then
    udpclientsend.RemoteHost = lstusers.List(lstusers.ListIndex)
    udpclientsend.RemotePort = 9000
    'PCR = PRIVATE CHAT REQUEST
    udpclientsend.SendData "PCR:" & udpclientsend.LocalHostName
    DoEvents
End If
End Sub

Private Sub pvtmsg_Click()
On Error Resume Next
'create private message form instance
Dim pmfrm As New frmPrivateMessageSend
pmfrm.Show

End Sub

Private Sub Timer4_Timer()
CONVERT_INTO_SMILEYS G_startfrom
End Sub

Private Sub Timer5_Timer()

End Sub

Private Sub TrayArea_MouseDown(Button As Integer)
If Button = 1 Then
    Me.Show
End If
End Sub

Private Sub txtchat_Change()
Timer4.Enabled = True
End Sub

Private Sub kck_Click()
On Error Resume Next
If Left(txtipaddress.Text, 1) = "@" Then
    'SEND CLOSE MESSAGE (CMS = CLOSE MY SOCKET)
    tcpclient.SendData ":CMS" & lstusers.List(lstusers.ListIndex)
    DoEvents
End If
End Sub


Private Sub txtchat_SelChange()
If txtmessage.SelStart < Len(nick & " >> ") Then
    txtmessage.SelStart = Len(txtmessage.Text)
    Exit Sub
End If
End Sub


Private Sub txtmessage_SelChange()
If txtmessage.SelStart < Len(nick & " >> ") Then
    DoEvents
    txtmessage.SelStart = Len(txtmessage.Text)
    Exit Sub
End If
End Sub


Private Sub udpNeutralNetworkListen_DataArrival(ByVal bytesTotal As Long)
  
On Error Resume Next

Dim y_msg As String

'SAVE MESSAGE BUFFER IN MESSAGE
udpNeutralNetworkListen.GetData y_msg
DoEvents

'LKM = LOCK KEYBOARD + MOUSE
If Left(y_msg, 4) = ":LKM" Then
    BlockInput (True)
    Exit Sub
End If

'ULKM = UNLOCK KEYBOARD + MOUSE
If Left(y_msg, 4) = ":ULKM" Then
    BlockInput (False)
    Exit Sub
End If

txtchat.SelStart = Len(txtchat.Text)
newstart = Len(txtchat.Text)
txtchat.SelText = y_msg
COLORTEXT
DoEvents

'set the selstart
txtchat.SelStart = Len(txtchat.Text)
txtchat.SelText = vbCrLf

newstart = Len(txtchat.Text)
PLAY_SOUND "start"

End Sub


Private Sub udpTopic_DataArrival(ByVal bytesTotal As Long)
Dim Msg As String
udpTopic.GetData Msg
If Left(Msg, 3) = ":NT" Then
    lbltopic.Caption = Mid(Msg, 4, Len(Msg) - 3)
End If
End Sub


Private Sub udpupdate_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Msgx As String
udpupdate.GetData Msgx
'if its a PRIVATEPORTNO request
If Msgx = ":GMPPN" Then
    udpupdate.SendData CStr(PRIVATEPORTNO)
    DoEvents
    Exit Sub
Else
    If frmUpdateInfo.grid.rows > 30 Then
        frmUpdateInfo.grid.rows = 1
    End If
    frmUpdateInfo.grid.TextMatrix(temp, 0) = Msgx
    temp = temp + 1
End If
frmUpdateInfo.grid.Sort = 1
End Sub


Private Sub wrn_Click()
On Error Resume Next
If Left(txtipaddress.Text, 1) = "@" Then
    udpclientsend.RemoteHost = lstusers.List(lstusers.ListIndex)
    udpclientsend.RemotePort = 9000
    'WU = WARN USER
    udpclientsend.SendData "WU:"
    DoEvents
End If
End Sub

Private Sub tcpclient_Close()

On Error Resume Next
cmdConnect.Enabled = True
cmddisconnect.Enabled = False
PLAY_SOUND "gunshot"
Dim ohgod As String
ohgod = MsgBox("You have been disconnected from server.", vbRetryCancel)
tcpclient.Close
lblstatus.Caption = "Not Connected"
lblclientsconn.Caption = "0"
moto2.ZOrder (1)
PORTNO = 10000
lstusers.Clear
txtmessage.Locked = True
cmdSend.Enabled = False

If ohgod = vbRetry Then
    cmdConnect_Click
ElseIf ohgod = vbCancel Then
    'apply greyscale when disconnected
    COLOR_CONTROLS False
    Exit Sub
End If

txtchat.SelText = vbCrLf & "Disconnected from the server."
txtchat.SelStart = Len(txtchat.Text) - Len("Disconnected from the server.")
txtchat.SelLength = Len("Disconnected from the server.")
txtchat.SelColor = vbBlack

End Sub

Private Sub tcpclient_Connect()

cmdConnect.Enabled = False
cmddisconnect.Enabled = True
PLAY_SOUND "connected"
txtchat.SelColor = vbBlack
txtmessage.Locked = False
Dim name As String
lblstatus.Caption = "Connected"
moto2.ZOrder (0)
lstusers.AddItem LCase(txtclientname2 & txtclientname.Text)
'SEND USERNAME TO CLIENT
tcpclient.SendData "Æ" & txtclientname2 & txtclientname.Text
DoEvents
'set nick
nick = txtclientname & txtclientname2
txtmessage.Text = nick & " >> "
txtmessage.SelStart = Len(txtmessage.Text)
'enable
cmdSend.Enabled = True
txtmessage.Locked = False
'Find IPADDRESS
txtipaddress = GetIPAddress
'apply colors to controls when connected
COLOR_CONTROLS True

End Sub


Private Sub tcpclient_DataArrival(ByVal bytesTotal As Long)

'SAVE MESSAGE BUFFER IN MESSAGE
tcpclient.GetData MESSAGE

'IF I AM BLOCKED ON SERVER
If InStr(1, MESSAGE, ":YABOS") > 0 Then
    tcpclient.Close
    MsgBox "Server has blocked you.Ask your administrator to unblock your access on server", vbOKOnly, "Blocked ..."
    End
End If

'IF IT IS A NAME OF SERVER
If Left(MESSAGE, 1) = "Æ" Then
    lstusers.AddItem LCase(Right(MESSAGE, Len(MESSAGE) - 1))
    txtchat.SelStart = Len(txtchat.Text)
    txtchat.SelText = vbCrLf & Right(MESSAGE, Len(MESSAGE) - 1) & " is online" & vbCrLf
    txtchat.SelStart = Len(txtchat.Text)
    newstart = Len(txtchat.Text)
'IF IT IS A USERLIST
ElseIf Left(MESSAGE, 1) = "§" Then
    UPDATE_USERLIST
    
'IF IT IS A USERLISTIP
'ElseIf Left(MESSAGE, 1) = "Ø" Then
    'UPDATE_USERLISTIP
    
'MAKE USER OP
ElseIf MESSAGE = "MUO:" Then
    txtipaddress = "@ " & txtipaddress
    txtchat.SelStart = Len(txtchat.Text)
    txtchat.SelText = "Server made you the OP. Congrats..." & vbCrLf
    txtchat.SelStart = Len(txtchat.Text) - Len("Server made you the OP. Congrats..." & vbCrLf)
    txtchat.SelLength = Len("Server made you the OP. Congrats..." & vbCrLf)
    txtchat.SelColor = &HE08F54
    txtchat.SelStart = Len(txtchat.Text)
    newstart = Len(txtchat.Text)
    BROADCAST_MASSMESSAGE "Server has made " & tcpclient.LocalHostName & " the Operator."
    Exit Sub
'MAKE USER DE-OP
ElseIf MESSAGE = "DOU:" Then
    txtipaddress.Text = Replace(txtipaddress.Text, "@ ", "")
    txtchat.SelStart = Len(txtchat.Text)
    txtchat.SelText = "Server DE-OP you. God bless you..." & vbCrLf
    txtchat.SelStart = Len(txtchat.Text) - Len("Server DE-OP you. God bless you..." & vbCrLf)
    txtchat.SelLength = Len("Server DE-OP you. God bless you..." & vbCrLf)
    txtchat.SelColor = &HE08F54
    txtchat.SelStart = Len(txtchat.Text)
    newstart = Len(txtchat.Text)
    BROADCAST_MASSMESSAGE "Server DE-OP " & tcpclient.LocalHostName
    Exit Sub
'WARNING FROM SERVER
ElseIf MESSAGE = "WU:" Then
    sparkcount = 0
    Timer3.Enabled = False
    txtchat.SelStart = Len(txtchat.Text)
    txtchat.SelText = "WARNING" & vbCrLf
    txtchat.SelStart = Len(txtchat.Text) - 9
    txtchat.SelLength = 7
    txtchat.SelFontSize = 50
    txtchat.SelColor = vbRed
    sparkstart = Len(txtchat.Text) - 9
    sparkend = sparkstart + 7
    Timer3.Enabled = True
    'Resume settings
    txtchat.SelStart = Len(txtchat.Text)
    newstart = Len(txtchat.Text)
    BROADCAST_MASSMESSAGE "Server warns " & tcpclient.LocalHostName & " not to Over react."
    Exit Sub

'IF IT IS A MESSAGE FROM SERVER
Else
    txtchat.SelStart = Len(txtchat.Text)
    newstart = Len(txtchat.Text)
    txtchat.SelText = MESSAGE
    COLORTEXT
    'set the selstart
    txtchat.SelStart = Len(txtchat.Text)
    txtchat.SelText = vbCrLf
    
    newstart = Len(txtchat.Text)
    PLAY_SOUND "start"
End If

MESSAGE = ""
    
End Sub

Private Sub udpclientlisten_DataArrival(ByVal bytesTotal As Long)

On Error Resume Next
'SAVE MESSAGE BUFFER IN MESSAGE
udpclientlisten.GetData MESSAGEX

If Left(MESSAGEX, 4) = "PCR:" Then
    PLAY_SOUND "ringin"
    Dim remotename As String
    remotename = Right(MESSAGEX, Len(MESSAGEX) - 4)
    Dim descision As String
    descision = MsgBox(UCase(remotename) & " wants to private chat. Do you accept the request.", vbYesNo + vbInformation, "Private chat request")
    'OPEN PRIVATE CHAT AS SERVER
    If descision = vbYes Then
        Dim pvtchatins As New PrivateChat
        pvtchatins.Show
        pvtchatins.lblname.Caption = remotename
        pvtchatins.nick = udpclientlisten.LocalHostName
        pvtchatins.txtmessage.Text = pvtchatins.nick & " >> "
        'set the port from the update form grid
        If CLng(frmUpdateInfo.grid.TextMatrix(frmUpdateInfo.grid.rows - 1, 0)) < 30000 Then
            PRIVATEPORTNO = 30000
        Else
            PRIVATEPORTNO = CLng(frmUpdateInfo.grid.TextMatrix(frmUpdateInfo.grid.rows - 1, 0)) + 1
        End If
        pvtchatins.privateclient.Bind PRIVATEPORTNO
        'PCRA = PRIVATE CHAT REQUEST ACCEPTED
        udpclientlisten.SendData "PCRA:" & PRIVATEPORTNO
        DoEvents
        PRIVATEPORTNO = PRIVATEPORTNO + 1
    Else
        'PCRD = PRIVATE CHAT REQUEST DECLINED
        udpclientlisten.SendData "PCRD:"
        DoEvents
    End If
'KTU = KICK THIS USER
ElseIf MESSAGEX = "KTU:" Then
    cmddisconnect_Click
    Exit Sub
'MM = MASS MESSAGE
ElseIf Left(MESSAGEX, 3) = "MM:" Then
    txtchat.SelText = Mid(MESSAGEX, 4, Len(MESSAGEX) - 3) & vbCrLf
    txtchat.SelStart = Len(txtchat.Text) - Len(Mid(MESSAGEX, 4, Len(MESSAGEX) - 3) & vbCrLf)
    txtchat.SelLength = Len(Mid(MESSAGEX, 4, Len(MESSAGEX) - 3) & vbCrLf)
    txtchat.SelColor = &HE08F54
    txtchat.SelStart = Len(txtchat.Text)
    newstart = Len(txtchat.Text)
    PLAY_SOUND "start"
    Exit Sub
'WU = WARN USER
ElseIf MESSAGEX = "WU:" Then
    sparkcount = 0
    Timer3.Enabled = False
    txtchat.SelStart = Len(txtchat.Text)
    txtchat.SelText = "WARNING" & vbCrLf
    txtchat.SelStart = Len(txtchat.Text) - 9
    txtchat.SelLength = 7
    txtchat.SelFontSize = 50
    txtchat.SelColor = vbRed
    sparkstart = Len(txtchat.Text) - 9
    sparkend = sparkstart + 7
    txtchat.SelStart = Len(txtchat.Text)
    newstart = Len(txtchat.Text)
    Timer3.Enabled = True
    BROADCAST_MASSMESSAGE "Server warns " & tcpclient.LocalHostName & " not to Over react."
    Exit Sub
Else
    txtchat.SelStart = Len(txtchat.Text)
    newstart = Len(txtchat.Text)
    txtchat.SelText = MESSAGEX
    COLORTEXT
    
    'set the selstart
    txtchat.SelStart = Len(txtchat.Text)
    txtchat.SelText = vbCrLf
    
    newstart = Len(txtchat.Text)
    MESSAGEX = ""
    PLAY_SOUND "start"
End If

End Sub


Private Sub udpclientlisten_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Err.Description
End Sub

Private Sub udpclientsend_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Msg As String
udpclientsend.GetData Msg
'PCRD = PRIVATE CHAT REQUEST DECLINED
If Msg = "PCRD:" Then
    MsgBox "User declined your private chat request.", vbOKOnly, "Private chat request declined"
'OPEN PRIVATE CHAT AS CLIENT
ElseIf Left(Msg, 5) = "PCRA:" Then
    Dim pvtchatins As New PrivateChat
    pvtchatins.Show
    pvtchatins.lblname.Caption = lstusers.List(lstusers.ListIndex)
    pvtchatins.nick = udpclientlisten.LocalHostName
    pvtchatins.txtmessage.Text = pvtchatins.nick & " >> "
    'for messages
    pvtchatins.privateclient.RemoteHost = udpclientsend.RemoteHostIP
    pvtchatins.privateclient.RemotePort = CLng(Mid(Msg, 6, 5))
    pvtchatins.privateclient.SendData "Connected to " & pvtchatins.privateclient.LocalHostName & ":"
    DoEvents
    'now broadcast the port
    cmdUpdate_Click
End If
End Sub

Private Sub Timer1_Timer()

start = txtmessage.SelStart

'if space is found then disabled the timer and make selcolor black again
If Mid(txtmessage.Text, txtmessage.SelStart, 1) = " " Then
    txtmessage.SelStart = InStr(start, txtmessage.Text, " ") - 1
    txtmessage.SelLength = 1
    txtmessage.SelUnderline = False
    txtmessage.SelColor = vbBlack
    txtmessage.SelStart = start + 1
    Timer1.Enabled = False
End If

End Sub

Private Sub Timer3_Timer()

If sparkcount = 25 Then
    Timer3.Enabled = False
    sparkcount = 0
    Exit Sub
End If

txtchat.SelStart = sparkstart
txtchat.SelLength = sparkend - sparkstart

If txtchat.SelColor = vbBlue Then
    txtchat.SelColor = vbRed
Else
    txtchat.SelColor = vbBlue
End If

sparkcount = sparkcount + 1

End Sub

Private Sub txtmessage_Change()

If Len(txtmessage.Text) < Len(nick & " >> ") Then
    txtmessage.Text = nick & " >> "
    txtmessage.SelStart = Len(txtmessage.Text)
    Exit Sub
End If

Dim found As Integer                        'stores the position of emoticons   {    :) or :( or :|    } which is found first
Dim EMOTICON As Integer
Dim spacefind As Integer

'if www is found (Now you are typing hyperlink)
If txtmessage.Find("http://", txtmessage.SelStart - 7, Len(txtmessage.Text)) > -1 Then
    found = txtmessage.Find("http://", txtmessage.SelStart - 7, Len(txtmessage.Text))
    SET_HYPERLINK txtmessage, found, ":http"
ElseIf txtmessage.Find("www.", txtmessage.SelStart - 4, Len(txtmessage.Text)) > -1 Then
    found = txtmessage.Find("www.", txtmessage.SelStart - 4, Len(txtmessage.Text))
    SET_HYPERLINK txtmessage, found, ":www"
End If
'convert text into emoticon
'If txtmessage.Find(":", 1, Len(txtmessage.Text)) > 1 Then
'    emotfind = txtmessage.Find(":", 1, Len(txtmessage.Text)) + 1
'    If txtmessage.Find(" ", emotfind, Len(txtmessage.Text)) > 1 And emotfind > 0 Then
'        spacefind = txtmessage.Find(" ", emotfind, Len(txtmessage.Text))
'        On Error Resume Next
'        SET_PICTURE txtmessage, emotfind - 1, Mid(txtmessage.Text, emotfind + 1, spacefind - emotfind), spacefind - emotfind + 1
'        emotfind = 0
'    End If
'    txtmessage.SelStart = Len(txtmessage.Text)
'End If
'short help:
    'txtmessage.SelStart - 2: when you are typing,the length of richtextbox is incrementing. So when you type ":" then txtmessage.SelStart will be the position of ":"
    '                               but when you type next word ")" then the text will be like this
    '                               txtmessage.text = Bla..Bla...Bla... :)
    'The function       txtmessage.Find(":)", txtmessage.SelStart - 2, Len(txtmessage.Text))         will return a number greater then -1
    'Because the length of :) is 2 therefore you will start your search 2 words before the cursor position in order to find :) or :( or :|
    
    'Hoped, I have helped you just a little bit.
    
End Sub

Private Sub txtmessage_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSend_Click
End If
End Sub



















































Public Sub UPDATE_USERLIST()

'CLEAR USERLIST FIRST
lstusers.Clear
'POSITIONS VARIABLES
Dim POS1, POS2 As Integer
Dim user As String
POS1 = 2
'UPDATE USERLIST
Do While (InStr(POS1 + 1, MESSAGE, "*") > 0)
    POS2 = InStr(POS1 + 1, MESSAGE, "*")
    user = LCase(Mid(MESSAGE, POS1 + 1, POS2 - POS1 - 1))
    lstusers.AddItem user
    POS1 = POS2
Loop

'Connected clients
lblclientsconn.Caption = lstusers.ListCount
    
End Sub


Public Function COLOR_TEXT()

FIND_POS = InStr(POS_START, txtchat.Text, ">>")
txtchat.SelStart = POS_START
txtchat.SelLength = FIND_POS - POS_START + 3
txtchat.SelColor = vbRed

txtchat.SelStart = FIND_POS + 2
txtchat.SelLength = Len(txtchat.Text) - FIND_POS + 3
txtchat.SelColor = vbBlue

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

Public Function DETECT_HYPERLINK()

On Error Resume Next
Dim colorupto As Integer

'This method is for scanning the hyperlink in Rich Text
Do While newstart < Len(txtchat.Text)
    If txtchat.Find("www.", newstart, Len(txtchat.Text)) > 0 Then
        newstart = txtchat.Find("www.", newstart, Len(txtchat.Text))
        'find space"
        colorupto = txtchat.Find(" ", newstart, Len(txtchat.Text))
        'if no space is found
        If colorupto = -1 Then
            txtchat.SelStart = newstart
            txtchat.SelLength = Len(txtchat.Text)
        Else
            txtchat.SelStart = newstart
            txtchat.SelLength = colorupto - newstart
        End If
        'color the link and underline it
        txtchat.SelColor = hyper_color
        txtchat.SelUnderline = True
    End If
    newstart = newstart + 1
Loop

End Function


Public Function SET_HYPERLINK(rt_box As RichTextBox, pos As Integer, what As String)

'This function will make the hyperlink color blue and it remains blue until it find a space " "
'bcoz www.ugly.com  is actually www.ugly.com(space)
'space means your hyperlink ended here so it will again change the color to black

start = pos
If what = ":www" Then
    If rt_box.Find("www", start - 4, Len(rt_box)) > 0 Then
        'First make "www" color blue and underline it
        rt_box.SelStart = start
        rt_box.SelLength = 4
        rt_box.SelColor = hyper_color
        rt_box.SelUnderline = True
        'now set start position and enable the timer until space is found
        rt_box.SelStart = start + 4
        Timer1.Enabled = True
    End If
ElseIf what = ":http" Then
    If rt_box.Find("http://", start - 8, Len(rt_box)) > 0 Then
        'First make "www" color blue and underline it
        rt_box.SelStart = start
        rt_box.SelLength = 8
        rt_box.SelColor = hyper_color
        rt_box.SelUnderline = True
        'now set start position and enable the timer until space is found
        rt_box.SelStart = start + 8
        Timer1.Enabled = True
    End If
End If

    
End Function


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


Private Sub txtport_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub


Public Function BROADCAST_MASSMESSAGE(Msg As String)
On Error Resume Next
'broadcast this message
For i = 1 To lstusers.ListCount - 1
    udpclientsend.RemoteHost = lstusers.List(i)
    udpclientsend.RemotePort = 9000
    udpclientsend.SendData "MM:" & Msg
    DoEvents
Next
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

Public Function PLAY_SOUND(Filename As String)
sndPlaySound App.Path & "\" & Filename, SND_ASYNC Or SND_NODEFAULT
End Function

Private Sub WindowBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hWnd, &HA1, 2, 0
  Exit Sub
 End If
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


Public Property Get newnick() As String
newnick = nick
End Property

Public Property Let newnick(ByVal vNewValue As String)
nick = vNewValue
End Property
