VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Server 
   BackColor       =   &H00EBF5F4&
   Caption         =   "Mynet Chat (Server)"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9760
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00EBF5F4&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7305
         Picture         =   "Server.frx":0CCA
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   45
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtportno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "10000"
         Top             =   580
         Width           =   2295
      End
      Begin VB.TextBox txtservername 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00EBF5F4&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   4680
         Picture         =   "Server.frx":0E0F
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   300
         Width           =   495
      End
      Begin VB.CommandButton cmdlisten 
         BackColor       =   &H00EBF5F4&
         Caption         =   "        S T A R T  S E R V E R"
         Height          =   615
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   2560
      End
      Begin VB.CommandButton cmddisconnect 
         BackColor       =   &H00EBF5F4&
         Caption         =   "        C L O S E  S E R V E R"
         Enabled         =   0   'False
         Height          =   615
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   2500
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   735
         Left            =   130
         Top             =   190
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00EBF5F4&
         Height          =   735
         Left            =   120
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Port :"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   615
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   630
         Left            =   240
         Picture         =   "Server.frx":2081
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Port :"
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
         Index           =   3
         Left            =   1210
         TabIndex        =   48
         Top             =   625
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Server :"
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
         Left            =   1200
         TabIndex        =   6
         Top             =   275
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Server :"
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
         Index           =   3
         Left            =   1215
         TabIndex        =   47
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   3315
      Top             =   4725
   End
   Begin VB.ListBox lstusersip 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   1515
      TabIndex        =   25
      Top             =   8445
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   960
      Width           =   4095
      Begin VB.Label lblstatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Listening."
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   180
         Width           =   1095
      End
      Begin VB.Image Image4 
         Height          =   270
         Left            =   220
         Picture         =   "Server.frx":256A
         Top             =   150
         Width           =   390
      End
      Begin VB.Shape shpstatus 
         BorderColor     =   &H00A56D39&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   165
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1210
         TabIndex        =   49
         Top             =   195
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4155
      TabIndex        =   11
      Top             =   960
      Width           =   3060
      Begin VB.Label lblclientsconn 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   180
         Width           =   500
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   240
         Picture         =   "Server.frx":2929
         Top             =   -20
         Width           =   480
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Users Online :"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Users Online :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   850
         TabIndex        =   50
         Top             =   195
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7275
      TabIndex        =   8
      Top             =   960
      Width           =   2500
      Begin VB.Label lbltime 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   180
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Server.frx":35F3
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Time :"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Time :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   610
         TabIndex        =   51
         Top             =   195
         Width           =   495
      End
   End
   Begin MSWinsockLib.Winsock tcpserver 
      Index           =   0
      Left            =   3840
      Top             =   4725
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   1410
      Width           =   9780
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00D0FDED&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "This Software is Copyright © by Hackers Underworld Corp ®."
         ForeColor       =   &H00404040&
         Height          =   230
         Left            =   50
         TabIndex        =   18
         Top             =   120
         Width           =   9700
      End
   End
   Begin VB.Frame Frame10 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      ForeColor       =   &H80000008&
      Height          =   650
      Left            =   0
      TabIndex        =   19
      Top             =   1725
      Width           =   9780
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "HAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   5290
         TabIndex        =   56
         Top             =   250
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "YNET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   3805
         TabIndex        =   54
         Top             =   250
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "HAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   2
         Left            =   5280
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   2
         Left            =   4800
         TabIndex        =   23
         Top             =   40
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "YNET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   0
         Left            =   3795
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome Waqas"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8040
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   7320
         Picture         =   "Server.frx":4865
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome Waqas"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8055
         TabIndex        =   46
         Top             =   375
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   0
         Left            =   3270
         TabIndex        =   21
         Top             =   40
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         Index           =   5
         Left            =   4815
         TabIndex        =   55
         Top             =   45
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         Index           =   4
         Left            =   3285
         TabIndex        =   53
         Top             =   60
         Width           =   495
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7230
      TabIndex        =   32
      Top             =   2325
      Width           =   2550
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   40
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "IP :"
         Top             =   130
         Width           =   480
      End
      Begin VB.TextBox txtipaddress 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   525
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "IP ADDRESS"
         Top             =   130
         Width           =   1990
      End
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      ForeColor       =   &H80000008&
      Height          =   4930
      Left            =   7230
      TabIndex        =   35
      Top             =   2760
      Width           =   2550
      Begin VB.ListBox lstusers 
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
         Height          =   4740
         Left            =   375
         TabIndex        =   36
         Top             =   150
         Width           =   2130
      End
      Begin VB.ListBox lstusersnumber 
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
         ForeColor       =   &H00000000&
         Height          =   4740
         ItemData        =   "Server.frx":4CA2
         Left            =   30
         List            =   "Server.frx":4CA9
         TabIndex        =   37
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      ForeColor       =   &H80000008&
      Height          =   4710
      Left            =   0
      TabIndex        =   26
      Top             =   2325
      Width           =   7215
      Begin MSWinsockLib.Winsock udpNeutralNetworkSend 
         Left            =   4320
         Top             =   2400
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin VB.CommandButton clearchat 
         BackColor       =   &H00EBF5F4&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   6000
         Picture         =   "Server.frx":4CB1
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Clear chat"
         Top             =   3860
         Width           =   705
      End
      Begin RichTextLib.RichTextBox txtchat 
         Height          =   4550
         Left            =   30
         TabIndex        =   28
         Top             =   120
         Width           =   7140
         _ExtentX        =   12594
         _ExtentY        =   8017
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Server.frx":557B
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
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   29
      Top             =   6960
      Width           =   7200
      Begin VB.CommandButton cmdsend 
         Height          =   420
         Left            =   5880
         Picture         =   "Server.frx":55F8
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   200
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox txtmessage 
         Height          =   580
         Left            =   30
         TabIndex        =   31
         Top             =   120
         Width           =   7120
         _ExtentX        =   12568
         _ExtentY        =   1032
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"Server.frx":6E0E
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
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF5F4&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   38
      Top             =   7650
      Width           =   9780
      Begin VB.CommandButton cmdBlockedUsers 
         BackColor       =   &H00EBF5F4&
         Height          =   420
         Left            =   7880
         Picture         =   "Server.frx":6E8B
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "update user list"
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdhelp 
         BackColor       =   &H00EBF5F4&
         Height          =   420
         Left            =   9120
         Picture         =   "Server.frx":72CB
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "help me ??  Evil"
         Top             =   150
         Width           =   495
      End
      Begin VB.PictureBox namecolor 
         BackColor       =   &H00800000&
         Height          =   255
         Left            =   1515
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   42
         ToolTipText     =   "chat name color"
         Top             =   220
         Width           =   255
      End
      Begin VB.PictureBox messagecolor 
         BackColor       =   &H00BF1AA3&
         Height          =   255
         Left            =   1875
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   41
         ToolTipText     =   "chat message color"
         Top             =   220
         Width           =   255
      End
      Begin VB.PictureBox hypercolor 
         BackColor       =   &H00C00000&
         Height          =   255
         Left            =   2235
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   40
         ToolTipText     =   "URL color"
         Top             =   220
         Width           =   255
      End
      Begin VB.CommandButton cmdupdclientlist 
         BackColor       =   &H00EBF5F4&
         Height          =   420
         Left            =   8505
         Picture         =   "Server.frx":86D5
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "update user list"
         Top             =   150
         Width           =   495
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   3240
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Colors"
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
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Colors"
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
         Index           =   0
         Left            =   130
         TabIndex        =   52
         Top             =   250
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Menu opt 
      Caption         =   "Options"
      Begin VB.Menu mkick 
         Caption         =   "Kick"
      End
      Begin VB.Menu msendmsg 
         Caption         =   "Send Message"
      End
      Begin VB.Menu mkop 
         Caption         =   "Make Op"
      End
      Begin VB.Menu dop 
         Caption         =   "De Op"
      End
      Begin VB.Menu wrn 
         Caption         =   "Warn"
      End
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302


Dim start As Integer
Dim newstart As Integer

'color variables
Dim name_color, message_color, hyper_color As OLE_COLOR

'put your nick
Dim nick As String

'               USERLIST TAG = §
'               NAME TAG     = Æ

'DECLARATIONS
Dim PORTNO As Long          'LISTEN PORT OF SERVER
Dim CLIENTNO As Integer
Dim CONNCLIENTNO As Integer 'CONNECTED CLIENT NO

Dim USERLIST As String
Dim USERLISTIP As String

'GET MESSAGE FROM USERS
Dim MESSAGE As String
'FOR COLORING TEXT
Dim POS_START As Integer
Dim FIND_POS As Integer

Dim i As Integer

Dim serverindex As Integer

Private Sub clearchat_Click()
txtchat.Text = ""
newstart = 0
End Sub

Private Sub cmdBlockedUsers_Click()
BlockedList.Show
End Sub

Private Sub cmddisconnect_Click()

'close all sockets
Dim i As Integer
For i = 1 To tcpserver.Count - 1
    tcpserver_Close i
    Unload tcpserver(i)
Next
cmddisconnect.Enabled = False
cmdlisten.Enabled = True
lblclientsconn.Caption = 1
PORTNO = CLng(txtportno.Text)

lblstatus.Caption = "Listening "
shpstatus.FillColor = vbRed
txtchat.SelText = "Server is closed."

End Sub

Private Sub cmdhelp_Click()
MsgBox "Welcome from EVILGENIOUS. Please vote for me if you like this chat", vbOKOnly
End Sub

Private Sub cmdlisten_Click()

On Error Resume Next
txtchat.SelText = vbCrLf & vbCrLf & vbCrLf
txtchat.SelText = "Starting Server" & vbTab & ": " & txtservername.Text & " ..." & vbCrLf
txtchat.SelText = vbCrLf & "Starting Time " & vbTab & ": " & Time & vbCrLf
txtchat.SelText = "Message Port " & vbTab & ": 10000" & vbCrLf & vbCrLf
txtchat.SelText = "Server started successfully." & vbCrLf
lblstatus.Caption = "Listening "
shpstatus.FillColor = vbGreen
cmdlisten.Enabled = False
cmddisconnect.Enabled = True
txtmessage.SetFocus
txtmessage.SelStart = Len(txtmessage.Text)
        
End Sub

Private Sub cmdsend_Click()

'Set SelStart = 0 to copy the text from start
txtmessage.SelStart = 0
'Set lenght upto the length of txtmessage
txtmessage.SelLength = Len(txtmessage.Text)

'Set SelStart = length of txtchat
txtchat.SelStart = Len(txtchat.Text)
'copy the contents to txtchat
newstart = Len(txtchat.Text)
txtchat.SelStart = newstart
txtchat.SelText = txtmessage.SelRTF

'send data immediately
BROADCAST txtmessage.Text

txtchat.SelText = vbCrLf
'color the text in txtchat
'COLORTEXT
'detect hyperlink
'DETECT_HYPERLINK
'clear and set the start of typing
txtmessage.Text = ""
txtmessage.SelStart = Len(txtmessage.Text)

'restore factory defaults
txtmessage.SelColor = vbBlack
txtmessage.SelUnderline = False

End Sub

Private Sub cmdupdclientlist_Click()
SEND_USER_LIST_TO_ALL_CLIENTS
End Sub

Private Sub dop_Click()
On Error Resume Next
tcpserver(lstusersnumber.List(lstusers.LISTINDEX)).SendData "DOU:"
End Sub

Private Sub Form_Load()

txtservername.Text = GetIPHostName
txtipaddress.Text = GetIPAddress
'ADD YOUR NAME TO USER LIST
lstusers.AddItem txtservername.Text
'INPUT MESSAGE FOR STARTING PORT NUMBER FROM THE SERVER
Dim listenportstartfrom As String
listenportstartfrom = InputBox("Please give the starting Port number." & vbCrLf & vbCrLf & "Dont use reserved ports like," & vbCrLf & vbCrLf & "    Http" & vbTab & "=" & vbTab & "80,8080" & vbCrLf & "    Ftp" & vbTab & "=" & vbTab & "1080" & vbCrLf & "    Smtp" & vbTab & "=" & vbTab & "25" & vbCrLf & vbCrLf & "Hint: give Port number greater than 5000" & vbCrLf, "Starting Port Number")
PORTNO = CLng(listenportstartfrom)
txtportno.Text = PORTNO

serverindex = 0

tcpserver(0).LocalPort = txtportno
tcpserver(0).Listen

nick = txtservername.Text
txtmessage.Text = nick & " >> "
    
End Sub

Private Sub lstusers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And lstusers.LISTINDEX <> 0 Then
    Me.PopupMenu opt
End If
End Sub

Private Sub hypercolor_Click()
'cd.ShowColor
'hypercolor.BackColor = cd.Color
'hyper_color = hypercolor.BackColor
End Sub

Private Sub messagecolor_Click()
'cd.ShowColor
'messagecolor.BackColor = cd.Color
'message_color = messagecolor.BackColor
End Sub

Private Sub mkop_Click()
On Error Resume Next
tcpserver(lstusersnumber.List(lstusers.LISTINDEX)).SendData "MUO:"
End Sub

Private Sub namecolor_Click()
'cd.ShowColor
'namecolor.BackColor = cd.Color
'name_color = namecolor.BackColor
End Sub

Private Sub mkick_Click()
'kick all the instances of user
KICK_USER lstusers.List(lstusers.LISTINDEX)
End Sub

Private Sub Timer1_Timer()
cmdupdclientlist_Click
End Sub

Private Sub wrn_Click()
On Error Resume Next
tcpserver(lstusersnumber.List(lstusers.LISTINDEX)).SendData "WU:"
End Sub

Private Sub tcpserver_Close(Index As Integer)
On Error Resume Next
Dim ClientName As String
ClientName = lstusers.List(Index)
For i = 1 To lstusers.ListCount - 1
    If lstusers.List(i) = ClientName Then
        'tcpserver(lstusersnumber.List(i)).Close
        'tcpserver(lstusersnumber.List(i)).Listen
        Unload tcpserver(i)
        lstusers.RemoveItem (i)
        lstusersnumber.RemoveItem (i)
        CLIENTNO = CLIENTNO - 1
        lblclientsconn.Caption = CONNCLIENTNO
        i = i - 1
    End If
Next
BROADCAST_TO_ALL_NETWORKS ClientName & " leave the chat"
SEND_USER_LIST_TO_ALL_CLIENTS
End Sub

Private Sub tcpserver_ConnectionRequest(Index As Integer, ByVal requestID As Long)

On Error Resume Next

'ACCEPT REQUEST OF NEW CLIENT
If Index = 0 Then
    CLIENTNO = CLIENTNO + 1
    serverindex = serverindex + 1
    Load tcpserver(serverindex)
    tcpserver(serverindex).LocalPort = "10000"
    tcpserver(serverindex).Accept requestID
End If

'SEND YOUR NAME TO CLIENT
tcpserver(serverindex).SendData "Æ" & txtservername.Text
DoEvents

CONNCLIENTNO = CONNCLIENTNO + 1
lblclientsconn.Caption = CONNCLIENTNO + 1
lstusersnumber.AddItem serverindex
    
End Sub

Private Sub tcpserver_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
On Error Resume Next
'SAVE MESSAGE BUFFER IN MESSAGE
tcpserver(Index).GetData MESSAGE

'IF IT IS THE NAME OF CLIENT
If Left(MESSAGE, 1) = "Æ" Then
    lstusers.AddItem Right(MESSAGE, Len(MESSAGE) - 1)
    BROADCAST_TO_ALL_NETWORKS Right(MESSAGE, Len(MESSAGE) - 1) & " is online"
    txtchat.SelStart = Len(txtchat.Text)
    txtchat.SelText = vbCrLf & Right(MESSAGE, Len(MESSAGE) - 1) & " is online" & vbCrLf
    txtchat.SelStart = Len(txtchat.Text)
    newstart = Len(txtchat.Text)
    Exit Sub
    'If this client is blocked list then disconnect it
    If BlockedList.FOUND(Right(MESSAGE, Len(MESSAGE) - 1)) = True Then
        'YABOS = YOU ARE BLOCKED ON SERVER
        tcpserver(Index).SendData ":YABOS"
        DoEvents
        Exit Sub
    End If
    DoEvents
    SEND_USER_LIST_TO_ALL_CLIENTS
'IF IT IS A SOCKET CLOSE MESSAGE
ElseIf Left(MESSAGE, 4) = ":CMS" Then
    KICK_USER Mid(MESSAGE, 5, Len(MESSAGE) - 4)
    Exit Sub
'IF IT IS A KEYBOARD + MOUSE CLOSE MESSAGE
ElseIf Left(MESSAGE, 4) = ":LKM" Then
    LOCK_KEY_MOU Mid(MESSAGE, 5, Len(MESSAGE) - 4)
    Exit Sub
'IF IT IS A KEYBOARD + MOUSE OPEN MESSAGE
ElseIf Left(MESSAGE, 4) = ":ULKM" Then
    UNLOCK_KEY_MOU Mid(MESSAGE, 5, Len(MESSAGE) - 4)
    Exit Sub
'IF IT IS A USER LIST REQUEST
ElseIf Left(MESSAGE, 5) = "SMUL:" Then
    SEND_USER_LIST_TO_ALL_CLIENTS
    Exit Sub
'IF IT IS A MESSAGE FROM CLIENT
Else
    txtchat.SelStart = Len(txtchat.Text)
    newstart = Len(txtchat.Text)
    txtchat.SelText = MESSAGE & vbCrLf
    'set the selstart
    txtchat.SelStart = Len(txtchat.Text)
    txtchat.SelText = vbCrLf
    newstart = Len(txtchat.Text)
    BROADCAST_TO_ALL_NETWORKS MESSAGE
End If

MESSAGE = ""
   
End Sub

Private Sub tcpserver_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

Dim ClientName As String
ClientName = lstusers.List(Index)
For i = 1 To lstusers.ListCount - 1
    If lstusers.List(i) = ClientName Then
        tcpserver(lstusersnumber.List(i)).Close
        tcpserver(lstusersnumber.List(i)).Listen
        lstusers.RemoveItem (i)
        lstusersnumber.RemoveItem (i)
        CLIENTNO = CLIENTNO - 1
        lblclientsconn.Caption = CONNCLIENTNO
        i = i - 1
    End If
Next

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
    cmdsend_Click
End If
End Sub






































Public Sub SEND_USER_LIST_TO_ALL_CLIENTS()

On Error Resume Next
'CONVERT USERLIST IN TO STRING
USERLIST = ""                                   'empty list first

For i = 0 To lstusers.ListCount - 1
    If i = 0 Then
        'INSERT USERLIST TAG IN USERLIST
        USERLIST = "§" & "*" & lstusers.List(i) & "*"
    Else
        USERLIST = USERLIST & lstusers.List(i) & "*"
    End If
Next

For i = 1 To lstusersnumber.ListCount - 1
    On Error Resume Next
    tcpserver(lstusersnumber.List(i)).SendData USERLIST
    DoEvents
Next

End Sub

Public Function REMOVE_CLIENT_FROM_LIST(LISTINDEX As Integer)

On Error Resume Next
lstusers.RemoveItem (LISTINDEX)
lstusersnumber.RemoveItem (LISTINDEX)
lstusersip.RemoveItem (LISTINDEX - 1)
CONNCLIENTNO = CONNCLIENTNO - 1

End Function

Public Function BROADCAST(msg As String)

On Error Resume Next
For i = 1 To lstusersnumber.ListCount - 1
    tcpserver(lstusersnumber.List(i)).SendData msg
    DoEvents
Next


End Function

Public Function BROADCAST_TO_ALL_NETWORKS(msg As String)

On Error Resume Next
For i = 1 To lstusersnumber.ListCount - 1
    udpNeutralNetworkSend.RemoteHost = lstusers.List(i)
    udpNeutralNetworkSend.RemotePort = 20000
    udpNeutralNetworkSend.SendData msg
    DoEvents
Next

End Function

Public Function COLOR_TEXT()

FIND_POS = InStr(POS_START, txtchat.Text, ">>")
txtchat.SelStart = POS_START
txtchat.SelLength = FIND_POS - POS_START + 3
txtchat.SelColor = vbRed

txtchat.SelStart = FIND_POS + 2
txtchat.SelLength = Len(txtchat.Text) - FIND_POS + 3
txtchat.SelColor = vbBlue

End Function


Public Function DETECT_HYPERLINK()

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


Public Function SET_HYPERLINK(rt_box As RichTextBox, pos As Integer)

'This function will make the hyperlink color blue and it remains blue until it find a space " "
'bcoz www.ugly.com  is actually www.ugly.com(space)
'space means your hyperlink ended here so it will again change the color to black

start = pos
If rt_box.Find("www", start - 4, Len(rt_box)) > 0 Then
    'First make "www" color blue and underline it
    rt_box.SelStart = start
    rt_box.SelLength = 4
    rt_box.SelColor = hyper_color
    rt_box.SelUnderline = True
    'now set start position and enable the timer until space is found
    rt_box.SelStart = start + 4
End If
    
End Function


Public Function COLORTEXT()

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
colorstart = txtchat.Find("www.", newstart2, Len(txtchat.Text))
newstart2 = colorstart
colorend = txtchat.Find(" ", newstart2, Len(txtchat.Text))

If colorstart > -1 Then
    txtchat.SelStart = colorstart
    If colorend = -1 Then
        txtchat.SelLength = Len(txtchat.Text) - colorstart
        txtchat.SelColor = hyper_color
    Else
        txtchat.SelLength = colorend - colorstart
        txtchat.SelColor = hyper_color
    End If
End If

End Function



Public Function KICK_USER(name As String)
'kick all the instances of user
On Error Resume Next
Dim ClientName As String
ClientName = name
For i = 1 To lstusers.ListCount - 1
    If UCase(lstusers.List(i)) = UCase(ClientName) Then
        Unload tcpserver(i)
        'tcpserver(lstusersnumber.List(i)).Close
        'tcpserver(lstusersnumber.List(i)).Listen
        lstusers.RemoveItem (i)
        lstusersnumber.RemoveItem (i)
        CLIENTNO = CLIENTNO - 1
        lblclientsconn.Caption = CONNCLIENTNO
        i = i - 1
    End If
Next
End Function

Public Function LOCK_KEY_MOU(name As String)
On Error Resume Next
udpNeutralNetworkSend.RemoteHost = name
udpNeutralNetworkSend.RemotePort = 20000
udpNeutralNetworkSend.SendData ":LKM"
DoEvents
End Function

Public Function UNLOCK_KEY_MOU(name As String)
On Error Resume Next
udpNeutralNetworkSend.RemoteHost = name
udpNeutralNetworkSend.RemotePort = 20000
udpNeutralNetworkSend.SendData ":ULKM"
DoEvents
End Function
