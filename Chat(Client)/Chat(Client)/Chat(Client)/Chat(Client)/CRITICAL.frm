VERSION 5.00
Begin VB.Form CRITICAL 
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   0  'None
   Caption         =   "CRITICAL"
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox WindowBorder 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      Picture         =   "CRITICAL.frx":0000
      ScaleHeight     =   420
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin MynetChat.chameleonButton cmdClose 
         Height          =   255
         Left            =   7130
         TabIndex        =   1
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
         MICON           =   "CRITICAL.frx":AB8A
         PICN            =   "CRITICAL.frx":ABA6
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
End
Attribute VB_Name = "CRITICAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
