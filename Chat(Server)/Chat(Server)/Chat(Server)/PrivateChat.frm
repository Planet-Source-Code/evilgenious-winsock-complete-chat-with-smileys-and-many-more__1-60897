VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PrivateChat 
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   50
      TabIndex        =   5
      Top             =   -120
      Width           =   5175
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SMDANYAL"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10
         TabIndex        =   6
         Top             =   120
         Width           =   5140
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   50
      TabIndex        =   2
      Top             =   3000
      Width           =   5175
      Begin VB.CommandButton cmdsend 
         Height          =   300
         Left            =   3910
         Picture         =   "PrivateChat.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   280
         Width           =   1095
      End
      Begin RichTextLib.RichTextBox txtmessage 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         _Version        =   393217
         TextRTF         =   $"PrivateChat.frx":1816
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
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   50
      TabIndex        =   0
      Top             =   360
      Width           =   5175
      Begin MSWinsockLib.Winsock privateserversend 
         Left            =   2040
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSWinsockLib.Winsock privateserverlisten 
         Left            =   1560
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin RichTextLib.RichTextBox txtchat 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4048
         _Version        =   393217
         ReadOnly        =   -1  'True
         TextRTF         =   $"PrivateChat.frx":1891
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
   End
End
Attribute VB_Name = "PrivateChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdsend_Click()
privateserversend.RemoteHost = "127.0.0.1"
privateserversend.RemotePort = 7000
privateserversend.SendData txtmessage.Text
End Sub

