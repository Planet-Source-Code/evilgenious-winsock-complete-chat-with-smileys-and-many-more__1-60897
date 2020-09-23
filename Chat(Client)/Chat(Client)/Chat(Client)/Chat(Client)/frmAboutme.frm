VERSION 5.00
Begin VB.Form frmAboutme 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Who am I ??"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   Icon            =   "frmAboutme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
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
      Left            =   6000
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "han"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3795
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "udassir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmAboutme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOk_Click()
Unload Me
End Sub
