VERSION 5.00
Begin VB.Form BlockedList 
   BackColor       =   &H00EBF5F4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Blocked User Area"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2235
   Icon            =   "BlockedList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   2235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   7320
   End
   Begin VB.ListBox lstBlockedUsers 
      Height          =   2595
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   ">>"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtClientName 
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
      TabIndex        =   1
      Text            =   "Username"
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Add User:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Block List:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "BlockedList"
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
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
  ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
  ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
  
  
  


Private Sub cmdAdd_Click()
lstBlockedUsers.AddItem txtClientName
End Sub


Public Function FOUND(ClientName As String) As Boolean

FOUND = False

For i = 0 To lstBlockedUsers.ListCount
    If UCase(lstBlockedUsers.List(i)) = UCase(ClientName) Then
        FOUND = True
    End If
Next

End Function


Private Sub cmdRemove_Click()
On Error Resume Next
lstBlockedUsers.RemoveItem (lstBlockedUsers.LISTINDEX)
End Sub

Private Sub Command3_Click()
Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
End Sub

Private Sub Timer1_Timer()
Me.Left = Server.Left + Server.Width
Me.Top = Server.Top
End Sub
