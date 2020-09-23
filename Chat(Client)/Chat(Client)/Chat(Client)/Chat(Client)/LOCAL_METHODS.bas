Attribute VB_Name = "LOCAL_METHODS"


Option Explicit

Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
'



Public Function COLOR_CONTROLS(FLAG As Boolean)

If FLAG = False Then
'greyscale all the controls
    With Client
        .txthostname.ForeColor = &HBFB6B5
        .txtport.ForeColor = &HBFB6B5
        .txtclientname.ForeColor = &HBFB6B5
        .txtclientname2.ForeColor = &HBFB6B5
        .lblstatus.ForeColor = &HBFB6B5
        .lblclientsconn.ForeColor = &HBFB6B5
        .lbltopic.ForeColor = &HBFB6B5
        .txtipaddress.ForeColor = &HBFB6B5
        .backgroundcolor.BackColor = &HBFB6B5
        .namecolor.BackColor = &HBFB6B5
        .messagecolor.BackColor = &HBFB6B5
        .messagecolor.BackColor = &HBFB6B5
        .hypercolor.BackColor = &HBFB6B5
        .cmdAbout.UseGreyscale = True
        .cmdupdclientlist.UseGreyscale = True
        .cmdemoticons.UseGreyscale = True
        .cmdmassmsg.UseGreyscale = True
        .cmdTransparent.UseGreyscale = True
        .wrn.UseGreyscale = True
        .pvtchat.UseGreyscale = True
        .pvtmsg.UseGreyscale = True
        .kck.UseGreyscale = True
        .cmdFont.UseGreyscale = True
        .cmdClear.UseGreyscale = True
        .cmdOpaque.UseGreyscale = True
        .cmdtopic.UseGreyscale = True
        .cmdPrivateChat.UseGreyscale = True
    End With
Else
'color all the controls
    With Client
        .txthostname.ForeColor = &HB00009
        .txtport.ForeColor = &HB00009
        .txtclientname.ForeColor = &HB00009
        .txtclientname2.ForeColor = &HB00009
        .lblstatus.ForeColor = &H8000&
        .lblclientsconn.ForeColor = &H1784D5
        .lbltopic.ForeColor = &HFF&
        .txtipaddress.ForeColor = &HAF6580
        .backgroundcolor.BackColor = &HFFFFFF
        .namecolor.BackColor = &H9A7B34
        .messagecolor.BackColor = &H669E4E
        .hypercolor.BackColor = &HFF8080
        .cmdAbout.UseGreyscale = False
        .cmdupdclientlist.UseGreyscale = False
        .cmdemoticons.UseGreyscale = False
        .cmdmassmsg.UseGreyscale = False
        .cmdTransparent.UseGreyscale = False
        .wrn.UseGreyscale = False
        .pvtchat.UseGreyscale = False
        .pvtmsg.UseGreyscale = False
        .kck.UseGreyscale = False
        .cmdFont.UseGreyscale = False
        .cmdClear.UseGreyscale = False
        .cmdOpaque.UseGreyscale = False
        .cmdtopic.UseGreyscale = False
        .cmdPrivateChat.UseGreyscale = False
    End With
End If

End Function

