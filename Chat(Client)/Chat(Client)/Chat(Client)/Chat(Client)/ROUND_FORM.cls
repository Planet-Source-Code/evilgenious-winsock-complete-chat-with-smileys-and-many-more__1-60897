VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ROUND_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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



Public Function ROUND_FORM(frmobj As Form, cornertwist As Integer, verticalborderwidth As Double, horizontalborderwidth As Double)
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

    hBrush = CreateSolidBrush(&H800000)
    lRet = FrameRgn(frmobj.hDC, mlTipBox, hBrush, verticalborderwidth, horizontalborderwidth)
    lRet = SetWindowRgn(frmobj.hWnd, mlTipBox, True)
    
    frmobj.Refresh ' This clears the drawing area of any e-junk from this above.

    ' This second one draws it.
    'iDrawTxt = DrawText(hDC, sHelp, Len(sHelp), tipRC, DT_LEFT Or DT_WORDBREAK)


End Function
