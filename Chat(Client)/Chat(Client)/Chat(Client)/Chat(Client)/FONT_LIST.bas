Attribute VB_Name = "FONTLIST"
'=========================================================================================
'  Module 1
'  Font declares, constants and types
'=========================================================================================
'  Adapted and Modified By: Behrooz Sangani
'  Published Date: 19/11/2001
'  Email: bs20014@yahoo.com
'  Send comments to the address above! This code is _
   just an improved msdn example of using fonts without _
   calling the font dialog box ...
'=========================================================================================
'  Based On: MSDN Examples
'=========================================================================================


'Font enumeration types
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE) As Byte
End Type

Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type

' ntmFlags field flags
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&

'  tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4

Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0

'  EnumFonts Masks
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4


Declare Function EnumFontFamilies Lib "gdi32" Alias _
    "EnumFontFamiliesA" _
    (ByVal hdc As Long, ByVal lpszFamily As String, _
    ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
    ByVal hdc As Long) As Long
'=========================================================================================
Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
    ByVal FontType As Long, lParam As ListBox) As Long 'Make font parameters
    Dim FaceName As String
    Dim FullName As String
    On Error Resume Next
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    lParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
    EnumFontFamProc = 1
End Function 'EnumFontFamProc
'=========================================================================================
Sub FillListWithFonts(LB As ListBox) 'Adds system fonts to list box
    Dim hdc As Long
    On Error Resume Next
    LB.Clear
    hdc = GetDC(LB.hwnd)
    EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, LB
    ReleaseDC LB.hwnd, hdc
End Sub 'FillListWithFonts(LB As ListBox)
'=========================================================================================
Sub FillComboWithFonts(CB As ComboBox) 'Adds system fonts to combo box
    Dim hdc As Long
    On Error Resume Next
    CB.Clear
    hdc = GetDC(CB.hwnd)
    EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, CB
    ReleaseDC CB.hwnd, hdc
End Sub 'FillComboWithFonts(CB As ComboBox)
'=========================================================================================



