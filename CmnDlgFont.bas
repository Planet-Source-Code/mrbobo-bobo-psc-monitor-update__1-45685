Attribute VB_Name = "CmnDlgFont"
'******************************************************************
'***************Copyright PSST 2003********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'Standard Commondialog Font Dialog - nothing special here
'this code is available EVERYWHERE
Option Explicit
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_NOSIZESEL = &H200000

Const REGULAR_FONTTYPE = &H400
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type LOGFONT
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
    lfFaceName As String * 31
End Type
Private Type CHOOSEFONT
    lStructSize As Long
    hwndOwner As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long

Public Type BBfont
    FontName As String
    Fontsize As Integer
    FontBold As Boolean
    FontItalic As Boolean
    FontUnderline As Boolean
    FontStrikethru As Boolean
    FontColor As Long
End Type
Public SelectFont As BBfont

Public Function ShowFont(Ownerform As Long, Optional NoFX As Boolean, Optional NoSize As Boolean) As Boolean
    Dim cf As CHOOSEFONT, lfont As LOGFONT, hMem As Long, pMem As Long
    Dim retval As Long
    lfont.lfHeight = 0
    lfont.lfItalic = SelectFont.FontItalic
    lfont.lfUnderline = SelectFont.FontUnderline
    lfont.lfStrikeOut = SelectFont.FontStrikethru
    lfont.lfEscapement = 0
    lfont.lfOrientation = 0
    lfont.lfHeight = SelectFont.Fontsize * 1.33
    If SelectFont.FontBold Then
        lfont.lfWidth = 700
    Else
        lfont.lfWidth = 0
    End If
    lfont.lfCharSet = DEFAULT_CHARSET
    lfont.lfOutPrecision = OUT_DEFAULT_PRECIS
    lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS
    lfont.lfQuality = DEFAULT_QUALITY
    lfont.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN
    lfont.lfFaceName = SelectFont.FontName & vbNullChar
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem)
    CopyMemory ByVal pMem, lfont, Len(lfont)
    cf.lStructSize = Len(cf)
    cf.hwndOwner = Ownerform
    cf.lpLogFont = pMem
    cf.iPointSize = SelectFont.Fontsize * 10
    cf.rgbColors = SelectFont.FontColor
    If SelectFont.FontBold Or SelectFont.FontItalic Then
        cf.nFontType = IIf(SelectFont.FontBold, BOLD_FONTTYPE, 0) Or IIf(SelectFont.FontItalic, ITALIC_FONTTYPE, 0)
    Else
        cf.nFontType = REGULAR_FONTTYPE
    End If
    cf.nSizeMin = 10
    cf.nSizeMax = 72
    If NoFX Then
        cf.Flags = CF_BOTH Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT
    Else
        cf.Flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT
    End If
    If NoSize Then cf.Flags = cf.Flags Or CF_NOSIZESEL
    retval = CHOOSEFONT(cf)
    If retval <> 0 Then
        ShowFont = True
        CopyMemory lfont, ByVal pMem, Len(lfont)
        With SelectFont
            .FontName = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
            .FontBold = False
            .FontItalic = False
            .FontUnderline = False
            .FontStrikethru = False
            .Fontsize = cf.iPointSize / 10
            If lfont.lfWeight = 700 Then .FontBold = True
            .FontItalic = lfont.lfItalic
            .FontUnderline = lfont.lfUnderline
            .FontStrikethru = lfont.lfStrikeOut
            .FontColor = cf.rgbColors
        End With
    Else
        ShowFont = False
    End If
    retval = GlobalUnlock(hMem)
    retval = GlobalFree(hMem)
End Function
