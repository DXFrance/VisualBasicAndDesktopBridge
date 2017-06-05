Attribute VB_Name = "Module1"
Option Explicit

Public glMgLeft        As Long         'Marge Gauche (Paysage)
Public glMgTop         As Long         'Marge Haute (Paysage)
Public glMgRight       As Long         'Marge droite (Paysage)
Public glMgBottom      As Long         'Marge basse (Paysage)
Public md_Ratio        As Double
Public ml_PrtLeft      As Long
Public ml_PrtRight     As Long
Public ml_PrtTop       As Long
Public ml_PrtBottom    As Long
Public ml_PrtScaleX    As Long
Public ml_PrtScaleY    As Long

Public m_fFactorPrinterMemo As Single   'Memorisation lors du printer
Public gbPrtFontBlack  As Boolean      'Force la couleurs les textes en noir l'impression

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
    lfFaceName As String * 32
End Type

Public Const FW_DONTCARE = 0
Public Const FW_THIN = 100
Public Const FW_EXTRALIGHT = 200
Public Const FW_ULTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_NORMAL = 400
Public Const FW_REGULAR = 400
Public Const FW_MEDIUM = 500
Public Const FW_SEMIBOLD = 600
Public Const FW_DEMIBOLD = 600
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_ULTRABOLD = 800
Public Const FW_HEAVY = 900
Public Const FW_BLACK = 900
Public Const ANSI_CHARSET = 0
Public Const ARABIC_CHARSET = 178
Public Const BALTIC_CHARSET = 186
Public Const CHINESEBIG5_CHARSET = 136
Public Const DEFAULT_CHARSET = 1
Public Const EASTEUROPE_CHARSET = 238
Public Const GB2312_CHARSET = 134
Public Const GREEK_CHARSET = 161
Public Const HANGEUL_CHARSET = 129
Public Const HEBREW_CHARSET = 177
Public Const JOHAB_CHARSET = 130
Public Const MAC_CHARSET = 77
Public Const OEM_CHARSET = 255
Public Const RUSSIAN_CHARSET = 204
Public Const SHIFTJIS_CHARSET = 128
Public Const SYMBOL_CHARSET = 2
Public Const THAI_CHARSET = 222
Public Const TURKISH_CHARSET = 162
Public Const OUT_DEFAULT_PRECIS = 0
Public Const OUT_DEVICE_PRECIS = 5
Public Const OUT_OUTLINE_PRECIS = 8
Public Const OUT_RASTER_PRECIS = 6
Public Const OUT_STRING_PRECIS = 1
Public Const OUT_STROKE_PRECIS = 3
Public Const OUT_TT_ONLY_PRECIS = 7
Public Const OUT_TT_PRECIS = 4
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_EMBEDDED = 128
Public Const CLIP_LH_ANGLES = 16
Public Const CLIP_STROKE_PRECIS = 2
Public Const ANTIALIASED_QUALITY = 4
Public Const DEFAULT_QUALITY = 0
Public Const DRAFT_QUALITY = 1
Public Const NONANTIALIASED_QUALITY = 3
Public Const PROOF_QUALITY = 2
Public Const DEFAULT_PITCH = 0
Public Const FIXED_PITCH = 1
Public Const VARIABLE_PITCH = 2
Public Const FF_DECORATIVE = 80
Public Const FF_DONTCARE = 0
Public Const FF_ROMAN = 16
Public Const FF_SCRIPT = 64
Public Const FF_SWISS = 32

Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (lplf As LOGFONT) As Long
Public Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Const TRANSPARENT = 1
Public Const OPAQUE = 2

Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long

'Public Const DT_BOTTOM = &H8
'Public Const DT_CALCRECT = &H400
'Public Const DT_CENTER = &H1
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
'Public Const DT_LEFT = &H0
'Public Const DT_NOCLIP = &H100
'Public Const DT_NOPREFIX = &H800
'Public Const DT_RIGHT = &H2
'Public Const DT_SINGLELINE = &H20
'Public Const DT_TABSTOP = &H80
'Public Const DT_TOP = &H0
'Public Const DT_VCENTER = &H4
'Public Const DT_WORDBREAK = &H10
Public Const DT_PATH_ELLIPSIS = &H4000
'Public Const DT_END_ELLIPSIS = &H8000&
Public Const DT_WORD_ELLIPSIS = &H40000

Public Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Public Const SO_ANSI_FIXED_FONT = 11
Public Const SO_ANSI_VAR_FONT = 12
Public Const SO_BLACK_BRUSH = 4          'A solid black brush.
Public Const SO_BLACK_PEN = 7            'A solid black pen.
Public Const SO_DEFAULT_GUI_FONT = 17    'Win 95/98 only: The default font for user objects under Windows.
Public Const SO_DEFAULT_PALETTE = 15     'The default system palette.
Public Const SO_DEVICE_DEFAULT_FONT = 14 'Win NT only: a device-dependent font.
Public Const SO_DKGRAY_BRUSH = 3         'A solid dark gray brush.
Public Const SO_GRAY_BRUSH = 2           'A solid gray brush.
Public Const SO_HOLLOW_BRUSH = 5         'Same as NULL_BRUSH.
Public Const SO_LTGRAY_BRUSH = 1         'A solid light gray brush.
Public Const SO_NULL_BRUSH = 5           'A null brush; i.e., a brush that does not draw anything on the device.
Public Const SO_NULL_PEN = 8             'A null pen; i.e., a pen that does not draw anything on the device.
Public Const SO_OEM_FIXED_FONT = 10      'The Original Equipment Manufacturer's default monospaced font.
Public Const SO_SYSTEM_FIXED_FONT = 16   'The system monospaced font under pre-3.x versions of Windows.
Public Const SO_SYSTEM_FONT = 13         'The system font (used for most system objects under Windows).
Public Const SO_WHITE_BRUSH = 0          'A solid white brush.
Public Const SO_WHITE_PEN = 6            'A solid white pen.

'Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function SetRect Lib "user32" (lpRect As Win32API.RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Public Const CLR_INVALID = -1
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Public Const SRCCOPY = &HCC0020
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long


Public Const PS_ALTERNATE = 8
Public Const PS_COSMETIC = 0
Public Const PS_DASH = 1
Public Const PS_DOT = 2
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_INSIDEFRAME = 6
Public Const PS_ENDCAP_FLAT = 512
Public Const PS_ENDCAP_MASK = 3840
Public Const PS_ENDCAP_ROUND = 0
Public Const PS_ENDCAP_SQUARE = 256
Public Const PS_GEOMETRIC = 65536
Public Const PS_JOIN_BEVEL = 4096
Public Const PS_JOIN_MASK = 61440
Public Const PS_JOIN_MITER = 8192
Public Const PS_JOIN_ROUND = 0
Public Const PS_NULL = 5
Public Const PS_SOLID = 0
Public Const PS_STYLE_MASK = 15
Public Const PS_TYPE_MASK = 983040
Public Const PS_USERSTYLE = 7

Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function Escape Lib "gdi32" (ByVal hdc As Long, ByVal nEs As Long, ByVal nCount As Long, ByVal lpInData As Any, lpOutData As Any) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Const PHYSICALWIDTH As Long = 110
Public Const PHYSICALHEIGHT As Long = 111
Public Const PHYSICALOFFSETX As Long = 112
Public Const PHYSICALOFFSETY As Long = 113
Public Const GETPHYSPAGESIZE = 12
Public Const GETPRINTINGOFFSET = 13
Public Const HORZSIZE As Long = 4
Public Const VERTSIZE As Long = 6
Public Const HORZRES As Long = 8
Public Const VERTRES As Long = 10
'Public Const LOGPIXELSX = 88
'Public Const LOGPIXELSY = 90


Public Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Public Const MM_LOMETRIC = 2 ' unité de mesure : le 1/10 mm - Each logical unit is mapped to 0.1 millimeter. Positive x is to the right; positive y is up.
Public Const MM_TWIPS = 6
'public Const MM_ANISOTROPIC as long 'Logical units are mapped to arbitrary units with arbitrarily scaled axes. Use the SetWindowExtEx and SetViewportExtEx functions to specify the units, orientation, and scaling.
'Public Const MM_HIENGLISH 'Each logical unit is mapped to 0.001 inch. Positive x is to the right; positive y is up.
'Public Const MM_HIMETRIC 'Each logical unit is mapped to 0.01 millimeter. Positive x is to the right; positive y is up.
'Public Const MM_ISOTROPIC 'Logical units are mapped to arbitrary units with equally scaled axes; that is, one unit along the x-axis is equal to one unit along the y-axis. Use the SetWindowExtEx and SetViewportExtEx functions to specify the units and the orientation of the axes. Graphics device interface (GDI) makes adjustments as necessary to ensure the x and y units remain the same size (When the window extent is set, the viewport will be adjusted to keep the units isotropic).
'Public Const MM_LOENGLISH 'Each logical unit is mapped to 0.01 inch. Positive x is to the right; positive y is up.
'Public Const MM_TEXT 'Each logical unit is mapped to one device pixel. Positive x is to the right; positive y is down.

Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Chord Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
   
Private Declare Sub ColorRGBToHLS Lib "shlwapi.dll" (ByVal lpRGB As Long, Hue As Long, Lum As Long, Sat As Long)
Private Declare Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal Hue As Long, ByVal Lum As Long, ByVal Sat As Long) As Long

Private Type TRIVERTEX
   x As Long
   y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type
Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Const GRADIENT_FILL_TRIANGLE = &H2&

'Public Enum GradientFillRectType
'   GRADIENT_FILL_RECT_H = 0
'   GRADIENT_FILL_RECT_V = 1
'End Enum

Public Const RGN_AND As Long = 1  'Zone commune aux deux régions
Public Const RGN_OR As Long = 2   'Réunion des deux régions.
Public Const RGN_XOR As Long = 3  'Zone qui n'est commune à aucune des deux régions.
Public Const RGN_DIFF As Long = 4 'Zone de la région 1 qui ne se trouve pas dans la région 2.
Public Const RGN_COPY As Long = 5 'Copie de la première région.
Public Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long

Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long
Private Type LOGBRUSH
lbStyle As Long
lbColor As Long
lbHatch As Long
End Type

Function Days(ByVal DayNo As Date) As Integer
  'used by function WeekNumber
  'http://support.microsoft.com/kb/200299
  Days = DayNo - DateSerial(Year(DayNo), 1, 0)
End Function

Function WeekNumber(ByVal InDate As Date) As Integer
  On Local Error Resume Next
  'http://support.microsoft.com/kb/200299
  Dim DayNo As Integer
  Dim StartDays As Integer
  Dim StopDays As Integer
  Dim StartDay As Integer
  Dim StopDay As Integer
  Dim VNumber As Integer
  Dim ThurFlag As Boolean

  DayNo = Days(InDate)
  StartDay = Weekday(DateSerial(Year(InDate), 1, 1), vbUseSystemDayOfWeek) '- 1
  StopDay = Weekday(DateSerial(Year(InDate), 12, 31), vbUseSystemDayOfWeek) '- 1
  ' Number of days belonging to first calendar week
  StartDays = 7 - (StartDay - 1)
  ' Number of days belonging to last calendar week
  StopDays = 7 - (StopDay - 1)
  ' Test to see if the year will have 53 weeks or not
  If StartDay = 4 Or StopDay = 4 Then ThurFlag = True Else ThurFlag = False
  VNumber = (DayNo - StartDays - 4) / 7
  ' If first week has 4 or more days, it will be calendar week 1
  ' If first week has less than 4 days, it will belong to last year's
  ' last calendar week
  If StartDays >= 4 Then
     WeekNumber = Fix(VNumber) + 2
  Else
     WeekNumber = Fix(VNumber) + 1
  End If
  ' Handle years whose last days will belong to coming year's first
  ' calendar week
  If WeekNumber > 52 And ThurFlag = False Then WeekNumber = 1
  ' Handle years whose first days will belong to the last year's
  ' last calendar week
  If WeekNumber = 0 Then
     WeekNumber = WeekNumber(DateSerial(Year(InDate) - 1, 12, 31))
  End If
End Function
'Private Sub GradientFillRect( _
'      ByVal lHDC As Long, _
'      tR As RECT, _
'      ByVal oStartColor As OLE_COLOR, _
'      ByVal oEndColor As OLE_COLOR, _
'      ByVal eDir As GradientFillRectType _
'   )
'Dim hBrush As Long
'Dim lStartColor As Long
'Dim lEndColor As Long
'Dim lR As Long
'
'   ' Use GradientFill:
''   If (HasGradientAndTransparency) Then
'      lStartColor = TranslateColor(oStartColor)
'      lEndColor = TranslateColor(oEndColor)
'
'      Dim tTV(0 To 1) As TRIVERTEX
'      Dim tGR As GRADIENT_RECT
'
'      setTriVertexColor tTV(0), lStartColor
'      tTV(0).x = tR.left
'      tTV(0).y = tR.top
'      setTriVertexColor tTV(1), lEndColor
'      tTV(1).x = tR.right
'      tTV(1).y = tR.bottom
'
'      tGR.UpperLeft = 0
'      tGR.LowerRight = 1
'
'      GradientFill lHDC, tTV(0), 2, tGR, 1, eDir
'
''   Else
''      ' Fill with solid brush:
''      hBrush = CreateSolidBrush(TranslateColor(oEndColor))
''      FillRect lHDC, tR, hBrush
''      DeleteObject hBrush
''   End If
'
'End Sub

Private Sub setTriVertexColor(tTV As TRIVERTEX, lColor As Long)
Dim lRed As Long
Dim lGreen As Long
Dim lBlue As Long
   lRed = (lColor And &HFF&) * &H100&
   lGreen = (lColor And &HFF00&)
   lBlue = (lColor And &HFF0000) \ &H100&
   setTriVertexColorComponent tTV.Red, lRed
   setTriVertexColorComponent tTV.Green, lGreen
   setTriVertexColorComponent tTV.Blue, lBlue
End Sub

Private Sub setTriVertexColorComponent(ByRef iColor As Integer, ByVal lComponent As Long)
   If (lComponent And &H8000&) = &H8000& Then
      iColor = (lComponent And &H7F00&)
      iColor = iColor Or &H8000
   Else
      iColor = lComponent
   End If
End Sub
Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function


Public Sub DrawLinePix(ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal StepXY As Boolean, _
                    ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long, ByVal Style As DrawStyleConstants, _
                    ByVal Epaisseur As Integer)

    Dim hPen      As Long
    Dim hOldPen   As Long
    Dim tLB As LOGBRUSH
    Call OleTranslateColor(Color, 0, Color)
    tLB.lbColor = Color
    
    X1 = ConvertTwipsToPixels(hdc, (X1 + glMgLeft) * md_Ratio, 0)
    Y1 = ConvertTwipsToPixels(hdc, (Y1 + glMgTop) * md_Ratio, 1)
    If Not StepXY Then
        X2 = X2 + glMgLeft
        Y2 = Y2 + glMgTop
    End If
    X2 = ConvertTwipsToPixels(hdc, X2 * md_Ratio, 0)
    Y2 = ConvertTwipsToPixels(hdc, Y2 * md_Ratio, 1)
    
    If Epaisseur > 1 Then
        hPen = ExtCreatePen(PS_GEOMETRIC Or PS_ENDCAP_ROUND Or Style, Epaisseur, tLB, 0, ByVal 0&)
    Else
        hPen = CreatePen(Style, Epaisseur, Color)
    End If
    hOldPen = SelectObject(hdc, hPen)

    MoveToEx hdc, X1, Y1, ByVal 0&
    If StepXY Then
        LineTo hdc, X1 + X2, Y1 + Y2
    Else
        LineTo hdc, X2, Y2
    End If
    SelectObject hdc, hOldPen
    DeleteObject hPen
    DeleteObject hOldPen
End Sub

'Public Sub DrawRectangle(ByRef hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal lWidth As Long, ByVal lHeight As Long, _
'                        ByVal lBackColor As Long, ByVal ColorLine As Long, ByVal Round As Long, ByVal lStyle As FillStyleConstants, _
'                        Optional ByVal BordGras As Integer = 0, Optional ByVal Degrad As eGradient = 0, Optional ByVal DoRadio As Boolean = True, _
'                        Optional ByVal degradC1 As OLE_COLOR = -1, Optional ByVal degradC2 As OLE_COLOR = -1, _
'                        Optional ByVal lBackColor2 As OLE_COLOR = -1)
'
'    '
'    'Degrad: 0 = non
'    '1= oui vers le clair vertical
'    '2=oui vers le fonce vertical
'    '3=oui vers le fonce horizontal
'
'    Dim hBrush    As Long
'    Dim hOldBrush As Long
'    Dim RetVal    As Long
'    Dim R         As Long
'    Dim G         As Long
'    Dim B         As Long
'    Dim hPen      As Long
'    Dim hOldPen   As Long
'    Dim lColor As Long
'    Dim lColorLine As Long
'    Dim tR     As RECT
'
'    If DoRadio Then
'        X1 = (X1 + glMgLeft) * md_Ratio
'        Y1 = (Y1 + glMgTop) * md_Ratio
'        lWidth = lWidth * md_Ratio
'        lHeight = lHeight * md_Ratio
'    End If
'    X1 = ConvertTwipsToPixels(hdc, X1, 0)
'    Y1 = ConvertTwipsToPixels(hdc, Y1, 1)
'    lWidth = ConvertTwipsToPixels(hdc, lWidth, 0)
'    lHeight = ConvertTwipsToPixels(hdc, lHeight, 1)
'
'    Call OleTranslateColor(lBackColor, 0, lColor)
'    Call OleTranslateColor(ColorLine, 0, lColorLine)
'
'    R = (lColor And &HFF&)
'    G = (lColor And &HFF00&) \ &H100&
'    B = (lColor And &HFF0000) \ &H10000
'    If lStyle = vbFSSolid Then
'        hBrush = CreateSolidBrush(RGB(R, G, B))
'    ElseIf lStyle = vbFSTransparent Then
'        hBrush = GetStockObject(SO_NULL_BRUSH)
'    Else
'        hBrush = CreateHatchBrush(lStyle - 2, RGB(R, G, B))
'        If lBackColor2 <> -1 Then
'            Dim hBrush2   As Long
'            Dim hOldBrush2 As Long
'            hBrush2 = CreateSolidBrush(lBackColor2)
'            hOldBrush2 = SelectObject(hdc, hBrush2)
'            RetVal = RoundRect(hdc, X1, Y1, X1 + lWidth, Y1 + lHeight, Round, Round)
'            RetVal = SelectObject(hdc, hOldBrush2)
'            RetVal = DeleteObject(hBrush2)
'        End If
'    End If
'    hOldBrush = SelectObject(hdc, hBrush)
'
'    R = (lColorLine And &HFF&)
'    G = (lColorLine And &HFF00&) \ &H100&
'    B = (lColorLine And &HFF0000) \ &H10000
'    hPen = CreatePen(PS_INSIDEFRAME, BordGras, RGB(R, G, B))
'    hOldPen = SelectObject(hdc, hPen)
'    SetBkColor hdc, RGB(R, G, B)
'
'    RetVal = RoundRect(hdc, X1, Y1, X1 + lWidth, Y1 + lHeight, Round, Round)
'    If Degrad > 0 And lStyle = vbFSSolid Then
'        If lColor > 0 Then
'            If degradC1 = -1 And degradC2 = -1 Then
'                'utilse pour drawsplit
'                With tR
'                    .left = X1
'                    .top = Y1 + 1
'                    .bottom = Y1 + lHeight
'                    .right = X1 + lWidth - 1
'                End With
'            Else
'                With tR
'                    .left = X1
'                    .top = Y1 + 1
'                    .bottom = Y1 + lHeight
'                    .right = X1 + lWidth
'                End With
'            End If
''            With tR
''                .left = X1
''                .top = Y1
''                .bottom = Y1 + lheight
''                .right = X1 + lWidth
''            End With
'            Select Case Degrad
'            Case eGd_VDarkToLight '1
'                If degradC1 = -1 Then degradC1 = Sombre(lColor, 20)
'                If degradC2 = -1 Then degradC2 = lColor
'                GradientFillRect hdc, tR, degradC1, degradC2, GRADIENT_FILL_RECT_V
'            Case eGd_VLightToDark '2
'                If degradC1 = -1 Then degradC1 = lColor
'                If degradC2 = -1 Then degradC2 = Sombre(lColor, 20)
'                GradientFillRect hdc, tR, degradC1, degradC2, GRADIENT_FILL_RECT_V
'            Case eGd_HLightToDark '3
'                GradientFillRect hdc, tR, lColor, Sombre(lColor, 20), GRADIENT_FILL_RECT_H
'            End Select
'        End If
'    End If
'    '
'    RetVal = SelectObject(hdc, hOldBrush)
'    RetVal = SelectObject(hdc, hOldPen)
'    RetVal = DeleteObject(hBrush)
'    RetVal = DeleteObject(hPen)
'End Sub

Public Function fctHeightText_fixeWidth(ByVal hdc As Long, ByVal lWidth As Long, ByVal Caption As String, ByVal lFontSize As Single, ByVal bBold As Boolean, _
                        ByVal bItalic As Boolean, ByVal sFont As String, Optional ByVal DoRadio As Boolean = True) As Long

    Dim lGras     As Long
    If bBold Then
        lGras = FW_BOLD
    Else
        lGras = FW_NORMAL
    End If
    
    
    Dim hFont     As Long
    Dim hOldFont  As Long
    Dim lSize As Long

    lSize = -MulDiv(lFontSize, GetDeviceCaps(hdc, LOGPIXELSY), 72)
    hFont = CreateFont(lSize, 0, 0, 0, lGras, bItalic, 0, _
            0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, _
            DEFAULT_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, sFont)

    hOldFont = SelectObject(hdc, hFont)
    
    If DoRadio Then
        lWidth = lWidth * md_Ratio
    End If
    
    Dim DrawingZone As RECT
    DrawingZone.left = 0
    DrawingZone.top = 0
    DrawingZone.right = ConvertTwipsToPixels(hdc, lWidth, 0)
    
    Dim lMemo As Long
    lMemo = DrawingZone.right
    
    'Calcule la taille a afficher (en coupant les mots et en remplaçant les tabulations par un certain nombre d'espaces)
    DrawText hdc, Caption, Len(Caption), DrawingZone, DT_WORDBREAK Or DT_CALCRECT

    fctHeightText_fixeWidth = DrawingZone.bottom - DrawingZone.top '- 50
    If fctHeightText_fixeWidth <= 0 Then fctHeightText_fixeWidth = 16
    fctHeightText_fixeWidth = ConvertPixelsToTwips(hdc, fctHeightText_fixeWidth, 1)
    

    SelectObject hdc, hOldFont
    DeleteObject hFont
End Function
Public Sub DrawCircle(ByRef hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal lRadius As Long, _
                        ByVal BackColor As Long, ByVal ColorLine As Long, ByVal lStyle As FillStyleConstants, _
                        Optional ByVal BordGras As Integer = 0, Optional ByVal DoRadio As Boolean = True)
    '
    Dim hBrush    As Long
    Dim hOldBrush As Long
    Dim RetVal    As Long
    Dim R As Long, G As Long, B As Long
    Dim hPen      As Long
    Dim hOldPen   As Long
    Dim lColor As Long
    Dim lColorLine As Long
    
    If DoRadio Then
        X1 = (X1 + glMgLeft) * md_Ratio
        Y1 = (Y1 + glMgTop) * md_Ratio
    End If
    X1 = ConvertTwipsToPixels(hdc, X1, 0)
    Y1 = ConvertTwipsToPixels(hdc, Y1, 1)

    Call OleTranslateColor(BackColor, 0, lColor)
    Call OleTranslateColor(ColorLine, 0, lColorLine)
    
    R = (lColor And &HFF&)
    G = (lColor And &HFF00&) \ &H100&
    B = (lColor And &HFF0000) \ &H10000
    If lStyle = vbFSSolid Then
        hBrush = CreateSolidBrush(RGB(R, G, B))
    Else
        hBrush = CreateHatchBrush(lStyle - 2, RGB(R, G, B))
    End If
    hOldBrush = SelectObject(hdc, hBrush)
    
    R = (lColorLine And &HFF&)
    G = (lColorLine And &HFF00&) \ &H100&
    B = (lColorLine And &HFF0000) \ &H10000
    hPen = CreatePen(PS_INSIDEFRAME, BordGras, RGB(R, G, B))
    hOldPen = SelectObject(hdc, hPen)
    '
    RetVal = Chord(hdc, X1 - lRadius / 2, Y1 - lRadius / 2, X1 + lRadius / 2, Y1 + lRadius / 2, X1 + lRadius / 2, Y1, X1 + lRadius / 2, Y1)
    '
    RetVal = SelectObject(hdc, hOldBrush)
    RetVal = SelectObject(hdc, hOldPen)
    RetVal = DeleteObject(hBrush)
    RetVal = DeleteObject(hPen)
End Sub
Public Function ConvertTwipsToPixels(ByVal lHDC As Long, ByVal lTwips As Long, ByVal lDirection As Long) As Long
    ' http://support.microsoft.com/kb/210590/fr
    Dim lPixelsPerInch As Long
    
    If (lDirection = 0) Then
        lPixelsPerInch = GetDeviceCaps(lHDC, LOGPIXELSX)
    Else
        lPixelsPerInch = GetDeviceCaps(lHDC, LOGPIXELSY)
    End If

    ConvertTwipsToPixels = lTwips / 1440 * lPixelsPerInch
End Function

Public Function ConvertPixelsToTwips(ByVal lHDC As Long, ByVal lPixels As Long, ByVal lDirection As Long) As Long
    ' http://support.microsoft.com/kb/152475/fr
    Dim lPixelsPerInch As Long
    
    If (lDirection = 0) Then
       lPixelsPerInch = GetDeviceCaps(lHDC, LOGPIXELSX)
    Else
       lPixelsPerInch = GetDeviceCaps(lHDC, LOGPIXELSY)
    End If
    
    ConvertPixelsToTwips = (lPixels / lPixelsPerInch) * 1440
End Function
Public Sub DrawTriangle(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal Direction As AlignConstants, ByVal lCouleur As Long, ByVal Distance As Integer)
    Dim hBrush    As Long
    Dim hOldBrush As Long
    Dim hPen      As Long
    Dim hOldPen   As Long
    Dim RetVal    As Long
    Dim R         As Long
    Dim G         As Long
    Dim B         As Long
    Dim P(3)      As POINTAPI
    '
    x = ConvertTwipsToPixels(hdc, (x + glMgLeft) * md_Ratio, 0)
    y = ConvertTwipsToPixels(hdc, (y + glMgTop) * md_Ratio, 1)
    Distance = ConvertTwipsToPixels(hdc, Distance * md_Ratio, 0)
    '
    R = (lCouleur And &HFF&)
    G = (lCouleur And &HFF00&) \ &H100&
    B = (lCouleur And &HFF0000) \ &H10000
    hBrush = CreateSolidBrush(RGB(R, G, B))
    hOldBrush = SelectObject(hdc, hBrush)
    '
    R = (lCouleur And &HFF&)
    G = (lCouleur And &HFF00&) \ &H100&
    B = (lCouleur And &HFF0000) \ &H10000
    hPen = CreatePen(6, 0, RGB(R, G, B))
    hOldPen = SelectObject(hdc, hPen)
    '
'    x = x + glMgLeft
    '
    Select Case Direction
    Case vbAlignLeft
        P(0).x = x - Distance
        P(0).y = y
        P(1).x = x - 2 * Distance
        P(1).y = y - Distance
        P(2).x = x - 2 * Distance
        P(2).y = y + Distance
    Case vbAlignRight
        P(0).x = x + Distance
        P(0).y = y
        P(1).x = x + 2 * Distance
        P(1).y = y - Distance
        P(2).x = x + 2 * Distance
        P(2).y = y + Distance
    Case vbAlignTop
        P(0).x = x - Distance
        P(0).y = y - Distance
        P(2).x = x + Distance
        P(2).y = y - Distance
        P(1).x = x
        P(1).y = y
    Case vbAlignBottom
        P(0).x = x - Distance
        P(0).y = y + Distance
        P(2).x = x + Distance
        P(2).y = y + Distance
        P(1).x = x
        P(1).y = y
    Case 5 'haut gauche
        P(0).x = x
        P(0).y = y
        P(2).x = x + Distance
        P(2).y = y
        P(1).x = x
        P(1).y = y + Distance
    Case 6 'haut droit
        P(0).x = x - Distance
        P(0).y = y
        P(2).x = x
        P(2).y = y
        P(1).x = x
        P(1).y = y + Distance
    Case 7 '-Bas droit
        P(0).x = x
        P(0).y = y
        P(1).x = x
        P(1).y = y - Distance
        P(2).x = x - Distance
        P(2).y = y
    End Select
    '
    Polygon hdc, P(0), 3
    '
    RetVal = SelectObject(hdc, hOldBrush)
    RetVal = SelectObject(hdc, hOldPen)
    RetVal = DeleteObject(hBrush)
    RetVal = DeleteObject(hPen)
End Sub

Public Sub DrawPolyPx(ByRef hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal lWidth As Long, ByVal lHeight As Long, _
                        ByVal BackColor As Long, ByVal ColorLine As Long, ByVal lStyle As FillStyleConstants, _
                        Optional ByVal BordGras As Integer = 0, Optional ByVal Fleche As Long)
    Dim hBrush    As Long
    Dim hOldBrush As Long
    Dim hPen      As Long
    Dim hOldPen   As Long
    Dim RetVal    As Long
    Dim R         As Long
    Dim G         As Long
    Dim B         As Long
    Dim P(5)      As POINTAPI
    '
    'If DoRadio Then
        X1 = (X1 + glMgLeft) * md_Ratio
        Y1 = (Y1 + glMgTop) * md_Ratio
        lWidth = lWidth * md_Ratio
        lHeight = lHeight * md_Ratio
    'End If
    
    X1 = ConvertTwipsToPixels(hdc, X1, 0)
    Y1 = ConvertTwipsToPixels(hdc, Y1, 1)
    lWidth = ConvertTwipsToPixels(hdc, lWidth, 0)
    lHeight = ConvertTwipsToPixels(hdc, lHeight, 1)
    Fleche = ConvertTwipsToPixels(hdc, Fleche * md_Ratio, 0)
    '
    R = (BackColor And &HFF&)
    G = (BackColor And &HFF00&) \ &H100&
    B = (BackColor And &HFF0000) \ &H10000
    If lStyle = vbFSSolid Then
        hBrush = CreateSolidBrush(RGB(R, G, B))
    Else
        hBrush = CreateHatchBrush(lStyle - 2, RGB(R, G, B))
    End If
    hOldBrush = SelectObject(hdc, hBrush)
    '
    R = (ColorLine And &HFF&)
    G = (ColorLine And &HFF00&) \ &H100&
    B = (ColorLine And &HFF0000) \ &H10000
    hPen = CreatePen(6, BordGras, RGB(R, G, B))
    hOldPen = SelectObject(hdc, hPen)
    '
    If Fleche * 2 > lWidth Then
        Fleche = lWidth / 2
    End If
    P(0).x = X1 + lWidth - Fleche
    P(0).y = Y1 + lHeight
    P(1).x = X1 + lWidth
    P(1).y = Y1 + lHeight / 2
    P(2).x = X1 + lWidth - Fleche
    P(2).y = Y1
    P(3).x = X1 + Fleche
    P(3).y = Y1
    P(4).x = X1
    P(4).y = Y1 + lHeight / 2
    P(5).x = X1 + Fleche
    P(5).y = Y1 + lHeight
    '
    Polygon hdc, P(0), 6
    '
    RetVal = SelectObject(hdc, hOldBrush)
    RetVal = SelectObject(hdc, hOldPen)
    RetVal = DeleteObject(hBrush)
    RetVal = DeleteObject(hPen)
End Sub

Public Function Sombre(ByVal Color As Long, Optional ByVal Luminance As Long = 50, Optional ByVal Inverse As Boolean = False) As Long
    On Error Resume Next
    
    Dim H As Long
    Dim L As Long
    Dim S As Long
    
    ColorRGBToHLS Color, H, L, S
    
    L = L - Luminance
    If L < 1 Then L = 1
    If Inverse Then If L < 30 Then L = 200
    If L > 240 Then L = 240
    
    Sombre = ColorHLSToRGB(H, L, S)

'    R = (Sombre And &HFF&)
'    G = (Sombre And &HFF00&) \ &H100&
'    B = (Sombre And &HFF0000) \ &H10000
End Function


