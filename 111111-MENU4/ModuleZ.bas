Attribute VB_Name = "ModuleZ"
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)


Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long

Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020


Public Const ODS_SELECTED = &H1
Public Const ODS_GRAYED = &H2
Public Const ODS_DISABLED = &H4
Public Const ODS_CHECKED = &H8
Public Const ODS_FOCUS = &H10
Public Const ODS_DEFAULT = &H20
Public Const ODS_HOTLIGHT = &H40
Public Const ODS_INACTIVE = &H80
Public Const ODS_NOACCEL = &H100
Public Const ODS_NOFOCUSRECT = &H200

Public Const DT_TOP = &H0
 Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000
Public Const DT_EDITCONTROL = &H2000
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_RTLREADING = &H20000
Public Const DT_WORD_ELLIPSIS = &H40000
 Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
 Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Const LOGPIXELSY = 90


Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long

Public Type SIZE
        cx As Long
        cy As Long
End Type

Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As SIZE) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Public SND1() As Byte

 Declare Function PlaySound_Res Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As Long, ByVal hModule As Long, ByVal dwFlags As Long) As Long
 Declare Function GetObj Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
 Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
 Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
 Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function LoadImage Lib "user32" Alias "LoadImageA" _
    (ByVal hInst As Long, ByVal lpsz As String, _
    ByVal iType As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal fOptions As Long) As Long
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2

Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const DI_NORMAL = &H3

Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DrawStateText Lib "user32" Alias "DrawStateA" _
        (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc _
        As Long, ByVal lParam As String, ByVal wParam As Long, _
        ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, _
        ByVal n4 As Long, ByVal un As Long) As Long
' DrawState API Constants
Public Const DST_COMPLEX = &H0
Public Const DST_TEXT = &H1
Public Const DST_PREFIXTEXT = &H2
Public Const DST_ICON = &H3
Public Const DST_BITMAP = &H4
Public Const DSS_NORMAL = &H0
Public Const DSS_UNION = &H10 ' // Grey Text
Public Const DSS_DISABLED = &H20
Public Const DSS_MONO = &H80
Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Function GetFont(ByVal nameX As String, nSize As Integer, ByVal charset As Long, ByVal italicX As Boolean, ByVal underlineX As Boolean, ByVal strikeX As Boolean, ByVal fontW As FontWeights) As Long
Dim DCX As Long
Dim FB As Long
DCX = GetDC(0)
GetFont = CreateFont(-MulDiv(nSize, GetDeviceCaps(DCX, LOGPIXELSY), 72), 0, 0, 0, fontW, italicX, underlineX, strikeX, charset, 0, 0, 2, 1, nameX)
ReleaseDC 0, DCX
End Function

Public Sub FillGRAD(ByVal hdc As Long, RCX As RECT, ByVal ColorFrom As Long, ByVal ColorTo As Long, ByVal HorVer As Boolean)
'TRUE-vertical
'FALSE-horizontal

Dim WIDTHX As Long
Dim HEIGHTX As Long
WIDTHX = RCX.Right - RCX.Left
HEIGHTX = RCX.Bottom - RCX.Top
If WIDTHX = 0 Or HEIGHTX = 0 Then Exit Sub
Dim CLRS() As Long
If HorVer Then
ReDim CLRS(WIDTHX - 1)
Else
ReDim CLRS(HEIGHTX - 1)
End If
GradateColors CLRS, ColorFrom, ColorTo
Dim JK As SIZE
Dim OOB As Long
If HorVer Then
GoSub vert
Else
GoSub horz
End If
Exit Sub

horz:
For i2 = 0 To HEIGHTX - 1
hbr = CreatePen(0, 1, CLRS(i2))
OOB = SelectObject(hdc, hbr)
MoveToEx hdc, RCX.Left, RCX.Top + i2, JK
LineTo hdc, RCX.Right, RCX.Top + i2
SelectObject hdc, OOB
DeleteObject hbr
Next
Return

vert:
For i = 0 To WIDTHX - 1
hbr = CreatePen(0, 1, CLRS(i))
OOB = SelectObject(hdc, hbr)
MoveToEx hdc, RCX.Left + i, RCX.Top, JK
LineTo hdc, RCX.Left + i, RCX.Bottom
SelectObject hdc, OOB
DeleteObject hbr
Next
Return

End Sub


Private Sub GradateColors(Colors() As Long, ByVal Color1 As Long, _
    ByVal Color2 As Long)

    Dim i As Long
    Dim dblR As Double, dblG As Double, dblB As Double
    Dim addR As Double, addG As Double, addB As Double
    Dim bckR As Double, bckG As Double, bckB As Double

    dblR = CDbl(Color1 And &HFF)
    dblG = CDbl(Color1 And &HFF00&) / 255
    dblB = CDbl(Color1 And &HFF0000) / &HFF00&
    bckR = CDbl(Color2 And &HFF&)
    bckG = CDbl(Color2 And &HFF00&) / 255
    bckB = CDbl(Color2 And &HFF0000) / &HFF00&

    addR = (bckR - dblR) / UBound(Colors)
    addG = (bckG - dblG) / UBound(Colors)
    addB = (bckB - dblB) / UBound(Colors)

    For i = 0 To UBound(Colors)
        dblR = dblR + addR
        dblG = dblG + addG
        dblB = dblB + addB
        If dblR > 255 Then dblR = 255
        If dblG > 255 Then dblG = 255
        If dblB > 255 Then dblB = 255
        If dblR < 0 Then dblR = 0
        If dblG < 0 Then dblG = 0
        If dblB < 0 Then dblB = 0
        Colors(i) = RGB(dblR, dblG, dblB)
    Next

End Sub
