VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DrawEngine2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum IconsSize
 [Size16x16] = 16
 [Size32x32] = 32
End Enum

Public Enum HighlightColor
[BlueGrad]
[RedGrad]
[GreenGrad]
[GrayGrad]
End Enum



Private FRMCLR As Long
Private TOCLR As Long

Private SZI As Long
Private FONTX As Long
Private WIRH As Long
Private Icons() As Long
Private IconsIndex() As Long
Private ICcount As Long
Public Property Let Icon(ByVal ItemID As Long, ByVal IconHandle As Long)
Dim plc As Long
If ICcount = 0 Then GoTo daljeXX
plc = FindPlace(IconsIndex, ItemID, True)
daljeXX:
AddEntry ItemID, plc, IconsIndex
AddEntry IconHandle, plc, Icons
ICcount = ICcount + 1
End Property
Public Property Get Icon(ByVal ItemID As Long) As Long
If ICcount = 0 Then Icon = -1: Exit Property
Dim plc As Long
plc = FindPlace(IconsIndex, ItemID, False)
If plc = -1 Then Icon = -1: Exit Property
Icon = Icons(plc)
End Property
Public Property Let IconSize(ByVal IconSz As IconsSize)
SZI = IconSz
End Property
Public Property Get IconSize() As IconsSize
IconSize = SZI
End Property
Public Property Let SelectedColor(ByVal color As HighlightColor)
Select Case color
Case [BlueGrad]
FRMCLR = &H3200000
TOCLR = &HFF0000

Case [RedGrad]
FRMCLR = &H32&
TOCLR = &HFF&

Case [GreenGrad]
FRMCLR = &H3200&
TOCLR = &HFF00&

Case [GrayGrad]
FRMCLR = &HBBBBBB
TOCLR = &H111111

End Select
End Property



Public Sub EraseIcons()
Erase Icons
Erase IconsIndex
ICcount = 0
End Sub
Public Sub DeleteIcon(ByVal ItemID As Long)
If ICcount = 0 Then Exit Sub
Dim plc As Long
plc = FindPlace(IconsIndex, ItemID, False)
If plc = -1 Then Exit Sub
DelEntry plc, IconsIndex
DelEntry plc, Icons
ICcount = ICcount - 1
End Sub

Public Property Let MenuFont(ByVal IconSz As IconsSize, ByVal FontName As String)
Dim TMPFNT As Long
Dim Fsz As Integer
If IconSz = Size32x32 Then
Fsz = 10
Else
Fsz = 8
End If
TMPFNT = GetFont(FontName, Fsz, 238, False, False, False, FW_NORMAL)
If TMPFNT = 0 Then Exit Property
If (FONTX <> 0) Then DeleteObject FONTX
FONTX = TMPFNT
End Property
Public Sub ProcessMeasureItem(measureIstruct As MenuObject.MEASUREITEMSTRUCT, ByVal NewMenuText As String, ByVal ParentHwnd As Long)
Dim SZ1 As SIZE
Dim DCX As Long
Dim OF As Long
DCX = GetDC(ParentHwnd)
OF = SelectObject(DCX, FONTX)
GetTextExtentPoint32 DCX, NewMenuText, Len(NewMenuText), SZ1
WIRH = measureIstruct.ItemHeight
measureIstruct.ItemHeight = SZI + 4
measureIstruct.itemWidth = SZ1.cx + SZI + 4
If (NewMenuText = "") And (measureIstruct.ItemID = 0) Then measureIstruct.ItemHeight = 2
SelectObject DCX, OF
ReleaseDC ParentHwnd, DCX
End Sub
Public Sub ProcessDrawItem(drawIstruct As MenuObject.DRAWITEMSTRUCT, ByVal NewMenuText As String, ByVal ParentHwnd As Long)
Dim RCTX As RECT
Dim RCTX2 As RECT
Dim DTP As DRAWTEXTPARAMS
DTP.cbSize = Len(DTP)

Dim SelectX As Boolean
Dim DisabledX As Boolean
CheckStates drawIstruct, DisabledX, SelectX


Dim orR As RECT
orR.Right = drawIstruct.rcItem.Right - drawIstruct.rcItem.Left
orR.Bottom = drawIstruct.rcItem.Bottom - drawIstruct.rcItem.Top

Dim DSK As Long
Dim DPC As Long
DSK = GetDesktopWindow
DPC = GetDC(DSK)

Dim MEMDC As Long
Dim BMPX As Long
MEMDC = CreateCompatibleDC(DPC)
BMPX = CreateCompatibleBitmap(DPC, orR.Right, orR.Bottom)
SelectObject MEMDC, BMPX


CopyMemory RCTX, orR, Len(RCTX)
CopyMemory RCTX2, RCTX, Len(RCTX)

SetBkMode MEMDC, 1
Dim BRX1 As Long
BRX1 = CreateSolidBrush(&HEEEEEE)

FillRect MEMDC, orR, BRX1

Dim ICO2 As Long
Dim SLL As Long
Dim PTRN As Long

If SelectX Then
SLL = 1
End If

ICO2 = Icon(drawIstruct.ItemID)

Dim RCTX3 As RECT
RCTX3.Top = 0
RCTX3.Left = 0
RCTX3.Right = SZI + 6
RCTX3.Bottom = SZI + 4
Dim GRY As Long
GRY = CreateSolidBrush(&HBBBBBB)
FillRect MEMDC, RCTX3, GRY
DeleteObject GRY


If SelectX And Not DisabledX Then
PTRN = DSS_MONO
ElseIf DisabledX Then
PTRN = DSS_DISABLED
End If

If (drawIstruct.ItemID = 0) And (NewMenuText = "") Then
'SEPARATOR
DrawEdge MEMDC, RCTX, EDGE_ETCHED, BF_RECT
Else
Call DrawState(MEMDC, 0, 0, ICO2, 0, _
                RCTX2.Left + 2, RCTX2.Top + 2, SZI, SZI, DST_ICON Or PTRN)
End If

If Not DisabledX And ICO2 <> -1 Then
Call DrawIconEx(MEMDC, RCTX2.Left + 2 + SLL, RCTX2.Top + 2 + SLL, ICO2, SZI, SZI, 0, 0, DI_NORMAL)
End If

SetTextColor MEMDC, &H0&
SelectObject MEMDC, FONTX

RCTX.Left = SZI + 6

If SelectX And Not DisabledX Then
FillGRAD MEMDC, RCTX, FRMCLR, TOCLR, True

If ICO2 <> -1 Then
RCTX3.Top = 0
RCTX3.Left = 0
Call DrawEdge(MEMDC, RCTX3, BDR_RAISEDOUTER, BF_RECT)
End If
SetTextColor MEMDC, &HEEEEEE
End If

RCTX.Left = SZI + 10

Dim h1x As Long
h1x = ((RCTX.Bottom - RCTX.Top) / 2) - (WIRH / 2)
If Not SelectX Or DisabledX Then
Call DrawStateText(MEMDC, 0&, 0&, NewMenuText, Len(NewMenuText), _
               RCTX.Left, h1x, 0&, 0&, DST_PREFIXTEXT Or DSS_DISABLED)
End If

If Not DisabledX Then
Call DrawTextEx(MEMDC, NewMenuText, Len(NewMenuText), RCTX, DT_VCENTER Or DT_SINGLELINE, DTP)
End If


BitBlt drawIstruct.hdc, drawIstruct.rcItem.Left, drawIstruct.rcItem.Top, _
drawIstruct.rcItem.Right - drawIstruct.rcItem.Left, _
drawIstruct.rcItem.Bottom - drawIstruct.rcItem.Top, MEMDC, 0, 0, SRCCOPY

DeleteObject BRX1
ReleaseDC DSK, DPC
DeleteDC MEMDC
DeleteObject BMPX


End Sub





Private Sub Class_Terminate()
EraseIcons
End Sub
