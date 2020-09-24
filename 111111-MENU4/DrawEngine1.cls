VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DrawEngine1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private FONTX As Long
Private BCKGND As Long
Private FORGND As Long
Private HGHLT As Long
Private SEPCLR As Long
Private DCLR As Long
Private NTF As Long
Private SEPH As Long
Private EDG As HighlightEdge
Private REDG As Long
Private TXTA As TextAlign




Private HGB As Long
Private BCB As Long
Private SEP As Long
Private SNDN As Long

Public Enum FontWeights
FW_DONTCARE = 0
FW_THIN = 100
FW_EXTRALIGHT = 200
FW_LIGHT = 300
FW_NORMAL = 400
FW_MEDIUM = 500
FW_SEMIBOLD = 600
FW_BOLD = 700
FW_EXTRABOLD = 800
FW_HEAVY = 900
End Enum

Public Enum HighlightEdge
[Raised]
[Sunken]
[Etched]
[Bump]
[RaisedOuter]
[RaisedInner]
[SunkenOuter]
[SunkenInner]
End Enum

Public Enum TextAlign
[LeftAlignment]
[CenterAlignment]
[RightAlignment]
End Enum

Public Sub MenuFont(ByVal FontName As String, ByVal FontSize As Integer, ByVal FontWeight As FontWeights, ByVal FontCharset As Long, ByVal FontItalic As Boolean, ByVal FontUnderline As Boolean, ByVal FontStrikethrough As Boolean)
Dim TMPFNT As Long
TMPFNT = GetFont(FontName, FontSize, FontCharset, FontItalic, Fontundeline, FontStrikethrough, FontWeight)
If TMPFNT = 0 Then Exit Sub
If (FONTX <> 0) Then DeleteObject FONTX
FONTX = TMPFNT
End Sub
Public Property Let Background(ByVal color As Long)
BCKGND = color
DeleteObject BCB
BCB = CreateSolidBrush(color)
End Property
Public Property Get Background() As Long
Background = BCKGND
End Property
Public Property Let ForegroundSelected(ByVal color As Long)
FORGND = color
End Property
Public Property Get ForegroundSelected() As Long
ForegroundSelected = FORGND
End Property
Public Property Let ForegroundNotSelected(ByVal color As Long)
NTF = color
End Property
Public Property Get ForegroundNotSelected() As Long
ForegroundNotSelected = NTF
End Property
Public Property Let HighlightColor(ByVal color As Long)
HGHLT = color
DeleteObject HGB
HGB = CreateSolidBrush(color)
End Property
Public Property Get HighlightColor() As Long
HighlightColor = HGHLT
End Property
Public Property Let SeparatorColor(ByVal color As Long)
SEPCLR = color
DeleteObject SEP
SEP = CreateSolidBrush(color)
End Property
Public Property Get SeparatorColor() As Long
SeparatorColor = SEPCLR
End Property
Public Property Let DisabledColor(ByVal color As Long)
DCLR = color
End Property
Public Property Get DisabledColor() As Long
DisabledColor = DCLR
End Property
Public Property Let SeparatorHeight(ByVal height As Long)
SEPH = height
End Property
Public Property Get SeparatorHeight() As Long
SeparatorHeight = SEPH
End Property
Public Property Let SelectedEdge(ByVal edgeprop As HighlightEdge)
EDG = edgeprop
Select Case EDG
Case [Raised]
REDG = EDGE_RAISED
Case [Sunken]
REDG = EDGE_SUNKEN
Case [Etched]
REDG = EDGE_ETCHED
Case [Bump]
REDG = EDGE_BUMP
Case [RaisedOuter]
REDG = BDR_RAISEDOUTER
Case [RaisedInner]
REDG = BDR_RAISEDINNER
Case [SunkenOuter]
REDG = BDR_SUNKENOUTER
Case [SunkenInner]
REDG = BDR_SUNKENINNER
End Select
End Property
Public Property Get SelectedEdge() As HighlightEdge
SelectedEdge = EDG
End Property
Public Property Let TextAlignment(ByVal alignprop As TextAlign)
TXTA = alignprop
End Property
Public Property Get TextAlignment() As TextAlign
TextAlignment = TXTA
End Property


Public Property Let Sound(ByVal soundnumber As Long)
If soundnumber < 0 Or soundnumber > 5 Then Exit Property
Dim tmn As Long
SNDN = soundnumber
tmn = soundnumber + 100
SND1 = LoadResData(tmn, "CUSTOM")
End Property
Public Property Get Sound() As Long
Sound = SNDN
End Property
Public Sub ProcessDrawItem(drawIstruct As MenuObject.DRAWITEMSTRUCT, ByVal NewMenuText As String, ByVal ParentHwnd As Long)

SetBkMode drawIstruct.hdc, 1
Dim Sbr As Long
Dim RCTX As RECT
Dim SelectX As Boolean
Dim DisabledX As Boolean

CopyMemory RCTX, drawIstruct.rcItem, Len(RCTX)
CheckStates drawIstruct, DisabledX, SelectX

FillRect drawIstruct.hdc, RCTX, BCB

If SelectX And Not DisabledX Then
FillRect drawIstruct.hdc, RCTX, HGB
Call DrawEdge(drawIstruct.hdc, RCTX, REDG, BF_RECT)
OffsetRect RCTX, -2, -2

If Not (Not CBool(SNDN)) Then
PlaySound_Res ByVal VarPtr(SND1(0)), 0, &H4 Or &H1
End If

End If

SelectObject drawIstruct.hdc, FONTX
Dim DTP As DRAWTEXTPARAMS
DTP.cbSize = Len(DTP)

If (drawIstruct.ItemID = 0) And (NewMenuText = "") Then
'SEPARATOR
FillRect drawIstruct.hdc, drawIstruct.rcItem, SEP
End If

If DisabledX Then
SetTextColor drawIstruct.hdc, DCLR
ElseIf SelectX Then
SetTextColor drawIstruct.hdc, FORGND
ElseIf Not SelectX Then
SetTextColor drawIstruct.hdc, NTF
End If


If TXTA = LeftAlignment Then
RCTX.Left = RCTX.Left + 3
ElseIf TXTA = RightAlignment Then
RCTX.Right = RCTX.Right - 3
End If

Call DrawTextEx(drawIstruct.hdc, NewMenuText, Len(NewMenuText), RCTX, DT_VCENTER Or TXTA Or DT_SINGLELINE, DTP)

End Sub
Public Sub ProcessMeasureItem(measureIstruct As MenuObject.MEASUREITEMSTRUCT, ByVal NewMenuText As String, ByVal ParentHwnd As Long)
Dim SZ1 As SIZE
Dim DCX As Long
Dim OF As Long
DCX = GetDC(ParentHwnd)
OF = SelectObject(DCX, FONTX)
GetTextExtentPoint32 DCX, NewMenuText, Len(NewMenuText), SZ1
measureIstruct.ItemHeight = SZ1.cy + 8
measureIstruct.itemWidth = SZ1.cx + 4
If (NewMenuText = "") And (measureIstruct.ItemID = 0) Then measureIstruct.ItemHeight = SEPH
SelectObject DCX, OF
ReleaseDC ParentHwnd, DCX
End Sub

Private Sub Class_Terminate()
DeleteObject FONTX
DeleteObject BCB
DeleteObject HGB
DeleteObject SEP
DeleteDC DCX
Erase SND1

End Sub
