VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private CCOUNT As Long
Private InternalData() As Byte
Private Menu1 As Long
Private MenuCMP As String
Public Property Let MenuData(ByVal data As String)
data = Replace(data, Chr(13) & Chr(10), "")
MenuCMP = data
End Property
Public Property Get MenuData() As String
MenuData = MenuCMP
End Property
Public Property Get Handle() As Long
Handle = Menu1
End Property
Public Sub SaveBinaryMenu(ByVal filename As String)
If Not CBool(Menu1) Then Exit Sub
If Dir(filename) <> "" Then Kill filename
Dim FFL As Long
FFL = FreeFile
Open filename For Binary As #FFL
Put #FFL, , InternalData
Close #1
End Sub

Public Sub KillMenu()
Menu1 = 0
DestroyMenu Menu1
Erase InternalData
End Sub


Public Function CompileMenu() As Long
On Error GoTo eend
If MenuCMP = "" Then Exit Function
If Menu1 <> 0 Then DestroyMenu Menu1
Erase InternalData
Dim POPS() As String

Dim TmpData() As Byte

POPS = Split(MenuCMP, "/", , vbBinaryCompare)

Dim INTDATA As New Collection
Dim WKind As String

Dim SP() As String
Dim ORX As Byte

CCOUNT = 0

For u = 1 To UBound(POPS)
WKind = Left(POPS(u), 4)
SP = Split(POPS(u), "::")

Select Case UCase(WKind)

Case "NPOP"
PutPOP SP, False, INTDATA

Case "LPOP"
PutPOP SP, True, INTDATA

Case "NITM"
PutITM SP, False, INTDATA

Case "LITM"
PutITM SP, True, INTDATA

Case "NSEP"
PutSeparator INTDATA, False, SP

Case "LSEP"
PutSeparator INTDATA, True, SP

Case Else
GoTo eend

End Select
Next u
FusionData INTDATA, InternalData
Menu1 = LoadMenuIndirect(ByVal VarPtr(InternalData(0)))
CompileMenu = Menu1
Exit Function
eend:
If Err <> 0 Then On Error GoTo 0
Erase InternalData
Menu1 = 0
End Function
Private Sub PutPOP(SP() As String, ByVal tf As Boolean, INTDATA As Collection)
If UBound(SP) > 1 Then
PutPopUp INTDATA, SP(1), tf, LongToInt(GetORX(SP(2)))
Else
PutPopUp INTDATA, SP(1), tf
End If
End Sub
Private Sub PutITM(SP() As String, ByVal tf As Boolean, INTDATA As Collection)
If UBound(SP) > 2 Then
PutItemUp INTDATA, SP(1), CLng(SP(2)), tf, LongToInt(GetORX(SP(3)))
Else
PutItemUp INTDATA, SP(1), CLng(SP(2)), tf
End If
End Sub


Private Function GetORX(ByVal stringX As String) As Long
On Error GoTo eendX1
Dim TS() As String
stringX = UCase(stringX)
TS = Split(stringX, "OR")
For u = 0 To UBound(TS)
TS(u) = Replace(TS(u), " ", "")
Select Case TS(u)
Case "CHECKED"
GetORX = GetORX Or &H8
Case "BREAK"
GetORX = GetORX Or &H40
Case "BARBREAK"
GetORX = GetORX Or &H20
Case "DISABLED"
GetORX = GetORX Or &H2
Case "GRAYED"
GetORX = GetORX Or 1
Case "OWNERDRAW"
GetORX = GetORX Or &H100
'Case "HELP"
'GetORX = GetORX Or &H4000&
End Select
Next u
eendX1:
If Err <> 0 Then On Error GoTo 0
End Function


Private Sub FusionData(col As Collection, data() As Byte)
Dim tmpDT() As Byte
Dim countXX As Long
ReDim data(CCOUNT + 3)
countXX = 4
For u = 1 To col.Count
tmpDT = col.Item(u)
CopyMemory data(countXX), tmpDT(0), UBound(tmpDT) + 1
countXX = countXX + UBound(tmpDT) + 1
Next u
End Sub



'********POPUP-- 10 00 CAPTION
'********ZADNJI POPUP-- 90 00 CAPTION

'********ITEM--00 00 XX XX CAPTION  (XX XX- ID)
'********ZADNJI ITEM-- 80 00 XX XX CAPTION (XX XX- ID)

'********SEPARATOR-- 00 08 00 00 00 00
'********ZADNJI SEPARATOR-- 80 08 00 00 00 00

'****GRAYED---&H01
'****DISABLED--&H02
'****CHECKED--&H08
'****BARBREAK--&H20
'****BREAK--&H40
'****OWNERDRAW-&H100


Private Sub PutPopUp(col As Collection, ByVal caption As String, ByVal last As Boolean, Optional ByVal Ored As Integer)
Dim DT() As Byte
Dim INTX As Integer
ReDim DT(LenB(caption) + 3)
If last Then
INTX = &H90& Or Ored
Else
INTX = &H10& Or Ored
End If
CopyMemory DT(0), INTX, 2
CopyMemory DT(2), ByVal StrPtr(caption), LenB(caption)
col.Add DT
CCOUNT = CCOUNT + UBound(DT) + 1
End Sub

Private Sub PutItemUp(col As Collection, ByVal caption As String, ByVal ID As Long, ByVal last As Boolean, Optional ByVal Ored As Integer)
Dim DT() As Byte
Dim INTX As Integer
ReDim DT(LenB(caption) + 5)
If last Then
INTX = &H80& Or Ored
Else
INTX = Ored
End If
CopyMemory DT(0), INTX, 2
CopyMemory DT(2), ByVal VarPtr(ID), 2
CopyMemory DT(4), ByVal StrPtr(caption), LenB(caption)
CCOUNT = CCOUNT + UBound(DT) + 1
col.Add DT
End Sub
Private Sub PutSeparator(col As Collection, ByVal last As Boolean, SP() As String)
Dim XORX As Byte
If UBound(SP) > 0 Then
If (GetORX(SP(1)) And &H100&) = &H100& Then XORX = 1
End If
Dim DT() As Byte
ReDim DT(5)
If last Then DT(0) = &H80&
DT(1) = &H8& Or XORX
col.Add DT
CCOUNT = CCOUNT + UBound(DT) + 1
End Sub

Private Sub Class_Terminate()
KillMenu
Erase InternalData
End Sub
