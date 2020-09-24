Attribute VB_Name = "ModuleMSG"

Public Sub CheckStates(drawIstruct As DRAWITEMSTRUCT, DisabledX As Boolean, SelectX As Boolean)
If ((drawIstruct.itemState And ODS_SELECTED) = ODS_SELECTED) Or ((drawIstruct.itemState And 257) = 257) Then SelectX = True
'Win2000 & XP
If (drawIstruct.itemState And (ODS_GRAYED + 256)) = (ODS_GRAYED + 256) Then DisabledX = True
If (drawIstruct.itemState And (ODS_DISABLED + 256)) = (ODS_DISABLED + 256) Then DisabledX = True

If (drawIstruct.itemState And (ODS_GRAYED)) = (ODS_GRAYED) Then DisabledX = True
If (drawIstruct.itemState And (ODS_DISABLED)) = (ODS_DISABLED) Then DisabledX = True
End Sub
Public Function FindPlace(data() As Long, ByVal Value As Long, Optional Place As Boolean) As Long
On Error GoTo kraj
If UBound(data) = 0 Then
Select Case Place
Case False
If Value = data(0) Then
FindPlace = 0
Exit Function
Else
FindPlace = -1
Exit Function
End If

Case True
If Value > data(0) Then FindPlace = 1
Exit Function

End Select
End If


Dim index As Long
Dim xl As Long
Dim reverse As Boolean
Dim addX As Long
index = CLng(UBound(data) / 2)
xl = CLng(index / 2)

Do
If data(index) < Value Then
index = index + xl
reverse = False: addX = 1
ElseIf data(index) > Value Then
index = index - xl
reverse = True: addX = -1
ElseIf data(index) = Value Then
FindPlace = index: Exit Function
End If
If xl = 0 Then Exit Do
xl = CLng(xl / 2)
Loop

'Zadnje traženje!
Do While (data(index) < Value) Xor reverse
index = index + addX
If data(index) = Value Then FindPlace = index: Exit Function
Loop

kraj:
If Err <> 0 Then On Error GoTo 0
If Place Then
If index < 0 Then
index = 0
ElseIf index > UBound(data) Then
ElseIf Value > data(index) Then index = index + 1
End If
FindPlace = index
Else
FindPlace = -1
End If
End Function

Public Sub AddEntry(ByVal Value As Long, ByVal index As Long, data() As Long)
On Error Resume Next
ReDim Preserve data(UBound(data) + 1)
If Err <> 0 Then ReDim data(0): CopyMemory data(0), Value, 4: Exit Sub
If index > UBound(data) Then
CopyMemory data(UBound(data)), Value, 4: Exit Sub
End If
CopyMemory ByVal VarPtr(data(index + 1)), ByVal VarPtr(data(index)), (UBound(data) - index) * 4
CopyMemory data(index), Value, 4
End Sub
Public Sub DelEntry(ByVal index As Long, data() As Long)
On Error Resume Next
If UBound(data) = 0 Then Erase data: Exit Sub
If Err <> 0 Then On Error GoTo 0: Exit Sub
If index = UBound(data) Then GoTo dalje
CopyMemory ByVal VarPtr(data(index)), ByVal VarPtr(data(index + 1)), (UBound(data) - index) * 4
CopyMemory ByVal VarPtr(data(UBound(data))), 0&, 4
dalje:
ReDim Preserve data(UBound(data) - 1)
End Sub
