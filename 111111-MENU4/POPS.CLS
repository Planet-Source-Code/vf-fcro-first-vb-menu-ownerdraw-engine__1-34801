VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private PopUps As New Collection



Public property let ReplacePopUp(DialogControl As DialogControlProp, ByVal index As Long) As Long
If index < 1 Or index > Childs.Count Then ReplaceChild = 1: Exit Function
Childs.Remove (index)
If Childs.Count = 0 Or index > Childs.Count Then
Childs.Add DialogControl
Else
Childs.Add DialogControl, , index
End If
End Function


