VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Menues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemstate As Long
    hMenu As Long
    hdc As Long
    rcItem As RECT
    ItemTextPointer As Long
End Type

Public Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    ItemHeight As Long
    ItemTextPointer As Long
End Type


Private INITTMENU As Long
Private MN() As Byte
Private Menu1 As Long
Private PAR As Long
Public Event MenuInitialization(ByVal Handle As Long)
Public Event MenuClick(ByVal ID As Long)
Public Event MenuSelect(ByVal ID As Long, ByVal MenuFlags As Long)
Public Event MeasureItem(measureIstruct As MEASUREITEMSTRUCT, ByVal MenuText As String, ByVal ParentHwnd As Long)
Public Event DrawItem(drawIstruct As DRAWITEMSTRUCT, ByVal MenuText As String, ByVal ParentHwnd As Long)
Friend Property Let InMenu(ByVal value As Long)
INITTMENU = value
End Property
Friend Property Get InMenu() As Long
InMenu = INITTMENU
End Property
Friend Function GetPar() As Long
GetPar = PAR
End Function
Friend Sub RaiseSL(ByVal ID As Long, ByVal MenuFlags As Long)
RaiseEvent MenuSelect(ID, MenuFlags)
End Sub
Friend Sub RaiseINIT(ByVal Handle As Long)
RaiseEvent MenuInitialization(Handle)
End Sub
Friend Sub RaiseClick(ByVal ID As Long)
RaiseEvent MenuClick(ID)
End Sub
Friend Sub RaiseMI(measureIstruct As MEASUREITEMSTRUCT, ByVal MenuText As String, ByVal ParentHwnd As Long)
RaiseEvent MeasureItem(measureIstruct, MenuText, ParentHwnd)
End Sub
Friend Sub RaiseDI(drawIstruct As DRAWITEMSTRUCT, ByVal MenuText As String, ByVal ParentHwnd As Long)
RaiseEvent DrawItem(drawIstruct, MenuText, ParentHwnd)
End Sub
Public Function LoadBinaryMenu(ByVal filename As String) As Long
Dim FN As Long
FN = FreeFile
Open filename For Binary Access Read As #FN
ReDim MN(LOF(FN) - 1)
Get #FN, , MN
Close #1
Menu1 = LoadMenuIndirect(ByVal VarPtr(MN(0)))
LoadBinaryMenu = Menu1
End Function
Public Property Get MenuCount() As Long
MenuCount = GetMenuItemCount(Menu1)
End Property
Public Property Get Handle() As Long
Handle = Menu1
End Property
Public Sub ImportMenu(ByVal Handle As Long)
Menu1 = Handle
End Sub
Public Property Let Enabled(ByVal ID As Long, ByVal itemstate As Boolean)
If Not itemstate Then
EnableMenuItem Menu1, ID, 0
Else
EnableMenuItem Menu1, ID, MF_DISABLED
End If
End Property
Public Property Get Enabled(ByVal ID As Long) As Boolean
If Not ((GetMenuState(Menu1, ID, 0) And MF_DISABLED) = MF_DISABLED) And _
 Not ((GetMenuState(Menu1, ID, 0) And MF_GRAYED) = MF_GRAYED) Then Enabled = True
End Property
Public Property Let Grayed(ByVal ID As Long, ByVal itemstate As Boolean)
If Not itemstate Then
EnableMenuItem Menu1, ID, 0
Else
EnableMenuItem Menu1, ID, MF_GRAYED
End If
End Property
Public Property Get Grayed(ByVal ID As Long) As Boolean
If (GetMenuState(Menu1, ID, 0) And MF_GRAYED) = MF_GRAYED Then Grayed = True
End Property
Public Property Let Checked(ByVal ID As Long, ByVal itemstate As Boolean)
If Not itemstate Then
CheckMenuItem Menu1, ID, 0
Else
CheckMenuItem Menu1, ID, MF_CHECKED
End If
End Property
Public Property Get Checked(ByVal ID As Long) As Boolean
If (GetMenuState(Menu1, ID, 0) And MF_CHECKED) = MF_CHECKED Then Checked = True
End Property

Public Sub ShowMenu(ByVal ParentHwnd As Long)
If Not CBool(Menu1) Then Exit Sub
PAR = ParentHwnd
NullProc
Dim OM As Long
OM = GetMenu(ParentHwnd)
SetMenu ParentHwnd, Menu1
DrawMenuBar ParentHwnd
If (OM <> Menu1) And (OM <> 0) Then
DestroyMenu OM
End If
If Not CBool(GetProp(ParentHwnd, "OLDPROC")) Then
SetProp ParentHwnd, "OBJECT", ObjPtr(Me)
SetProp ParentHwnd, "OLDPROC", SetWindowLong(ParentHwnd, GWL_WNDPROC, AddressOf ParentProc)
End If
End Sub
Public Sub TrackMenu(ByVal ID As Long, ByVal x As Long, ByVal Y As Long, ByVal ParentHwnd As Long)
If Not CBool(Menu1) Then Exit Sub
Dim RC As RECT
PAR = ParentHwnd
NullProc
If Not CBool(GetProp(ParentHwnd, "OLDPROC")) Then
SetProp ParentHwnd, "OBJECT", ObjPtr(Me)
SetProp ParentHwnd, "OLDPROC", SetWindowLong(ParentHwnd, GWL_WNDPROC, AddressOf ParentProc)
End If
Call TrackPopupMenu(GetSubMenu(Menu1, ID), 0, x, Y, 0, ParentHwnd, RC)
End Sub
Public Sub KillMenu()
DestroyMenu Menu1
SetMenu PAR, ByVal 0&
DrawMenuBar PAR
Menu1 = 0
End Sub
Private Sub Class_Terminate()
NullProc
End Sub
Private Sub NullProc()
Dim oPr As Long
oPr = GetProp(PAR, "OLDPROC")
If oPr <> 0 Then
Call SetWindowLong(PAR, GWL_WNDPROC, oPr)
End If
RemoveProp PAR, "OLDPROC"
RemoveProp PAR, "OBJECT"
End Sub



