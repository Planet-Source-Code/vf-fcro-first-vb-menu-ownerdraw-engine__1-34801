Attribute VB_Name = "Module1"
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function LoadMenuIndirect Lib "user32" Alias "LoadMenuIndirectA" (ByVal lpMenuTemplate As Long) As Long
Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hinstance As Long, ByVal lpString As Long) As Long
Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetSystemMenu Lib "user32" _
(ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Public Const MF_BITMAP = &H4&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_BYPOSITION = &H400&
Public Const MF_POPUP = &H10&
Public Const MF_GRAYED = &H1&
Public Const MF_HELP = &H4000&
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&
Public Const MF_END = &H80
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_CALLBACKS = &H8000000


Public Const TPM_BOTTOMALIGN = &H20&
Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_HORNEGANIMATION = &H800&
Public Const TPM_HORPOSANIMATION = &H400&
Public Const TPM_NOANIMATION = &H4000&
Public Const TPM_NONOTIFY = &H80&
Public Const TPM_RECURSE = &H1&
Public Const TPM_RETURNCMD = &H100&
Public Const TPM_RIGHTALIGN = &H8&
Public Const TPM_RIGHTBUTTON = &H2&
Public Const TPM_VCENTERALIGN = &H10&
Public Const TPM_VERNEGANIMATION = &H2000&
Public Const TPM_VERPOSANIMATION = &H1000&
Public Const TPM_VERTICAL = &H40&

Public Const WM_DRAWITEM = &H2B
Public Const WM_COMMAND = &H111
Public Const WM_MEASUREITEM = &H2C

Public Const WM_PAINT = &HF

Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUGETOBJECT = &H124
Public Const WM_MENURBUTTONUP = &H122
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCOMMAND = &H126
Public Const WM_MENUDRAG = &H123
Public Const WM_MENUCHAR = &H120
Public Const WM_NCPAINT = &H85
Public Const WM_ERASEBKGND = &H14
Public Const WM_SIZE = &H5
Public Const WM_SIZING = &H214

Public Const GWL_EXSTYLE = (-20)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)


Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
 Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Dim OBJ1 As Menues
Dim tLen As Long
Dim mStr As String


Public Function ParentProc(ByVal hwnd As Long, ByVal umsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case umsg

Case WM_COMMAND
Dim MSG As Long
Dim ID As Long
ID = GetLO(wParam)
MSG = GetHI(wParam)
If lParam <> 0 Then GoTo eend
CopyMemory OBJ1, GetProp(hwnd, "OBJECT"), 4
OBJ1.RaiseClick ID
DoEvents
CopyMemory OBJ1, 0&, 4
Exit Function


Case WM_MEASUREITEM
If Not CBool(wParam) Then
Dim MIT As MEASUREITEMSTRUCT
CopyMemory MIT, ByVal lParam, Len(MIT)
CopyMemory OBJ1, GetProp(hwnd, "OBJECT"), 4
tLen = lstrlenW(MIT.ItemTextPointer)
mStr = Space(tLen)
CopyMemory ByVal StrPtr(mStr), ByVal (MIT.ItemTextPointer), tLen * 2
OBJ1.RaiseMI MIT, mStr, OBJ1.GetPar
ParentProc = True
CopyMemory ByVal lParam, MIT, Len(MIT)
CopyMemory OBJ1, 0&, 4
Exit Function
End If

Case WM_DRAWITEM
If Not CBool(wParam) Then
Dim DIT As DRAWITEMSTRUCT
CopyMemory DIT, ByVal lParam, Len(DIT)
CopyMemory OBJ1, GetProp(hwnd, "OBJECT"), 4
tLen = lstrlenW(DIT.ItemTextPointer)
mStr = Space(tLen)
CopyMemory ByVal StrPtr(mStr), ByVal (DIT.ItemTextPointer), tLen * 2
OBJ1.RaiseDI DIT, mStr, OBJ1.GetPar
ParentProc = True
CopyMemory ByVal lParam, DIT, Len(DIT)
CopyMemory OBJ1, 0&, 4
Exit Function
End If

Case WM_INITMENU
CopyMemory OBJ1, GetProp(hwnd, "OBJECT"), 4
OBJ1.InMenu = wParam
OBJ1.RaiseINIT wParam
CopyMemory OBJ1, 0&, 4

Case WM_MENUSELECT
Dim IDD As Long
Dim FLGS As Long
IDD = GetLO(wParam)
FLGS = GetHI(wParam)
CopyMemory OBJ1, GetProp(hwnd, "OBJECT"), 4
OBJ1.RaiseSL IDD, FLGS
CopyMemory OBJ1, 0&, 4


End Select
eend:
ParentProc = CallWindowProc(GetProp(hwnd, "OLDPROC"), hwnd, umsg, wParam, lParam)
End Function

Public Function GetHI(ByVal value As Long) As Long
CopyMemory GetHI, ByVal (VarPtr(value) + 2), 2
End Function
Public Function GetLO(ByVal value As Long) As Long
CopyMemory GetLO, ByVal (VarPtr(value)), 2
End Function
Public Function PutLOHI(ByVal LO As Integer, ByVal HI As Integer) As Long
CopyMemory ByVal VarPtr(PutLOHI), LO, 2
CopyMemory ByVal (VarPtr(PutLOHI) + 2), HI, 2
End Function
Public Function LongToInt(ByVal value As Long) As Integer
CopyMemory LongToInt, ByVal VarPtr(value), 2
End Function
Public Function IntToLong(ByVal value As Integer) As Long
CopyMemory ByVal VarPtr(IntToLong), value, 2
End Function
