VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   Caption         =   "Menu Engine/Compiler V1.01B by Vanja Fuckar,EMAIL:INGA@VIP.HR    With Two Owner Draw Engines!"
   ClientHeight    =   8580
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8580
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Complex Menu Sample 2"
      Height          =   375
      Left            =   8520
      TabIndex        =   16
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   375
      Left            =   10800
      TabIndex        =   15
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   7440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   2280
      Width           =   4335
   End
   Begin MSComctlLib.ImageList ImageList16x16 
      Left            =   5880
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":02B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":040E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0568
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":06C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":081C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0976
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Owner Draw Engine 2 Small Icons"
      Height          =   195
      Index           =   2
      Left            =   7920
      TabIndex        =   11
      Top             =   1200
      Width           =   3135
   End
   Begin MSComctlLib.ImageList ImageList32x32 
      Left            =   5880
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":3282
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":5A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":81E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":A998
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":D14A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":D464
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":FC16
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":123C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":14B7A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Owner Draw Engine 2 Large Icons"
      Height          =   195
      Index           =   1
      Left            =   7920
      TabIndex        =   10
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sample 1"
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   7800
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Owner Draw Engine 1"
      Height          =   195
      Index           =   0
      Left            =   7920
      TabIndex        =   8
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H008080FF&
      Caption         =   "Save Binary File"
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H008080FF&
      Caption         =   "Load Binary File"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000C000&
      Caption         =   "Track Pop Up"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Load Text File"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Save Text File"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   240
      Width           =   7335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Compile And Menu Bar"
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   7800
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Changes Beetween Owner Draw Engines Requires Recompilation!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   7440
      TabIndex        =   14
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Menu Compiler Help"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7440
      TabIndex        =   13
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Menu Compiler Code:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7335
   End
   Begin VB.Menu TemporaryRequired 
      Caption         =   ""
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MM As New MenuMaker
Private WithEvents MENU1 As Menues
Attribute MENU1.VB_VarHelpID = -1
Private OWNER1 As New DrawEngine1
Private OWNER2 As New DrawEngine2
Private Mhandle As Long
Dim sFile As String
Dim sPath As String



Private Sub Command1_Click()
MM.MenuData = Text1
Mhandle = MM.CompileMenu
MENU1.ImportMenu Mhandle
MENU1.ShowMenu Me.hWnd
End Sub

Private Sub Command10_Click()
If MM.Handle = 0 Then MsgBox "Compile First!", vbInformation, "Info": Exit Sub
Dim aa As Boolean
aa = GetSaveFilePath(hWnd, "", 0, "", "", "", "Save As Text File", sPath)
If aa = False Then Exit Sub
MM.SaveBinaryMenu sPath
End Sub





Private Sub Command2_Click()
If Text1 = "" Then Exit Sub
Dim aa As Boolean
aa = GetSaveFilePath(hWnd, "", 0, "", "", "", "Save As Binary File", sPath)
If aa = False Then Exit Sub
Dim FFL As Long
FFL = FreeFile
If Dir(sPath) <> "" Then Kill sPath
Open sPath For Binary As #FFL
Put #1, , Text1.Text
Close #1
End Sub



Private Sub Command3_Click()
Text1 = LoadResString(1)
End Sub

Private Sub Command4_Click()
Dim aa As Boolean
aa = GetOpenFilePath(hWnd, "", 0, sFile, "", "Load Text File", sPath)
If aa = False Then Exit Sub
Dim FFL As Long
FFL = FreeFile
Dim STRX As String
Open sPath For Binary Access Read As #FFL
STRX = Space(LOF(FFL))
Get #FFL, , STRX
Close #1
Text1 = STRX
STRX = ""
End Sub


Private Sub Command5_Click()
MsgBox "This is Non Ownerdraw Menu Sample!", vbInformation, "Information"
Text1 = LoadResString(3)
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
MENU1.TrackMenu 0, 250, 250, Me.hWnd
End Sub



Private Sub Command9_Click()
Dim aa As Boolean
aa = GetOpenFilePath(hWnd, "", 0, sFile, "", "Load Binary File", sPath)
If aa = False Then Exit Sub
MENU1.LoadBinaryMenu sPath
MENU1.ShowMenu hWnd
End Sub

Private Sub Form_Load()
Text2 = LoadResString(2)
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Set MENU1 = New Menues
End Sub


Private Sub Form_Unload(Cancel As Integer)
MsgBox "Next OwnerDraw Engine will be published very soon!", vbExclamation, "Information"
End Sub

Private Sub MENU1_DrawItem(drawIstruct As MenuObject.DRAWITEMSTRUCT, ByVal MenuText As String, ByVal ParentHwnd As Long)
If Option1(0).Value = True Then
'OVIM OMOGUCUJETE DA DRAWENGINE OBRADI OWNERDRAW PORUKU!!!!
'THIS COMMAND WILL ENABLE DRAWENGINE TO PROCESS OWNERDRAW MESSAGE!
OWNER1.ProcessDrawItem drawIstruct, MenuText, ParentHwnd
ElseIf Option1(1).Value = True Or Option1(2).Value = True Then
OWNER2.ProcessDrawItem drawIstruct, MenuText, ParentHwnd
End If

End Sub

Private Sub MENU1_MeasureItem(measureIstruct As MenuObject.MEASUREITEMSTRUCT, ByVal MenuText As String, ByVal ParentHwnd As Long)
If Option1(0).Value = True Then
'OVIM OMOGUCUJETE DA DRAWENGINE OBRADI OWNERDRAW PORUKU!!!!
'THIS COMMAND WILL ENABLE DRAWENGINE TO PROCESS OWNERDRAW MESSAGE!
OWNER1.ProcessMeasureItem measureIstruct, MenuText, ParentHwnd
ElseIf Option1(1).Value = True Or Option1(2).Value = True Then
OWNER2.ProcessMeasureItem measureIstruct, MenuText, ParentHwnd
End If

End Sub

Private Sub MENU1_MenuClick(ByVal ID As Long)
MsgBox "Menu Click ID:" & ID, vbInformation, "MENU OBJECT 1"
End Sub

Private Sub MENU1_MenuInitialization(ByVal Handle As Long)
If Option1(0).Value = True Then
OWNER1.Background = &HA25366
OWNER1.ForegroundSelected = &H66FFFF
OWNER1.ForegroundNotSelected = &HCCEE&
OWNER1.HighlightColor = &H2222AA
OWNER1.MenuFont "trebuchet ms", 9, FW_BOLD, 238, True, False, False
OWNER1.SeparatorColor = &HEE&
OWNER1.DisabledColor = &H5522&
OWNER1.SeparatorHeight = 1
OWNER1.Sound = 3
OWNER1.SelectedEdge = RaisedOuter
OWNER1.TextAlignment = CenterAlignment

ElseIf Option1(2).Value = True Then
OWNER2.MenuFont(Size16x16) = "trebuchet ms"
OWNER2.EraseIcons
OWNER2.IconSize = Size16x16
OWNER2.Icon(1) = ImageList16x16.ListImages(1).Picture
OWNER2.Icon(3) = ImageList16x16.ListImages(2).Picture
OWNER2.Icon(4) = ImageList16x16.ListImages(3).Picture
OWNER2.Icon(5) = ImageList16x16.ListImages(4).Picture
OWNER2.Icon(2) = ImageList16x16.ListImages(5).Picture
OWNER2.Icon(6) = ImageList16x16.ListImages(6).Picture
OWNER2.Icon(7) = ImageList16x16.ListImages(7).Picture
OWNER2.Icon(8) = ImageList16x16.ListImages(8).Picture
OWNER2.SelectedColor = BlueGrad

ElseIf Option1(1).Value = True Then
OWNER2.MenuFont(Size32x32) = "verdana"
OWNER2.EraseIcons
OWNER2.IconSize = Size32x32
OWNER2.Icon(1) = ImageList32x32.ListImages(5).Picture
OWNER2.Icon(3) = ImageList32x32.ListImages(2).Picture
OWNER2.Icon(4) = ImageList32x32.ListImages(4).Picture
OWNER2.Icon(5) = ImageList32x32.ListImages(8).Picture
OWNER2.Icon(2) = ImageList32x32.ListImages(6).Picture
OWNER2.Icon(6) = ImageList32x32.ListImages(9).Picture
OWNER2.Icon(7) = ImageList32x32.ListImages(7).Picture
OWNER2.Icon(8) = ImageList32x32.ListImages(10).Picture
OWNER2.SelectedColor = GreenGrad
End If


End Sub




