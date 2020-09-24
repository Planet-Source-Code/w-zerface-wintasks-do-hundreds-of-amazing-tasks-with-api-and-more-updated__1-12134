VERSION 5.00
Begin VB.Form fresolution 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Screen Resolution and Color"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   Icon            =   "fresolution.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Display Properties..."
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "fresolution.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "This program will change the monitor resolution to the specified height/width/color in pixils/bits"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "fresolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const DM_DISPLAYFLAGS = &H200000
Const DM_DISPLAYFREQUENCY = &H400000

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpInitData As DEVMODE, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (lpszDeviceName As Any, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Const BITSPIXEL = 12

' /* Flags for ChangeDisplaySettings */
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST = &H2
Const CDS_FULLSCREEN = &H4
Const CDS_GLOBAL = &H8
Const CDS_SET_PRIMARY = &H10
Const CDS_RESET = &H40000000
Const CDS_SETRECT = &H20000000
Const CDS_NORESET = &H10000000

' /* Return values for ChangeDisplaySettings */
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
Const DISP_CHANGE_FAILED = -1
Const DISP_CHANGE_BADMODE = -2
Const DISP_CHANGE_NOTUPDATED = -3
Const DISP_CHANGE_BADFLAGS = -4
Const DISP_CHANGE_BADPARAM = -5

Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

Dim d() As DEVMODE, lNumModes As Long

Private Sub Command1_Click()
    Dim L As Long, flags As Long, x As Long
    x = List1.ListIndex
    d(x).dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
    flags = CDS_UPDATEREGISTRY
    L = ChangeDisplaySettings(d(x), flags)
    Select Case L
        Case DISP_CHANGE_RESTART
            L = MsgBox("This change will not take effect until you reboot the system.  Reboot now?", vbYesNo)
            If L = vbYes Then
                flags = 0
                L = ExitWindowsEx(EWX_REBOOT, flags)
            End If
        Case DISP_CHANGE_SUCCESSFUL
        Case Else
            MsgBox "Error changing resolution! Returned: " & L
    End Select
    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5)
End Sub

Private Sub Form_Load()
    Dim L As Long, lMaxModes As Long
    Dim lBits As Long, lWidth As Long, lHeight As Long
    lBits = GetDeviceCaps(hdc, BITSPIXEL)
    lWidth = Screen.Width \ Screen.TwipsPerPixelX
    lHeight = Screen.Height \ Screen.TwipsPerPixelY
    lMaxModes = 8
    ReDim d(0 To lMaxModes) As DEVMODE
    lNumModes = 0
    L = EnumDisplaySettings(ByVal 0, lNumModes, d(lNumModes))
    Do While L
        List1.AddItem d(lNumModes).dmPelsWidth & "x" & d(lNumModes).dmPelsHeight & "x" & d(lNumModes).dmBitsPerPel
        If lBits = d(lNumModes).dmBitsPerPel And _
           lWidth = d(lNumModes).dmPelsWidth And _
           lHeight = d(lNumModes).dmPelsHeight Then
            List1.ListIndex = List1.NewIndex
        End If
        lNumModes = lNumModes + 1
        If lNumModes > lMaxModes Then
            lMaxModes = lMaxModes + 8
            ReDim Preserve d(0 To lMaxModes) As DEVMODE
        End If
        L = EnumDisplaySettings(ByVal 0, lNumModes, d(lNumModes))
    Loop
    lNumModes = lNumModes - 1
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command1_Click
End Sub
