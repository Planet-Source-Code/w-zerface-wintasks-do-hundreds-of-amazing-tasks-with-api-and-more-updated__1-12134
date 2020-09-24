VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form mainf 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "WinTasks"
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   Icon            =   "mainf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer4 
      Interval        =   10
      Left            =   2400
      Top             =   2280
   End
   Begin VB.Timer auto 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2040
      Top             =   2400
   End
   Begin VB.Timer Timer15 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   2760
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1920
      Top             =   2760
   End
   Begin VB.PictureBox pctPrg 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00008000&
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   0
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   6
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4920
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3840
      Top             =   720
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   1080
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   3480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Window 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mouse is over:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CPU Usage:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes in Windows:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label dates 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   -120
      X2              =   3360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Times 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu about 
         Caption         =   "&About..."
         Shortcut        =   ^A
      End
      Begin VB.Menu help 
         Caption         =   "&Help"
      End
      Begin VB.Menu exer 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu comm 
      Caption         =   "&Controls"
      Begin VB.Menu comp 
         Caption         =   "&Windows"
         Begin VB.Menu sd 
            Caption         =   "&Shut Down"
            Shortcut        =   {F2}
         End
         Begin VB.Menu re 
            Caption         =   "&Restart"
            Shortcut        =   {F3}
         End
         Begin VB.Menu lo 
            Caption         =   "&Log Off"
            Shortcut        =   {F4}
         End
         Begin VB.Menu sep99 
            Caption         =   "-"
         End
         Begin VB.Menu cprog 
            Caption         =   "&Control Program     F9"
         End
         Begin VB.Menu Rw 
            Caption         =   "&Refresh Windows"
            Shortcut        =   {F11}
         End
         Begin VB.Menu Secpassset 
            Caption         =   "&Secure Computer Password..."
         End
         Begin VB.Menu scputer 
            Caption         =   "&Secure Computer"
            Shortcut        =   {F7}
         End
         Begin VB.Menu amm 
            Caption         =   "&Auto Mouse Move"
            Shortcut        =   ^U
         End
         Begin VB.Menu erb 
            Caption         =   "&Empty Recycling Bin"
         End
         Begin VB.Menu set 
            Caption         =   "&Set"
            Begin VB.Menu sr 
               Caption         =   "Screen &Resoluton"
               Shortcut        =   +^{F1}
            End
            Begin VB.Menu ss123 
               Caption         =   "Screen &Saver"
               Begin VB.Menu son 
                  Caption         =   "&On"
                  Checked         =   -1  'True
               End
               Begin VB.Menu soff 
                  Caption         =   "O&ff"
               End
            End
            Begin VB.Menu taskvis 
               Caption         =   "&Taskbar"
               Begin VB.Menu vistes 
                  Caption         =   "&Show"
                  Checked         =   -1  'True
               End
               Begin VB.Menu hidyis 
                  Caption         =   "&Hide"
               End
            End
            Begin VB.Menu di123 
               Caption         =   "&Desktop Icons"
               Begin VB.Menu dshow 
                  Caption         =   "&Show"
                  Checked         =   -1  'True
               End
               Begin VB.Menu dhide 
                  Caption         =   "&Hide"
               End
            End
            Begin VB.Menu acd 
               Caption         =   "&Alt - Ctrl - Del"
               Begin VB.Menu ena 
                  Caption         =   "&Enabled"
                  Checked         =   -1  'True
               End
               Begin VB.Menu dis 
                  Caption         =   "&Disabled"
               End
            End
            Begin VB.Menu ocr 
               Caption         =   "&Open CD-Rom"
            End
         End
         Begin VB.Menu sep2000 
            Caption         =   "-"
         End
         Begin VB.Menu ma 
            Caption         =   "&Minimize All"
            Shortcut        =   ^M
         End
         Begin VB.Menu wp123 
            Caption         =   "&Programs"
            Begin VB.Menu fff 
               Caption         =   "F&ind Files or Folders..."
               Shortcut        =   ^F
            End
            Begin VB.Menu winfi 
               Caption         =   "&WinFiles"
               Shortcut        =   ^W
            End
            Begin VB.Menu re1 
               Caption         =   "E&xplore..."
               Shortcut        =   ^E
            End
            Begin VB.Menu sep48 
               Caption         =   "-"
            End
            Begin VB.Menu fmail2 
               Caption         =   "&Fakemail / Mail Bomb"
            End
            Begin VB.Menu enum 
               Caption         =   "&Enumerator"
            End
            Begin VB.Menu sfinder 
               Caption         =   """*"" &Finder"
            End
            Begin VB.Menu atd 
               Caption         =   "&ASCII to Decimal"
            End
            Begin VB.Menu vc 
               Caption         =   "&Veda Creator"
            End
         End
      End
      Begin VB.Menu internet 
         Caption         =   "&Internet"
         Begin VB.Menu email 
            Caption         =   "&E-mail"
            Begin VB.Menu ce 
               Caption         =   "&Check Email..."
               Shortcut        =   ^C
            End
            Begin VB.Menu se123 
               Caption         =   "&Send Email..."
               Shortcut        =   ^S
            End
            Begin VB.Menu sep13 
               Caption         =   "-"
            End
            Begin VB.Menu fmail 
               Caption         =   "&Fakemail \ Mail Bomb"
            End
         End
         Begin VB.Menu browser 
            Caption         =   "&Internet Browser..."
            Shortcut        =   ^B
         End
         Begin VB.Menu con 
            Caption         =   "&Connect"
         End
         Begin VB.Menu dcon 
            Caption         =   "&Diconnect"
         End
         Begin VB.Menu ip2 
            Caption         =   "&Internet Properties..."
            Shortcut        =   ^P
         End
      End
      Begin VB.Menu cp 
         Caption         =   "&Control Panel"
         Begin VB.Menu arp 
            Caption         =   "&Add/Remove Programs..."
         End
         Begin VB.Menu std1 
            Caption         =   "&Set Time/Date..."
         End
         Begin VB.Menu rs123 
            Caption         =   "&Regional Settings..."
         End
         Begin VB.Menu anh 
            Caption         =   "A&dd new hardware..."
         End
         Begin VB.Menu disp4 
            Caption         =   "&Display Properties..."
         End
         Begin VB.Menu ip1 
            Caption         =   "&Internet Properties..."
         End
         Begin VB.Menu kp 
            Caption         =   "&Keyboard Properties..."
         End
         Begin VB.Menu mp 
            Caption         =   "&Mouse Properties..."
         End
         Begin VB.Menu mp2 
            Caption         =   "M&odem Properties..."
         End
         Begin VB.Menu sysp 
            Caption         =   "S&ystem Properties..."
         End
         Begin VB.Menu np 
            Caption         =   "&Network Properties..."
         End
         Begin VB.Menu pp 
            Caption         =   "&Password Properties..."
         End
         Begin VB.Menu sp123 
            Caption         =   "S&ounds Properties..."
         End
      End
      Begin VB.Menu files 
         Caption         =   "&Files"
         Begin VB.Menu copyfile 
            Caption         =   "&Copy"
         End
         Begin VB.Menu run 
            Caption         =   "&Run..."
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu tp 
         Caption         =   "&This Program"
         Begin VB.Menu fot 
            Caption         =   "&Force On top     F8"
         End
         Begin VB.Menu aot 
            Caption         =   "&Always on top"
            Checked         =   -1  'True
         End
         Begin VB.Menu bcolor 
            Caption         =   "&BackColor..."
         End
         Begin VB.Menu size 
            Caption         =   "&Size"
            Begin VB.Menu compact 
               Caption         =   "&Compact"
               Checked         =   -1  'True
            End
            Begin VB.Menu full 
               Caption         =   "&Medium"
            End
            Begin VB.Menu full2 
               Caption         =   "&Full"
            End
         End
      End
   End
End
Attribute VB_Name = "mainf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim securepass
Dim a123
Private Declare Function RegOpenKeyExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegisterServiceProcess Lib "Kernel32.dll" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Private Declare Function Getasynckeystate Lib "user32" Alias "GetAsyncKeyState" (ByVal VKEY As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPIF_SENDWININICHANGE = &H2
Const Internet_Autodial_Force_Unattended As Long = 2
Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal _
uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As _
Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Declare Function SetCursorPos Lib "user32.dll" (ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    x As Long
    Y As Long
End Type
Private XT(1) As Single, YT As Single, M As Single
Private XScreen As Single, YScreen As Single
Private i As Integer, II As Integer
Private Const MAX_DELOCATION = 250
Private Const PUPIL_DISTANCE = 30
Option Explicit
Private CPU As New CPUUsage
Private Avg As Long                         ' Average of CPU Usage
Private Sum As Long
Private Index As Long
Dim timeval
Private Declare Function GetTickCount Lib "Kernel32.dll" () As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40
'Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2
Const VK_LWIN = &H5B
Private Const EWX_SHUTDOWN As Long = 1
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Const EWX_REBOOT As Long = 2
Private Const EWX_LOGOFF As Long = 0
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const conSwNormal = 1
Private Const SW_SHOW = 5
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Dim dgf
Const SPI_SCREENSAVERRUNNING = 97
Private Declare Function BringWindowToTop Lib "user32.dll" (ByVal hWnd As Long) As Long
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

'Declare the constants for the API function. These constants can be
'found in the header file Shellapi.h.

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Private Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.

'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA







Private Sub about_Click()
Form7.Show
End Sub

Private Sub amm_Click()
If amm.Checked = False Then
auto.Enabled = True
amm.Checked = True
Else
auto.Enabled = False
amm.Checked = False
End If
End Sub

'''
Private Sub anh_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", 5)
End Sub

Private Sub aot_Click()
If aot.Checked = True Then
SetWindowPos hWnd, conHwndNoTopmost, 100, 100, 205, 141, conSwpNoActivate Or conSwpShowWindow
aot.Checked = False
Me.Left = 0
Me.Top = 0
mainf.Height = 615
mainf.Width = 2070
GoTo lll:
End If
If aot.Checked = False Then
aot.Checked = True
SetWindowPos hWnd, conHwndTopmost, 0, 0, 205, 141, conSwpNoActivate Or conSwpShowWindow
Me.Left = 0
Me.Top = 0
mainf.Height = 615
mainf.Width = 2070
End If
lll:
End Sub

Private Sub arp_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", 5)
End Sub

Private Sub atd_Click()
Covert.Show
End Sub

Private Sub auto_Timer()
Dim retvals
retvals = SetCursorPos(Rnd * 1000, Rnd * 700)
End Sub

Private Sub bcolor_Click()
CommonDialog1.flags = 1
CommonDialog1.Color = mainf.BackColor
CommonDialog1.ShowColor
mainf.BackColor = CommonDialog1.Color
End Sub

Private Sub browser_Click()
ShellExecute hWnd, "open", "", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub ce_Click()
Shell ("C:\Program Files\Outlook Express\msimn.exe")
End Sub

Private Sub compact_Click()
Timer5.Enabled = False
Timer3.Enabled = False
compact.Checked = True
full2.Checked = False
full.Checked = False
mainf.Height = 615
mainf.Width = 2070
End Sub

Private Sub con_Click()
Dim lResult As Long
lResult = InternetAutodial(Internet_Autodial_Force_Unattended, 0&)
End Sub

Private Sub copyfile_Click()
Form5.Show
End Sub

Private Sub cprog_Click()
form6.Show
End Sub

Private Sub dcon_Click()
Dim lResult As Long
lResult = InternetAutodialHangup(0&)
End Sub

Private Sub dhide_Click()
Dim hWnd As Long
hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hWnd, 0
dhide.Checked = True
dshow.Checked = False
End Sub

Private Sub dis_Click()
callme (True)
ena.Checked = False
dis.Checked = True
End Sub

Private Sub disp4_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5)
End Sub

Private Sub dshow_Click()
Dim hWnd As Long
hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hWnd, 5
dhide.Checked = False
dshow.Checked = True
End Sub

Private Sub ena_Click()
callme (False)
ena.Checked = True
dis.Checked = False
End Sub

Private Sub enum_Click()
Form1.Show
End Sub

Private Sub erb_Click()
Dim retvaL
retvaL = SHEmptyRecycleBin(Form1.hWnd, "", SHERB_NOPROGRESSUI)

End Sub

Private Sub exer_Click()
Unload Me
End Sub

Private Sub fff_Click()
Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(70, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub fmail_Click()
f2.Show
End Sub

Private Sub fmail2_Click()
f2.Show
End Sub

Private Sub Form_Load()
a123 = 0
mainf.Height = 615
mainf.Width = 2070
 XScreen = Screen.Width / Screen.TwipsPerPixelX
    YScreen = Screen.Height / Screen.TwipsPerPixelY
    II = 1
SetWindowPos hWnd, conHwndTopmost, 0, 0, 205, 141, conSwpNoActivate Or conSwpShowWindow
'CPU.InitCPUUsage
Me.Left = 0
Me.Top = 0
mainf.Height = 615
mainf.Width = 2070
nid.cbSize = Len(nid)
   nid.hWnd = mainf.hWnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = mainf.Icon
   nid.szTip = "Windows Control Program" & vbNullChar

   'Call the Shell_NotifyIcon function to add the icon to the taskbar
   'status area.
   Shell_NotifyIcon NIM_ADD, nid
Dim checktiME
On Error Resume Next
Open "c:\windows\check.win" For Append As #1
Close #1
Open "c:\windows\check.win" For Input As #1
Input #1, checktiME
Close #1

If Not checktiME = "ok" Then
Open "c:\windows\check.win" For Output As #1
checktiME = "ok"
Write #1, checktiME
Close #1
Open "C:\windows\desktop\Wintasks ReadMe.txt" For Output As #1
Print #1, "Wintasks was created by Martin McCormick.  Wintasks is Freeware and can be copied without limitations.  Please send your questions or comments to slimshady_5_5_5@hotmail.com.  If you delete this text file and run Wintasks again, this file will NOT appear again."
Close #1
End If
Open "C:\windows\secpass.dat" For Append As #1
Close #1
Open "C:\windows\secpass.dat" For Input As #1
Input #1, securepass
Close #1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
    Dim Msg As Long
    Dim sFilter As String
    Msg = x / Screen.TwipsPerPixelX
    Select Case Msg
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       mainf.Visible = True
       mainf.WindowState = 0
       AppActivate ("Win")
       Case WM_LBUTTONDBLCLK
      
       mainf.Visible = True
       mainf.WindowState = 0
       AppActivate ("Win")
       Case WM_RBUTTONDOWN
          Dim ToolTipString As String
           
          If ToolTipString <> "" Then
             nid.szTip = ToolTipString & vbNullChar
             Shell_NotifyIcon NIM_MODIFY, nid
          End If
       Case WM_RBUTTONUP
       Case WM_RBUTTONDBLCLK
    End Select

End Sub



Private Sub Form_Terminate()
'Shell ("c:\windows\desktop\wintasks.exe")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If dgf = 1 Then
Cancel = 1
Else
Dim qwe
qwe = MsgBox("Are you sure you want to close WinTasks?", vbQuestion + vbYesNo, "Are you sure?")
If qwe = vbYes Then GoTo h:
If qwe = vbNo Then
Cancel = 1
GoTo h2:
End If
End If
h:
Shell_NotifyIcon NIM_DELETE, nid
h2:
End Sub

Private Sub fot_Click()
If fot.Checked = False Then
Timer15.Enabled = True
fot.Checked = True
GoTo lll:
Else
fot.Checked = False
Timer15.Enabled = False
End If


lll:
End Sub

Private Sub full_Click()
full2.Checked = False
full.Checked = True
compact.Checked = False
mainf.Height = 2000
mainf.Width = 2500
Timer3.Enabled = True
Timer5.Enabled = False
End Sub

Private Sub full2_Click()
Timer3.Enabled = True
mainf.Height = 2900
mainf.Width = 2500
full2.Checked = True
full.Checked = False
compact.Checked = False
Timer5.Enabled = True
End Sub

Private Sub help_Click()
Dim asda
asda = MsgBox("If you have a problem with this program please email slimshady_5_5_5@hotmail.com   NOTE: The 'Check Email' will not work on computers that do not have Outlook Express.", vbInformation + vbOKOnly, "Help")
End Sub

Private Sub hidyis_Click()
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
vistes.Checked = False
hidyis.Checked = True
End Sub

Private Sub ip1_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
End Sub

Private Sub ip2_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
End Sub

Private Sub kp_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", 5)
End Sub

Private Sub lo_Click()
Dim abc
abc = MsgBox("Are you sure you want to log off the computer?", vbYesNo + vbQuestion, "Log Off")
If abc = vbYes Then
Dim lngresult
lngresult = ExitWindowsEx(EWX_LOGOFF, 0&)
End If

End Sub

Private Sub ma_Click()
Call keybd_event(VK_LWIN, 0, 0, 0)
Call keybd_event(77, 0, 0, 0)
Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub mp_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0", 5)
End Sub

Private Sub mp2_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL modem.cpl", 5)
End Sub

Private Sub np_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", 5)
End Sub

Private Sub ocr_Click()
Dim lngReturn As Long
Dim strReturn As Long
lngReturn = mciSendString("set CDAudio door open", strReturn, 127, 0)
End Sub

Private Sub pp_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL password.cpl", 5)
End Sub

Private Sub re_Click()
Dim abc
abc = MsgBox("Are you sure you want to restart the computer?", vbYesNo + vbQuestion, "Restart")
If abc = vbYes Then
Dim lngresult
lngresult = ExitWindowsEx(EWX_REBOOT, 0&)
End If
End Sub

Private Sub re1_Click()
Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(69, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub rs123_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", 5)
End Sub

Private Sub run_Click()
runf.Show
End Sub

Private Sub Rw_Click()
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
callme (False)
Dim hWnd As Long
hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hWnd, 5
comm.Enabled = True
file.Enabled = True
fclear.Show
End Sub

Private Sub scputer_Click()
Dim hWnd As Long
hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hWnd, 0
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
Call keybd_event(VK_LWIN, 0, 0, 0)
Call keybd_event(77, 0, 0, 0)
Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
comm.Enabled = False
file.Enabled = False
dgf = 1
Timer1.Enabled = True
callme (True)
End Sub

Private Sub sd_Click()
Dim abc
abc = MsgBox("Are you sure you want to shut down the computer?", vbYesNo + vbQuestion, "Shutdown")
If abc = vbYes Then
Dim lngresult
lngresult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End If
End Sub

Private Sub se123_Click()
ShellExecute hWnd, "open", "mailto:", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub Secpassset_Click()
Dim checkres
checkres = InputBox("Enter OLD Secure Computer Password:", "Enter Old Password")
If checkres = securepass Then
securepass = InputBox("Enter NEW Secure Computer Password:", "Correct! Enter New Password", securepass)
Open "c:\windows\secpass.dat" For Output As #1
Write #1, securepass
Close #1
Else
Dim sda
sda = MsgBox("Incorrect Password!", vbOKOnly + vbCritical, "Incorrect!")
End If
End Sub

Private Sub sfinder_Click()
frmPassword.Show
End Sub

Private Sub soff_Click()
ToggleScreenSaverActive (False)
son.Checked = False
soff.Checked = True
End Sub

Private Sub son_Click()
ToggleScreenSaverActive (True)
son.Checked = True
soff.Checked = False
End Sub

Private Sub sp123_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", 5)
End Sub

Private Sub sr_Click()
fresolution.Show
End Sub

Private Sub std1_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)
End Sub

Private Sub sysp_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)
End Sub
Private Sub callme(huh As Boolean)
Dim gd
gd = SystemParametersInfo(97, huh, CStr(1), 0)
End Sub


Private Sub Timer1_Timer()
Dim fgt
fgt = InputBox("Please enter access password:", "Enter Password")
If fgt = securepass Then
dgf = 0
Timer1.Enabled = False
Call enabler
End If
End Sub

Private Sub Timer15_Timer()
SetWindowPos hWnd, conHwndTopmost, 0, 0, 138, 41, conSwpNoActivate Or conSwpShowWindow

End Sub

Private Sub Timer2_Timer()
If mainf.WindowState = 1 Then mainf.Visible = False
End Sub

Private Sub Timer3_Timer()
   
    
     Dim tmp As Long
    tmp = CPU.GetCPUUsage
    Sum = Sum + tmp
    Index = Index + 1
    Avg = Int(Sum / Index)
    'Draw the bar
    pctPrg.Cls
    pctPrg.Line (0, 0)-(tmp, 18), , BF
    pctPrg.Line (Avg, 0)-(Avg, 18), &HFF
    pctPrg.Line (Avg + 1, 0)-(Avg + 1, 18), &HFF
    DoEvents
dates.Caption = Format(Date, "mm/dd/yyyy")
Times.Caption = Format(Time, "hh:mm:ss")
Dim lngTickCount As Long
lngTickCount = GetTickCount
Label3.Caption = CStr(Round((lngTickCount / 1000 / 60))) & " Minutes in Windows"
End Sub

Private Sub Timer4_Timer()
Dim keystaTE
keystaTE = Getasynckeystate(vbKeyF8)
If (keystaTE And &H1) = &H1 Then
fot_Click
End If

keystaTE = Getasynckeystate(vbKeyF9)
If (keystaTE And &H1) = &H1 Then
form6.Show
End If
End Sub

Private Sub Timer5_Timer()
Dim cp As POINTAPI, hWnd As Long, s As String
    GetCursorPos cp
     hWnd = WindowFromPoint(cp.x, cp.Y)
    s = Space(128)
    GetWindowText hWnd, s, 128
    If Asc(Left(s, 1)) = 0 Then GetClassName hWnd, s, 128
    Window.Caption = s
    DoEvents
End Sub

Private Sub Times_Click()
a123 = a123 + 1
If a123 = 10 Then
Unload Me
End
End If
End Sub

Private Sub vc_Click()
veda.Show
End Sub

Private Sub vistes_Click()
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
vistes.Checked = True
hidyis.Checked = False
End Sub
Private Sub enabler()
Call Rw_Click
End Sub

Private Sub winfi_Click()
Shell ("C:\windows\winfile.exe")
End Sub
Public Function ToggleScreenSaverActive(Active As Boolean) _
   As Boolean
Dim lActiveFlag As Long
Dim retvaL As Long

lActiveFlag = IIf(Active, 1, 0)
retvaL = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, _
   lActiveFlag, 0, 0)
ToggleScreenSaverActive = retvaL > 0

End Function


