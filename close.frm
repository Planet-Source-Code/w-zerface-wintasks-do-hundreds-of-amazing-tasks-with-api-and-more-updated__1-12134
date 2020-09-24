VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control  Program"
   ClientHeight    =   4545
   ClientLeft      =   1620
   ClientTop       =   2175
   ClientWidth     =   4875
   Icon            =   "close.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4875
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   240
      Top             =   4440
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Restore"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Minimize"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ma&ximize"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Hide"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Bring to Top"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Show"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin MSComctlLib.ListView View1 
      Height          =   3495
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Parent windows"
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6165
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&numThem"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Leave"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   3960
      Top             =   2640
   End
   Begin VB.Menu mnufile1 
      Caption         =   "&File"
      Begin VB.Menu locker 
         Caption         =   "&Locked"
         Checked         =   -1  'True
      End
      Begin VB.Menu refresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu exer200011 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Show 
         Caption         =   "&Show Window using ShowWindow API"
      End
      Begin VB.Menu Show_BWTT 
         Caption         =   "Show &Winsow using BringWindowToTop API"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu Max 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu Min 
         Caption         =   "Mi&nimize"
      End
      Begin VB.Menu Restore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu Hide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu Close 
         Caption         =   "&Close this Window"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "menu2"
      Visible         =   0   'False
      Begin VB.Menu BnClick 
         Caption         =   "&Click"
      End
   End
   Begin VB.Menu s2 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu AR1 
         Caption         =   "Auto Refresh"
      End
   End
End
Attribute VB_Name = "form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40
Private ClassResize As New CResize

'API to open the browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
    String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub AR1_Click()
If AR1.Checked = False Then
AR1.Checked = True
GoTo lll:
End If
If AR1.Checked = True Then
AR1.Checked = False
End If
lll:
End Sub

Private Sub BnClick_Click()
    'SendMessage Val(View2.SelectedItem), BM_CLICK, 0, 0
    
End Sub

Private Sub Close_Click()
    'close window code goes here:
    Dim lhwnd As Long
    
    On Error Resume Next
    lhwnd = Val(View1.SelectedItem)
    SendMessage lhwnd, WM_CLOSE, 0, 0

End Sub

Private Sub Command1_Click()
    'Free the memory occupied by the Object
    Set ClassResize = Nothing
    Unload Me

End Sub

Private Sub Command10_Click()
refresh_Click
End Sub

Private Sub Command2_Click()
    Command2.Caption = "&Refresh"
    View1.ListItems.Clear
   ' View2.ListItems.Clear
    View1.GridLines = True
    Dim myLong As Long
    VCount = 1
    myLong = EnumWindows(AddressOf WndEnumProc, View1)

End Sub

Private Sub Command3_Click()
Close_Click
    
End Sub

Private Sub Command4_Click()
Show_Click
End Sub

Private Sub Command5_Click()
Show_BWTT_Click
End Sub

Private Sub Command6_Click()
Hide_Click
End Sub

Private Sub Command7_Click()
Max_Click
End Sub

Private Sub Command8_Click()
Min_Click
End Sub

Private Sub Command9_Click()
Restore_Click
End Sub

Private Sub exer200011_Click()
Unload Me
End Sub

Private Sub Form_Load()
    'With ClassResize
        '.hParam = Form1.Height
        '.wParam = Form1.Width
        '.Map Command1, RS_Top_Left
        '.Map Command2, RS_Top_Left
        '.Map Command3, RS_Top_Left
        '.Map Label2, RS_TopOnly
        '.Map Label3, RS_LeftOnly
        '.Map View1, RS_HeightOnly
        '.Map View2, RS_HeightOnly
        '.Map Check1, RS_Top_Left
    'End With
    'Form1.Width = 11000
    
    'Me.Left = (Screen.Width - Me.Width) / 2
    'Me.Top = (Screen.Height - Me.Height) / 2
    
    View1.View = lvwReport
    With View1.ColumnHeaders
        .Add , , "Handle", 1000
        .Add , , "Class Name", 1500
        .Add , , "Text", 4500
    End With
    VCount = 1
    'View2.View = lvwReport
    'With View2.ColumnHeaders
    '    .Add , , "Handle", 1000
    '    .Add , , "Class Name", 1500
    '    .Add , , "Text", 4500
    '    .Add , , "IsPassword field", 1000
    '
    'End With
    ICount = 1
    Options.Visible = False
Command2_Click
End Sub

Private Sub Form_Resize()
    ClassResize.rSize Form1
    
    'OK now resize if you must!
     'View2.Left = Int(Form1.Width / 2)
     'View1.Width = View2.Left - 255
     'View2.Width = Int(Form1.Width / 2) - 255
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
End Sub

Private Sub Hide_Click()
    ShowWindow Val(View1.SelectedItem), SW_HIDE
End Sub

Private Sub Label2_Click()
    Dim Ret As Long
    Ret = ShellExecute(Me.hWnd, "Open", "http://go.to/abubakar", "", App.Path, 1)

End Sub


Private Sub locker_Click()
If locker.Checked = True Then
locker.Checked = False
Timer2.Enabled = False
Else
locker.Checked = True
Timer2.Enabled = True
End If
End Sub

Private Sub Max_Click()
    ShowWindow Val(View1.SelectedItem), SW_MAXIMIZE
    
End Sub

Private Sub Min_Click()
    ShowWindow Val(View1.SelectedItem), SW_MINIMIZE
End Sub

Private Sub refresh_Click()
Command2_Click
End Sub

Private Sub Restore_Click()
    ShowWindow Val(View1.SelectedItem), SW_RESTORE
End Sub

Private Sub Show_BWTT_Click()
    Dim lhwnd As Long
    
    On Error GoTo bugging
    lhwnd = Val(View1.SelectedItem)
    'ShowWindow lhwnd, SW_SHOW
    BringWindowToTop lhwnd
    
    Exit Sub
bugging:
    Rem Do Nothing
    
End Sub

Private Sub Show_Click()
    'show window code goes here:
    Dim lhwnd As Long
    On Error Resume Next

    lhwnd = Val(View1.SelectedItem)
    ShowWindow lhwnd, SW_SHOW
End Sub

Private Sub SpyMenu_Click()
    Dim st As RECT
    
    Spy_Form.Show
    SpyHwnd = Val(View1.SelectedItem)
    Spy_Form.Tree.Nodes.Clear
    'If its a MDI type window and its child windows are maximized
    'then 'GetMenuItemInfo' crashes the 'EnumerationX'.
    'I tried to cascade the windows of other app but that doesnt
    'happen, do you know how I can do this?
    'MsgBox CascadeWindows(SpyHwnd, MDITILE_SKIPDISABLED, st, 0, 0)
    'SendMessage SpyHwnd, WM_MDICASCADE, MDITILE_SKIPDISABLED, 0
    'SendMessage SpyHwnd, WM_MDITILE, MDITILE_HORIZONTAL, 0
    
    SMenu GetMenu(SpyHwnd), Spy_Form.Tree
        
End Sub

Private Sub Timer1_Timer()
If AR1.Checked = True Then
Call Command2_Click
End If
End Sub

Private Sub Timer2_Timer()
SetWindowPos hWnd, conHwndTopmost, 200, 100, 330, 345, conSwpNoActivate Or conSwpShowWindow
End Sub

Private Sub View1_Click()
    GotoChild
End Sub

Private Sub View1_DblClick()
Command3_Click
End Sub

Private Sub View1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then GotoChild
                'So that you are able to see child windows easily by
                'scrolling through up-down arrow keys instead of
                'clicking the parent window handle every time.
    
End Sub
Private Sub GotoChild()
    On Error GoTo HandleErrorPlz
    
    Dim Num As Long
    Dim myLong As Long
    Num = Val(View1.SelectedItem)
    'View2.ListItems.Clear
    'View2.GridLines = True
    ICount = 1
    'myLong = EnumChildWindows(Num, AddressOf WndEnumChildProc, View2)

HandleErrorPlz:
    'Exit Sub ' As simple as that :)
End Sub

Private Sub View1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And View1.ListItems.Count > 0 Then
        If GetMenu(Val(View1.SelectedItem)) > 0 Then
            'SpyMenu.Enabled = True
        Else
            'SpyMenu.Enabled = False
        End If
               
        PopupMenu Options
    End If
    
End Sub
'Private Sub View2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'    If Button = vbRightButton And View2.ListItems.Count > 0 Then
 '       PopupMenu menu2
 '   End If
'
'End Sub
