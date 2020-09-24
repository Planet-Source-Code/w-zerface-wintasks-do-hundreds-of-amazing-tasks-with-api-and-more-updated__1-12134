VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form runf 
   BorderStyle     =   0  'None
   Caption         =   "Run..."
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   5085
   Icon            =   "runf.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "runf.frx":000C
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Open:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Type the name of a program, folder, document, or Internet Resource and Windows will open it for you."
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu exer 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "runf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const conSwNormal = 1

Private Sub Command1_Click()
On Error GoTo h:
ShellExecute hWnd, "open", Text1.Text, vbNullString, vbNullString, conSwNormal
GoTo L:
h:
Dim asd
asd = MsgBox("Cannot open!", vbOKOnly + vbCritical, "Error")
GoTo k:
L:
Unload Me
k:
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
cd1.flags = 1
cd1.ShowOpen
Text1.Text = cd1.FileName
End Sub

Private Sub exer_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command1_Click
End Sub
