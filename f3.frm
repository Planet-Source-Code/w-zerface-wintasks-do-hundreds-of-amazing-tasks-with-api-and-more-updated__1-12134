VERSION 5.00
Begin VB.Form f3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save Settings"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   Icon            =   "f3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   120
      Picture         =   "f3.frx":000C
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "3. Press 'Save' to save the settings."
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Line Line2 
      X1              =   1800
      X2              =   3840
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label3 
      Caption         =   "2. Type a filename to save the settings in:"
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   1800
      Y1              =   720
      Y2              =   3240
   End
   Begin VB.Label Label2 
      Caption         =   "1. Pick a directory to save the settings file in:"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "To save all information in the Fakemail form as a file, follow the directions below."
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "f3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo lll:
Dim file
file = Dir1.Path & "\" & Text1.Text
a = f2.smtp_host.Text
b = f2.port.Text
c = f2.f_email.Text
d = f2.to_mail.Text
e = f2.sender_name.Text
f = f2.r_name.Text
g = f2.subject.Text
h = f2.body.Text
Open file For Output As #1
Write #1, a, b, c, d, e, f, g, h
Close #1
Dim s
s = MsgBox("Settings saved in " & file, vbOKOnly + vbInformation, "Saved!")
Unload Me
GoTo kkk:
lll:
Dim fg
fg = MsgBox("Cannot save!", vbOKOnly + vbCritical, "Error!")
kkk:
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
On Error GoTo h:
Dir1.Path = Drive1.Drive
GoTo k:
h:
Dim d
d = MsgBox("Cannot open drive!", vbOKOnly + vbCritical, "Error!")
Drive1.Drive = "c:"
k:
End Sub


Private Sub Form_Load()
Dir1.Path = "c:"

End Sub

Private Sub Text1_LostFocus()
If Not Text1.Text Like "*.fms" Then
Text1.Text = Text1.Text & ".fms"
End If
End Sub
