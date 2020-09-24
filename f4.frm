VERSION 5.00
Begin VB.Form f4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Settings"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   Icon            =   "f4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   1920
      Pattern         =   "*.fms"
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   120
      Picture         =   "f4.frx":000C
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Open"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "3. Press Open:"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   1800
      Y1              =   720
      Y2              =   3240
   End
   Begin VB.Label Label3 
      Caption         =   "2. Pick file to open:"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "To open settings for Fakemail that are stored in a file follow the directions below."
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "1. Pick a directory to save the settings file in:"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "f4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command1_Click()
'Dim file
'file = Dir1.Path & "\" & Text1.Text
'a = Form2.smtp_host.Text
'b = Form2.port.Text
'c = Form2.f_email.Text
'd = Form2.to_mail.Text
'e = Form2.sender_name.Text
'f = Form2.r_name.Text
'g = Form2.subject.Text
'h = Form2.body.Text
'Open file For Output As #1
'Write #1, a, b, c, d, e, g, h
'Close #1
'Dim s
's = MsgBox("Settings opened from " & file, vbOKOnly + vbInformation, "Saved!")
'Unload Me
'End Sub

Private Sub Command1_Click()
On Error GoTo lll:
Dim file
file = Dir1.Path & "\" & File1.FileName
Open file For Input As #1
Input #1, a, b, c, d, e, f, g, h
Close #1
f2.smtp_host.Text = a
f2.port.Text = b
f2.f_email.Text = c
f2.to_mail.Text = d
f2.sender_name.Text = e
f2.r_name.Text = f
f2.subject.Text = g
f2.body.Text = h
Dim s
s = MsgBox("Settings opened from " & file, vbOKOnly + vbInformation, "Opened!")
Unload Me
GoTo kkk:
lll:
Dim fg
fg = MsgBox("Cannot open!", vbOKOnly + vbCritical, "Error!")
kkk:
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
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

