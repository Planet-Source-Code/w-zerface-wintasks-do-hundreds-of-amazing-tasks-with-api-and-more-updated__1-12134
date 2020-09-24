VERSION 5.00
Begin VB.Form Covert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ascii Charactor to decimal Converter"
   ClientHeight    =   660
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4365
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "converter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton c1 
      Caption         =   "Press here to display the charactor equalivent  "
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   80
      TabIndex        =   0
      Tag             =   "ntoc"
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Type any key to find out what its Ascii decimal equalivent is."
      Height          =   255
      Left            =   80
      TabIndex        =   2
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Covert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub c1_Click()
ss = Text1.Text
If ss > 255 Then GoTo ddd:
Text1.SetFocus
KeyAscii = ss
ddd:
End Sub

Private Sub change_Click()
If Text1.Tag = "ntoc" Then
Change.Caption = "Charactor to #"
Text1.Tag = "cton"
Label1.Caption = "Type a # to find its Ascii charactor."
c1.Visible = True
Form1.Height = 1700
Form1.Caption = "Ascii # to Charactor Converter"



GoTo bbb:
End If
If Text1.Tag = "cton" Then
Change.Caption = "# to Charactor"
Text1.Tag = "ntoc"
Label1.Caption = "Type any key to find out what its Ascii decimal equalivent is."
Form1.Height = 1335
c1.Visible = False
Form1.Caption = "Ascii Charactor to # Converter"



End If
bbb:
End Sub

Private Sub exer_Click()
Unload Me
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Text1.Text = "This is not an Ascii charactor but its KeyCode is " & KeyCode
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If Text1.Tag = "ntoc" Then
Text1.Text = ""
Text1.Text = " = Ascii decimal #: " & KeyAscii
If KeyAscii = 13 Then
KeyAscii = 0
End If
If KeyAscii = 27 Then
KeyAscii = 0
End If
End If
End Sub

