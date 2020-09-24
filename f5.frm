VERSION 5.00
Begin VB.Form f5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pick SMTP Host"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   Icon            =   "f5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Find SMTP..."
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Paradise"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Yahoo"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MSN"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   735
   End
   Begin VB.OptionButton Option14 
      Caption         =   "AOL"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.OptionButton Option9 
      Caption         =   "Juno"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton Option13 
      Caption         =   "Microsoft"
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.OptionButton Option12 
      Caption         =   "NetZero"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton Option10 
      Caption         =   "Geocities"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton Option8 
      Caption         =   "Email"
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Xtra"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Bellsouth"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton Option4 
      Caption         =   "McCracken"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Hotmail"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.OptionButton Option11 
      Caption         =   "Blue Light"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Pick an SMTP Host:"
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   45
      Width           =   2055
   End
End
Attribute VB_Name = "f5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
ShellExecute hWnd, "open", "http://www.maxban.com/tools/smtp.html?Domain=yahoo.com", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Option1_Click()
f2.smtp_host.Text = "smtp.email.msn.com"
End Sub

Private Sub Option14_Click()
f2.smtp_host.Text = "aolmail.aol.com"
End Sub

Private Sub Option10_Click()
f2.smtp_host.Text = "mail.geocities.com"
End Sub

Private Sub Option11_Click()
f2.smtp_host.Text = "mail.bluelight.com"
End Sub

Private Sub Option12_Click()
'inbound-mail.netzero.net
f2.smtp_host.Text = "inbound-mail.netzero.net"
End Sub

Private Sub Option13_Click()
'mail5.microsoft.com
f2.smtp_host.Text = "mail5.microsoft.com"
End Sub

Private Sub Option2_Click()
f2.smtp_host.Text = "smtp.mail.yahoo.com"
End Sub

Private Sub Option3_Click()
f2.smtp_host.Text = "mail.hotmail.com"
End Sub

Private Sub Option4_Click()
'209.174.47.10
f2.smtp_host.Text = "mccracken.skokie735.k12.il.us"

End Sub

Private Sub Option5_Click()
f2.smtp_host.Text = "mail.atl.bellsouth.net"
End Sub

Private Sub Option6_Click()
f2.smtp_host.Text = "smtp.paradise.net.nz"
End Sub

Private Sub Option7_Click()
f2.smtp_host.Text = "smtp.xtra.co.nz"
End Sub

Private Sub Option8_Click()
f2.smtp_host.Text = "mail-intake-1.mail.com"
End Sub

Private Sub Option9_Click()
f2.smtp_host.Text = "mx.boston.juno.com"
End Sub
