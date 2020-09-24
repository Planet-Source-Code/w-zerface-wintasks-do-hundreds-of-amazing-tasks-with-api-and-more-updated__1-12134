VERSION 5.00
Begin VB.Form veda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vedas"
   ClientHeight    =   8625
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "veda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   7680
      TabIndex        =   32
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6720
      TabIndex        =   30
      Text            =   "1"
      Top             =   7800
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   480
      Top             =   960
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   10080
      TabIndex        =   29
      Text            =   "5"
      Top             =   7680
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   10080
      TabIndex        =   26
      Text            =   "1"
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   1680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Randomize"
      Height          =   375
      Left            =   8760
      TabIndex        =   25
      Top             =   8160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   8400
      TabIndex        =   24
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   8040
      TabIndex        =   23
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   7680
      TabIndex        =   22
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   7320
      TabIndex        =   21
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   6960
      TabIndex        =   20
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   6600
      TabIndex        =   19
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   6240
      TabIndex        =   18
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   5880
      TabIndex        =   17
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   5520
      TabIndex        =   16
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   5160
      TabIndex        =   15
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   4800
      TabIndex        =   14
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   4440
      TabIndex        =   13
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   4080
      TabIndex        =   12
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   3720
      TabIndex        =   11
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   3360
      TabIndex        =   10
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   3000
      TabIndex        =   9
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   2640
      TabIndex        =   8
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   2280
      TabIndex        =   7
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1920
      TabIndex        =   6
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   5
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   4
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   3
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GO"
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   8160
      Width           =   300
   End
   Begin VB.Shape star 
      BorderColor     =   &H000000FF&
      Height          =   45
      Left            =   4920
      Top             =   3240
      Width           =   45
   End
   Begin VB.Label Label3 
      Caption         =   "Char. #s:"
      Height          =   255
      Left            =   5880
      TabIndex        =   31
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Random Key:"
      Height          =   255
      Left            =   10080
      TabIndex        =   28
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Copies:"
      Height          =   255
      Left            =   10080
      TabIndex        =   27
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Menu ab 
      Caption         =   "About"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "veda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mx
Dim my
Dim counted
Private Sub ab_Click()
'Dim k
'k = MsgBox("Created By Martin McCormick", vbOKOnly + vbInformation, "About")
End Sub

Private Sub Command1_Click()
On Error GoTo kkkL:

veda.Cls
For i = 0 To 23
If Text1(i).Text = "" Then Text1(i).Text = 0
Next i
Dim lasty
Dim lastx
lastx = star.Left + 12
lasty = star.Top + 12
For d = 1 To Text2.Text
Line (lastx, lasty)-(lastx + (Text1(0).Text * 100), lasty)
lasty = lasty
lastx = lastx + (Text1(0).Text * 100)
Line (lastx, lasty)-(lastx, lasty - (Text1(1).Text * 100))
lastx = lastx
lasty = lasty - (Text1(1).Text * 100)
Line (lastx, lasty)-(lastx - (Text1(2).Text * 100), lasty)
lasty = lasty
lastx = lastx - (Text1(2).Text * 100)
Line (lastx, lasty)-(lastx, lasty + (Text1(3).Text * 100))
lastx = lastx
lasty = lasty + (Text1(3).Text * 100)

'''
Line (lastx, lasty)-(lastx + (Text1(4).Text * 100), lasty)
lasty = lasty
lastx = lastx + (Text1(4).Text * 100)
Line (lastx, lasty)-(lastx, lasty - (Text1(5).Text * 100))
lastx = lastx
lasty = lasty - (Text1(5).Text * 100)
Line (lastx, lasty)-(lastx - (Text1(6).Text * 100), lasty)
lasty = lasty
lastx = lastx - (Text1(6).Text * 100)
Line (lastx, lasty)-(lastx, lasty + (Text1(7).Text * 100))
lastx = lastx
lasty = lasty + (Text1(7).Text * 100)
'''
Line (lastx, lasty)-(lastx + (Text1(8).Text * 100), lasty)
lasty = lasty
lastx = lastx + (Text1(8).Text * 100)
Line (lastx, lasty)-(lastx, lasty - (Text1(9).Text * 100))
lastx = lastx
lasty = lasty - (Text1(9).Text * 100)
Line (lastx, lasty)-(lastx - (Text1(10).Text * 100), lasty)
lasty = lasty
lastx = lastx - (Text1(10).Text * 100)
Line (lastx, lasty)-(lastx, lasty + (Text1(11).Text * 100))
lastx = lastx
lasty = lasty + (Text1(11).Text * 100)
'''
Line (lastx, lasty)-(lastx + (Text1(12).Text * 100), lasty)
lasty = lasty
lastx = lastx + (Text1(12).Text * 100)
Line (lastx, lasty)-(lastx, lasty - (Text1(13).Text * 100))
lastx = lastx
lasty = lasty - (Text1(13).Text * 100)
Line (lastx, lasty)-(lastx - (Text1(14).Text * 100), lasty)
lasty = lasty
lastx = lastx - (Text1(14).Text * 100)
Line (lastx, lasty)-(lastx, lasty + (Text1(15).Text * 100))
lastx = lastx
lasty = lasty + (Text1(15).Text * 100)
'''
Line (lastx, lasty)-(lastx + (Text1(16).Text * 100), lasty)
lasty = lasty
lastx = lastx + (Text1(16).Text * 100)
Line (lastx, lasty)-(lastx, lasty - (Text1(17).Text * 100))
lastx = lastx
lasty = lasty - (Text1(17).Text * 100)
Line (lastx, lasty)-(lastx - (Text1(18).Text * 100), lasty)
lasty = lasty
lastx = lastx - (Text1(18).Text * 100)
Line (lastx, lasty)-(lastx, lasty + (Text1(19).Text * 100))
lastx = lastx
lasty = lasty + (Text1(19).Text * 100)
'''
Line (lastx, lasty)-(lastx + (Text1(20).Text * 100), lasty)
lasty = lasty
lastx = lastx + (Text1(20).Text * 100)
Line (lastx, lasty)-(lastx, lasty - (Text1(21).Text * 100))
lastx = lastx
lasty = lasty - (Text1(21).Text * 100)
Line (lastx, lasty)-(lastx - (Text1(22).Text * 100), lasty)
lasty = lasty
lastx = lastx - (Text1(22).Text * 100)
Line (lastx, lasty)-(lastx, lasty + (Text1(23).Text * 100))
lastx = lastx
lasty = lasty + (Text1(23).Text * 100)
Next d
'''
GoTo k:
kkkL:
s = MsgBox("error", vbOKOnly + vbExclamation, "Error")
Call Command3_Click
k:
End Sub

Private Sub Command2_Click()
For i = 0 To 23
Text1(i).Text = Int(Rnd * Text3.Text)
Next i
End Sub

Private Sub Command3_Click()
veda.Cls
counted = 0
For i = 0 To 23
Text1(i).Text = ""
Next i
Text1(0).SetFocus
End Sub

Private Sub Form_Click()
star.Left = mx
star.Top = my
End Sub

Private Sub Form_Load()
counted = 0
For i = 0 To 23
Text1(i).Text = ""
Next i
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
mx = x
my = Y
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
Dim q
q = Index + 1
counted = counted + 1
If Not q = 24 Then
If counted = Text4.Text Then
Text1(q).SetFocus
counted = 0
End If
End If
End Sub

Private Sub Text4_Change()
counted = 0
End Sub

Private Sub Timer1_Timer()
Command1_Click
Command2_Click
End Sub

Private Sub Timer2_Timer()
ab.Visible = False
Timer2.Enabled = False
End Sub
