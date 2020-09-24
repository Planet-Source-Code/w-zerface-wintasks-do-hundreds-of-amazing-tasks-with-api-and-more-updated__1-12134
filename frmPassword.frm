VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "* Find"
   ClientHeight    =   360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1200
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   1200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReady 
      Caption         =   "&Ready"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtDelay 
      Height          =   285
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "2"
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4200
      Top             =   0
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Time Delay "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' This API used to get the current mouse pointer location
Private Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long

' This API is used to get the window handle by giving the
' location
Private Declare Function WindowFromPoint Lib "user32" _
(ByVal xPoint As Long, ByVal yPoint As Long) As Long

' By passing window handle this will return the class name of
' window object
Private Declare Function GetClassName Lib "user32" Alias _
"GetClassNameA" (ByVal hWnd As Long, ByVal _
lpClassName As String, ByVal nMaxCount As Long) As Long

'This function is used to send a message to another window
Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

' This is a data type to store location
Private Type POINTAPI
        x As Long
        Y As Long
End Type

' We want to send a message that we need to set the password
' char of that object to zero. That means it will apear in
' normal text. This is the message
Private Const EM_SETPASSWORDCHAR = &HCC

' This procedure is called by timer to do all the operation
Private Sub ShowPassword()

Dim curWindow As Long           'To store window handle
Dim sClassName As String * 255  'To store class name of the obj
Dim sname As String             'Reduced class name
Dim lpPoint As POINTAPI         'To store mouse location

Call GetCursorPos(lpPoint)      'Get current mouse location
'Get window handle from that location
curWindow = WindowFromPoint(lpPoint.x, lpPoint.Y)
'Get class name of that window
Call GetClassName(curWindow, sClassName, 255)
'Well reduced class name
sname = Trim(Left(sClassName, InStr(sClassName, vbNullChar) - 1))

'Check whether this object is Edit control or TextBox if yes
If sname = "Edit" Or InStr(sname, "TextBox") > 0 Then
    'Send message to convert its password char to normal
    Call SendMessage(curWindow, EM_SETPASSWORDCHAR, 0, 0)
    cmdReady.BackColor = vbGreen
    cmdReady.Caption = "Ready"
Else
    'If not say that cannot get the password
    MsgBox "Sorry! Cannot get the password", vbExclamation, "Error !"
End If
Timer1.Enabled = False      ' Stop the timer
End Sub

'Display help message
Private Sub cmdHowTo_Click()

Dim strMessage As String
strMessage = "Set time delay. Then press Ready button. " _
  & vbCrLf & "Then place the mouse pointer on password box" _
  & vbCrLf & "of any application." _
  & vbCrLf & "After the specified time Status Message will tell " _
  & vbCrLf & "you that you're now ready to get the password." _
  & vbCrLf & "Then click that password box, " _
  & vbCrLf & "it will apear in normal text."
           
MsgBox strMessage, vbOKOnly + vbInformation, "Password Hacker How to"
End Sub

Private Sub cmdReady_Click()
Timer1.Interval = Val(txtDelay.Text) * 1000
Timer1.Enabled = True
cmdReady.BackColor = vbRed
cmdReady.Caption = "Wait"
End Sub

Private Sub Form_Load()
'cmdReady.BackColor = vbRed
'lblStatus.Caption = "Wait"
End Sub

Private Sub Timer1_Timer()
Call ShowPassword
End Sub

Private Sub txtDelay_LostFocus()
If Not IsNumeric(txtDelay.Text) Then
    MsgBox "Please enter numeric value", vbInformation, "Error !"
    txtDelay = 5
End If
End Sub
