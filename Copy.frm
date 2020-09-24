VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy and\or Rename File"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   Icon            =   "Copy.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy and/or &Rename"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "Copy.frx":000C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "File to copy and/or rename:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Copy and/or rename file to:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "First type the path of the file to copy and/or rename and then type the path and new name."
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function copyfile Lib "Kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Sub Command1_Click()
Dim retvaL
Dim abc
retvaL = copyfile(Text1.Text, Text2.Text, 1)
If retvaL = 0 Then abc = MsgBox("Cannot copy/rename " & Text1.Text, vbOKOnly + vbCritical, "Error")
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyCode = 13 Then Call Command1_Click
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyCode = 13 Then Call Command1_Click
End Sub
