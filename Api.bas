Attribute VB_Name = "Api"
Option Explicit
'Const&Functions used for the FormMove methods
Public Const LP_HT_CAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

