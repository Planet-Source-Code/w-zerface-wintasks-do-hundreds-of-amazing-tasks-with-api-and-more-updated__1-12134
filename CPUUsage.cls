VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPUUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const REG_DWORD = 4
Private Const HKEY_DYN_DATA = &H80000006
'Initiate the key
Public Sub InitCPUUsage()
    Dim Data As Long, Typ As Long, size As Long
    Dim hKey As Long, hRet As Long

    hRet = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", hKey)
    hRet = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, REG_DWORD, Data, 4)
    hRet = RegCloseKey(hKey)
End Sub
'Get the cpu info via gfx meter
Public Function GetCPUUsage() As Long
    Dim Data As Long, Typ As Long, size As Long
    Dim hKey As Long
    Dim hRet As Long

    hRet = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", hKey)
    hRet = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0&, REG_DWORD, Data, 4)
    GetCPUUsage = Data
    hRet = RegCloseKey(hKey)
End Function

