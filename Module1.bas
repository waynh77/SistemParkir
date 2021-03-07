Attribute VB_Name = "Module1"
Public db As String
Public dbCon As String
Public db2 As String
Public dbCon2 As String
Public jmlrow As Single
Public cuRR As String
Public sql1 As String

Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As _
Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName _
As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
      (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Sub bukadaTa()
db = "Akseskontrol"
dbCon = "odbc;uid=sa;pwd=eq-park;database=akseskontrol;dsn=akseskontrol"
End Sub

Public Function kutip(Kalimat As String) As String
Kalimat = Replace(Kalimat, "'", "''")
kutip = "'" & Kalimat & "'"
End Function

Public Sub AddFilter(ByVal dlg As CommonDialog, ByVal _
    filter_title As String, ByVal filter_value As String)
    Dim txt As String

    txt = dlg.Filter
    If Len(txt) > 0 Then txt = txt & "|"
    txt = txt & filter_title & " (" & filter_value & ")|" & _
        filter_value
    dlg.Filter = txt
End Sub

