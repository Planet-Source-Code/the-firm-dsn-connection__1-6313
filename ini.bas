Attribute VB_Name = "Ini"
'******* USE ********

'******* write to ini_file ********

'Private Sub cmdstore_Click()
'Dim lpAppName As String, lpFileName As String, lpKeyName As String, lpString As String
'Dim U As Long
'lpAppName = appname.Text
'lpKeyName = keyname.Text
'lpString = txtvalue.Text
'lpFileName = file
'U = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
'If U = 0 Then
'Beep
'End If
'End Sub


'******** read from ini_file *******

'Private Sub cmdget_Click()
'Dim X As Long
'Dim Temp As String * 50
'Dim lpAppName As String, lpKeyName As String, lpDefault As String, lpFileName As String
'lpAppName = appname.Text
'lpKeyName = keyname.Text
'lpDefault = file
'lpFileName = file
'X = GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, Temp, Len(Temp), lpFileName)

'If X = 0 Then
'    Beep
'Else
'    txtvalue.Text = Trim(Temp)
'End If
'End Sub


Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Global file
Global appname
Global keyname
Global value

Public Sub writeini()
Dim lpAppName As String, lpFileName As String, lpKeyName As String, lpString As String
Dim U As Long
lpAppName = appname
lpKeyName = keyname
lpString = value
lpFileName = file
U = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
If U = 0 Then
Beep
End If
End Sub

Public Sub readini()
Dim X As Long
Dim Temp As String * 50
Dim lpAppName As String, lpKeyName As String, lpDefault As String, lpFileName As String
lpAppName = appname
lpKeyName = keyname
lpDefault = no
lpFileName = file
X = GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, Temp, Len(Temp), lpFileName)

If X = 0 Then
    Beep
Else
    result = Trim(Temp)
End If
End Sub


