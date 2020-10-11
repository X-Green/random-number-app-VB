Attribute VB_Name = "Module_API"
Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
   ByVal lpKeyName As Any, ByVal lpDefault As String, _
   ByVal lpReturnedString As String, ByVal nSize As Long, _
   ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" _
   Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
   ByVal lpKeyName As Any, ByVal lpString As Any, _
   ByVal lpFileName As String) As Long

