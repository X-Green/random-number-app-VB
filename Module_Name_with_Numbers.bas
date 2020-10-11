Attribute VB_Name = "Module_Name_with_Numbers"
Option Explicit

Public NN(1 To 65535) As String

Public Function GetNameNumber() As String
    Dim ret As Long, buff As String, i%, SE As String
    For i = 1 To T
        SE = i
        buff = String(255, 0)
        ret = GetPrivateProfileString("N", SE, "", buff, 256, "c:aa.ini")
        NN(i) = buff
    Next i
End Function

Public Function ShowNameNumber(x As Integer) As String
    ShowNameNumber = NN(x)
End Function
