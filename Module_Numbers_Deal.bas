Attribute VB_Name = "Module_Numbers_Deal"
Option Explicit

Public T As Integer
Public numbersonce As Boolean
Public X(1 To 32767) As Integer
Public Time As Integer
Public n%(1 To 5)
Public NU%
Public NameNumber(0 To 32767) As String
Public iString As String
Public FinalNumber As Integer

'���������1-T
Public Function Ran(T) As Integer
    Dim X As Integer
    Do
    Randomize
    X = Int(T * Rnd + 1)
    Loop Until X <> n(1) And X <> n(2) And X <> n(3) And X <> n(4) And X <> n(5)
    Ran = X
End Function
'�����б� 1 - T���ظ�
Public Sub MakeList(T As Integer, NU As Integer)
    Dim a%(1 To 32767)
    Dim b%, i%, ii%, iii%, myvalue%
    For i = 1 To T
        a(i) = i
    Next i
    For iii = 1 To T - NU
        Do
        b = Ran(T)
        X(iii) = a(b)
        Loop Until X(iii) <> 0
        a(b) = 0
    Next iii
    Time = 0
End Sub

'����Ran()
'��ʾ�б�

Public Function NumberListShow() As Integer
    If X(Time + 1) <> 0 Then
        Time = Time + 1
        NumberListShow = X(Time)
    Else: MsgBox ("��ǰ�б��Ѿ���ʾ��")
        MakeList T, NU
        Time = 0
    End If

End Function



Public Function Final_String()
    Dim vhdvdggcg As Integer
    If numbersonce Then
        vhdvdggcg = NumberListShow
    Else
        vhdvdggcg = Ran(T)
    End If
    If NameNumber(vhdvdggcg) = String(255, Left(NameNumber(vhdvdggcg), 1)) Then
        Final_String = vhdvdggcg
    Else
        Final_String = NameNumber(vhdvdggcg)
    End If
    FinalNumber = vhdvdggcg
End Function

Public Sub GetInfFrom_ini()
    Dim ret As Long
    Dim buff As String
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "��һ��", "0", buff, 256, CFGPath)
    n(1) = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "�ڶ���", "0", buff, 256, CFGPath)
    n(2) = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "������", "0", buff, 256, CFGPath)
    n(3) = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "���ĸ�", "0", buff, 256, CFGPath)
    n(4) = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "�����", "0", buff, 256, CFGPath)
    n(5) = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("ColourRGB", "R", "256", buff, 256, CFGPath)
    r1 = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("ColourRGB", "G", "256", buff, 256, CFGPath)
    g1 = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("ColourRGB", "B", "256", buff, 256, CFGPath)
    b1 = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Total", "�༶������", "0", buff, 256, CFGPath)
    T = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "ȱϯ����", "5", buff, 256, CFGPath)
    NU = buff

'===========������������ȡ============
    Dim success As Boolean
    Dim STR As String
    Dim i As Integer
    Dim iString As String
    For i = 1 To T
        STR = String(255, 0)
        iString = i
        success = GetPrivateProfileString("����", iString, "", STR, 255, "C:\ClassHelper\RandomNumber\����.ini")
        NameNumber(i) = STR
    Next i
End Sub


