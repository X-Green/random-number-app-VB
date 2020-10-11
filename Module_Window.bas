Attribute VB_Name = "Module_Window"
Option Explicit
Public r1 As Integer, g1 As Integer, b1 As Integer
Public CFGPath As String

Public Sub DragTitle(frm As Form, Button As Integer, shift As Integer, X As Single, Y As Single)
    Static i!, j!
    Dim a!, b!
    If Button = 1 Then
        a = frm.Left - i + X
        b = frm.Top - j + Y
        frm.Move a, b
    Else
        i = X
        j = Y
    End If
End Sub
