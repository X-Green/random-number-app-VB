VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ѧ��"
   ClientHeight    =   5100
   ClientLeft      =   6045
   ClientTop       =   4500
   ClientWidth     =   7800
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1 Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7800
   Begin VB.CommandButton Command1 
      Caption         =   "��һ��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2280
      Top             =   2760
   End
   Begin VB.CommandButton ��ѧ�� 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����� 1 - 47"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3960
      TabIndex        =   0
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Preview 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Preview 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label ��� 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   96
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim F As String, S As String



Private Sub Form_Load()
    ' Shell "cmd.exe /c Md C:\ClassHelper\RandomNumber\"
    CFGPath = ("C:\ClassHelper\RandomNumber\" & "cfg.ini")
    GetInfFrom_ini
    S = "�����1-" & T & "��"
    ��ѧ��.Caption = S
    Form1.BackColor = RGB(r1, g1, b1)
    Functions.Show
    If T < 6 Then
        MsgBox ("��һ��ʹ��ǰ��������")
        Form3.Show
        'Form1.Hide
    End If
    
    numbersonce = True
    MakeList T, NU
    NameNumber(0) = " "
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub



Private Sub Timer1_Timer()
    Functions.Left = Form1.Left + Form1.Width
    Functions.Top = Form1.Top
End Sub

Public Sub ��ѧ��_Click()
    ���.Caption = Final_String
    Preview(0).Caption = FinalNumber
End Sub

