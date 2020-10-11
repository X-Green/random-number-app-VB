VERSION 5.00
Begin VB.Form Functions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   $"Functions.frx":0000
   ClientHeight    =   4275
   ClientLeft      =   14280
   ClientTop       =   4965
   ClientWidth     =   1335
   LinkTopic       =   "Form5"
   ScaleHeight     =   4275
   ScaleWidth      =   1335
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "学号不重复"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PPT模式"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "设置"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   4275
      Left            =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Functions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    numbersonce = Not numbersonce
    If numbersonce Then
        MakeList T, NU
        MsgBox ("已生成列表")
    End If
End Sub

Private Sub Command1_Click()
    Form3.Show vbModeless
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    Form2.Show vbModeless
    Form1.Hide
    Functions.Hide
End Sub

