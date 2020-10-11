VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   3825
   ClientLeft      =   7395
   ClientTop       =   5145
   ClientWidth     =   5475
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5475
   Begin VB.CommandButton Save 
      Caption         =   "OK"
      Height          =   735
      Left            =   5040
      TabIndex        =   6
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox TextC 
      Height          =   270
      Index           =   3
      Left            =   4080
      TabIndex        =   5
      Text            =   "255"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox TextC 
      Height          =   270
      Index           =   2
      Left            =   4080
      TabIndex        =   4
      Text            =   "255"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox TextC 
      Height          =   270
      Index           =   1
      Left            =   4080
      TabIndex        =   3
      Text            =   "255"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Back 
      Caption         =   "恢复默认"
      Height          =   735
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton skinchange 
      Caption         =   "一键换肤"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton Numbers 
      Caption         =   "学号相关"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Shape Shape4 
      Height          =   135
      Left            =   5040
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape Shape3 
      Height          =   135
      Left            =   4560
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape Shape2 
      Height          =   135
      Left            =   4320
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      FillColor       =   &H00404040&
      Height          =   135
      Left            =   4080
      Top             =   2280
      Width           =   135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Back_Click()
    Form1.BackColor = RGB(255, 255, 255)
    Text1.Text = 255
    Text2.Text = 255
    Text3.Text = 255
    Dim success As Long
    success = WritePrivateProfileString("ColourRGB", "R", Text1.Text, "c:aa.ini")
    success = WritePrivateProfileString("ColourRGB", "G", Text2.Text, "c:aa.ini")
    success = WritePrivateProfileString("ColourRGB", "B", Text3.Text, "c:aa.ini")

End Sub

Private Sub Numbers_Click()
    Form4.Show vbModel
End Sub

Private Sub Save_Click()
    r1 = TextC(1).Text
    g1 = TextC(2).Text
    b1 = TextC(3).Text
    If r1 > 255 Or g1 > 255 Or b1 > 255 Or r1 < 0 Or g1 < 0 Or b1 < 0 Then
    MsgBox ("颜色的值只能是0与255间的整数")
    Else
    Form1.BackColor = RGB(r1, g1, b1)
    Dim success As Long
    success = WritePrivateProfileString("ColourRGB", "R", TextC(1).Text, "c:aa.ini")
    success = WritePrivateProfileString("ColourRGB", "G", TextC(2).Text, "c:aa.ini")
    success = WritePrivateProfileString("ColourRGB", "B", TextC(3).Text, "c:aa.ini")
    End If
    
End Sub

Private Sub skinchange_Click()
    r1 = Int(Rnd() * 256)
    g1 = Int(Rnd() * 256)
    b1 = Int(Rnd() * 256)
    Form1.BackColor = RGB(r1, g1, b1)
    Text1.Text = r1
    Text2.Text = g1
    Text3.Text = b1
    Dim success As Long
    success = WritePrivateProfileString("ColourRGB", "R", Text1.Text, "c:aa.ini")
    success = WritePrivateProfileString("ColourRGB", "G", Text2.Text, "c:aa.ini")
    success = WritePrivateProfileString("ColourRGB", "B", Text3.Text, "c:aa.ini")
End Sub
