VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form4"
   ClientHeight    =   4665
   ClientLeft      =   6990
   ClientTop       =   4935
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4665
   ScaleWidth      =   6045
   Begin VB.CommandButton Command1 
      Caption         =   "姓名修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   6
      Text            =   "1"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton SaveChanges 
      Caption         =   "保存修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Text            =   "0"
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Text            =   "0"
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Text            =   "17"
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "4"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "3"
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "↑班级总人数↑"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim SN As String
    Dim i As Integer
    Dim iString As String
    Dim success As Boolean
    For i = 1 To T
        iString = i
        SN = InputBox(iString)
        NameNumber(i) = SN
        success = WritePrivateProfileString("姓名", iString, SN, "c:aa.ini")
    Next i
End Sub

Private Sub Form_Load()
    Dim ret As Long
    Dim buff As String
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "第一个", "0", buff, 256, "c:aa.ini")
    Text1.Text = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "第二个", "0", buff, 256, "c:aa.ini")
    Text2.Text = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "第三个", "0", buff, 256, "c:aa.ini")
    Text3.Text = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "第四个", "0", buff, 256, "c:aa.ini")
    Text4.Text = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "第五个", "0", buff, 256, "c:aa.ini")
    Text5.Text = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Total", "班级总人数", "47", buff, 256, "c:aa.ini")
    Text6.Text = buff
End Sub

Public Sub SaveChanges_Click()
    Dim i%
    If Text6.Text > 65534 Or Text6.Text < 6 Then
        MsgBox ("保存失败：             泥垢了")
    Else
        n(1) = Text1.Text
        n(2) = Text2.Text
        n(3) = Text3.Text
        n(4) = Text4.Text
        n(5) = Text5.Text
        T = Text6.Text
        Dim WR As Long
        WR = WritePrivateProfileString("Numbers_Filtered", "第一个", Text1.Text, "c:aa.ini")
        WR = WritePrivateProfileString("Numbers_Filtered", "第二个", Text2.Text, "c:aa.ini")
        WR = WritePrivateProfileString("Numbers_Filtered", "第三个", Text3.Text, "c:aa.ini")
        WR = WritePrivateProfileString("Numbers_Filtered", "第四个", Text4.Text, "c:aa.ini")
        WR = WritePrivateProfileString("Numbers_Filtered", "第五个", Text5.Text, "c:aa.ini")
        WR = WritePrivateProfileString("Numbers_Total", "班级总人数", Text6.Text, "c:aa.ini")
        Form3.Hide
        Form4.Hide
        NU = 0
        For i = 1 To 5
            If n(i) <> 0 Then
            NU = NU + 1
            Else
            End If
        Next i
        MsgBox (("缺席：" & NU) & "人")
        Form1.Show vbModel
        Functions.Check1.Value = 0
        numbersonce = False
        S = "随机：1-" & T
        Form1.抽学号.Caption = S
    End If
    
End Sub

