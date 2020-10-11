VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00EDECCB&
   Caption         =   "偏好设置"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6660
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   4605
   ScaleWidth      =   6660
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Cancel 
      Caption         =   "放弃修改"
      Height          =   615
      Left            =   240
      TabIndex        =   20
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox TextC 
      Height          =   270
      Index           =   1
      Left            =   4200
      TabIndex        =   18
      Text            =   "255"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox TextC 
      Height          =   270
      Index           =   2
      Left            =   4200
      TabIndex        =   17
      Text            =   "255"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox TextC 
      Height          =   270
      Index           =   3
      Left            =   4200
      TabIndex        =   16
      Text            =   "255"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton ColourChangewithID 
      Height          =   735
      Left            =   5160
      TabIndex        =   15
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton skinchange 
      Caption         =   "一键换肤"
      Height          =   735
      Left            =   2760
      TabIndex        =   14
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Back 
      Caption         =   "恢复默认"
      Height          =   735
      Left            =   2760
      TabIndex        =   13
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Text            =   "3"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Text            =   "4"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Text            =   "17"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Text            =   "0"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Text            =   "0"
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton SaveChanges 
      Caption         =   "保存修改"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Text            =   "1"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "姓名修改"
      Height          =   975
      Left            =   4320
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "关于..."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "↑缺席/不在↑"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "↑班级总人数↑"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EDEDED&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "外观"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "学号相关设置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1800
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim Labeldown As Integer




Private Sub Back_Click()
    Form1.BackColor = RGB(255, 255, 255)
    TextC(1).Text = 255
    TextC(2).Text = 255
    TextC(3).Text = 255
    Dim success As Long
    success = WritePrivateProfileString("ColourRGB", "R", TextC(1).Text, CFGPath)
    success = WritePrivateProfileString("ColourRGB", "G", TextC(2).Text, CFGPath)
    success = WritePrivateProfileString("ColourRGB", "B", TextC(3).Text, CFGPath)
End Sub

Private Sub Cancel_Click()
    Form3.Hide
    Form1.Show
    Functions.Show
End Sub

Private Sub ColourChangewithID_Click()
    r1 = TextC(1).Text
    g1 = TextC(2).Text
    b1 = TextC(3).Text
    If r1 > 255 Or g1 > 255 Or b1 > 255 Or r1 < 0 Or g1 < 0 Or b1 < 0 Then
    MsgBox ("颜色的值只能是0与255间的整数")
    Else
    Form1.BackColor = RGB(r1, g1, b1)
    Dim success As Long
    success = WritePrivateProfileString("ColourRGB", "R", TextC(1).Text, CFGPath)
    success = WritePrivateProfileString("ColourRGB", "G", TextC(2).Text, CFGPath)
    success = WritePrivateProfileString("ColourRGB", "B", TextC(3).Text, CFGPath)
    End If
End Sub

Private Sub Command1_Click()
    Dim SN As String
    Dim i As Integer
    Dim success As Boolean
    Dim iString As String
    Shell "cmd.exe /c Md C:\ClassHelper\RandomNumber\"
    For i = 1 To T
        iString = i
        success = WritePrivateProfileString("姓名", iString, NameNumber(i), "C:\ClassHelper\RandomNumber\姓名.ini")
    Next i
    success = Shell("Explorer " & "C:\ClassHelper\RandomNumber\姓名.ini", vbNormalFocus)
    SaveChanges_Click
    GetInfFrom_ini
End Sub


Private Sub Form_Load()
    Label1_MouseDown 1, 1, 1, 1, 1
'===================下面是ini里的信息==========
    Dim ret As Long
    Dim buff As String
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "第一个", "0", buff, 256, CFGPath)
    Text1.Text = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "第二个", "0", buff, 256, CFGPath)
    Text2.Text = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "第三个", "0", buff, 256, CFGPath)
    Text3.Text = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "第四个", "0", buff, 256, CFGPath)
    Text4.Text = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Filtered", "第五个", "0", buff, 256, CFGPath)
    Text5.Text = buff
    buff = String(255, 0)
    ret = GetPrivateProfileString("Numbers_Total", "班级总人数", "47", buff, 256, CFGPath)
    Text6.Text = buff
End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 1 To 3
        If i <> Labeldown Then
            Label1(i).BackColor = RGB(255, 255, 255)
        End If
    Next i
End Sub


Private Sub Label1_MouseDown(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 1 To 3
        Label1(i).BackColor = RGB(255, 255, 255)
    Next i
    Label1(Index).BackColor = RGB(237, 237, 237)
    Labeldown = Index
'====================以上是颜色部分==============
    If Index = 1 Then
        Page1_load
        Page2_unload
        Page3_unload
    Else
        If Index = 2 Then
            Page1_unload
            Page2_load
            Page3_unload
        Else
            If Index = 3 Then
                Page1_unload
                Page2_unload
                Page3_load
            Else
            End If
        End If
    End If
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
    If Index <> Labeldown Then
        Label1(Index).BackColor = RGB(249, 243, 249)
    End If
End Sub

Public Sub SaveChanges_Click()
    Dim i%
    If Text6.Text > 65534 Or Text6.Text < 6 Then
        MsgBox ("保存失败：一个班级不可能有这么多/少的人")
    Else
        n(1) = Text1.Text
        n(2) = Text2.Text
        n(3) = Text3.Text
        n(4) = Text4.Text
        n(5) = Text5.Text
        T = Text6.Text
        Dim WR As Long
        WR = WritePrivateProfileString("Numbers_Filtered", "第一个", Text1.Text, CFGPath)
        WR = WritePrivateProfileString("Numbers_Filtered", "第二个", Text2.Text, CFGPath)
        WR = WritePrivateProfileString("Numbers_Filtered", "第三个", Text3.Text, CFGPath)
        WR = WritePrivateProfileString("Numbers_Filtered", "第四个", Text4.Text, CFGPath)
        WR = WritePrivateProfileString("Numbers_Filtered", "第五个", Text5.Text, CFGPath)
        WR = WritePrivateProfileString("Numbers_Total", "班级总人数", Text6.Text, CFGPath)
        Form3.Hide
        NU = 0
        For i = 1 To 5
            If n(i) <> 0 Then
            NU = NU + 1
            Else
            End If
        Next i
        MsgBox (("缺席：" & NU) & "人")
        iString = NU
        WR = WritePrivateProfileString("Numbers_Filtered", "缺席總數", iString, CFGPath)
        Functions.Check1.Value = 0
        numbersonce = False
        Form1.抽学号.Caption = "随机：1-" & T & "号"
    End If
End Sub
Public Sub Page1_load()
    Text1.Visible = True
    Text2.Visible = True
    Text3.Visible = True
    Text4.Visible = True
    Text5.Visible = True
    Command1.Visible = True
    Text6.Visible = True
    Label4.Visible = True
    Label3.Visible = True
    
End Sub

Public Sub Page1_unload()
    Text1.Visible = False
    Text2.Visible = False
    Text3.Visible = False
    Text4.Visible = False
    Text5.Visible = False
    Command1.Visible = False
    Text6.Visible = False
    Label3.Visible = False
    Label4.Visible = False

End Sub

Public Sub Page2_load()
    Back.Visible = True
    skinchange.Visible = True
    TextC(1).Visible = True
    TextC(2).Visible = True
    TextC(3).Visible = True
    ColourChangewithID.Visible = True
End Sub

Public Sub Page2_unload()
    Back.Visible = False
    skinchange.Visible = False
    TextC(1).Visible = False
    TextC(2).Visible = False
    TextC(3).Visible = False
    ColourChangewithID.Visible = False

End Sub

Private Sub skinchange_Click()
    r1 = Int(Rnd() * 256)
    g1 = Int(Rnd() * 256)
    b1 = Int(Rnd() * 256)
    Form1.BackColor = RGB(r1, g1, b1)
    TextC(1).Text = r1
    TextC(2).Text = g1
    TextC(3).Text = b1
    Dim success As Integer
    success = WritePrivateProfileString("ColourRGB", "R", TextC(1).Text, CFGPath)
    success = WritePrivateProfileString("ColourRGB", "G", TextC(2).Text, CFGPath)
    success = WritePrivateProfileString("ColourRGB", "B", TextC(3).Text, CFGPath)
End Sub

Public Sub Page3_load()
    Label2.Caption = "作者：戎羿190632                             Copyleft(c) LightningCreeper.Inc. All rights free"
End Sub

Public Sub Page3_unload()
    Label2.Caption = ""
End Sub
