VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "PPTing"
   ClientHeight    =   2745
   ClientLeft      =   18270
   ClientTop       =   3090
   ClientWidth     =   2355
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Red 
      BackColor       =   &H00000000&
      Caption         =   "∑µªÿ"
      Height          =   375
      Left            =   240
      MaskColor       =   &H00000000&
      Picture         =   "Form2.frx":10CA
      TabIndex        =   1
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton Green 
      BackColor       =   &H00000000&
      Caption         =   "≥È—ß∫≈"
      Height          =   375
      Left            =   240
      MaskColor       =   &H00808080&
      Picture         =   "Form2.frx":1BB4
      TabIndex        =   0
      Top             =   1680
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.Label Frame1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   645
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Const HWND_BOTTOM = 1
Const SWP_NOMOVE = &H2

Private Sub Form_Load()
   Me.BackColor = &HFF0000
   Dim rtn As Long
   Dim BorderStyler
   BorderStyler = 0
   rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
   rtn = rtn Or WS_EX_LAYERED
   SetWindowLong hwnd, GWL_EXSTYLE, rtn
   SetLayeredWindowAttributes hwnd, &HFF0000, 100, LWA_COLORKEY


    SetWindowPos Form2.hwnd, -1, 0, 0, 0, 0, 3
    Dim X As Integer, Y As Integer
    X = Screen.Width / Screen.TwipsPerPixelX
    Y = Screen.Height / Screen.TwipsPerPixelY
    Form2.Left = X * 15 - Form2.Width
    Form2.Top = Y * 15 * 0.382 - Form2.Height
End Sub

Private Sub Green_Click()
    Frame1.Caption = Final_String
End Sub

Private Sub Red_Click()
    Form1.Show
    Form2.Hide
    Functions.Show
End Sub

