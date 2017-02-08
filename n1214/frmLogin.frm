VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "登陆窗口"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4590
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdExit 
      Caption         =   "取消"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "登陆"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtPassWord 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1200
      Width           =   1500
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   1305
      Left            =   360
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1320
   End
   Begin VB.Label lbDate 
      Caption         =   "Label4"
      Height          =   615
      Left            =   3600
      TabIndex        =   7
      Top             =   2500
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "STI版权所有"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2500
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "密  码："
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "用户名："
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strUserName As String
Dim strPassWord As String
Const username As String = "admin"
Const passWord As String = "666666"

Public Fusername As String


Private Sub Command2_Click()

End Sub

Private Sub cmdExit_Click()
    End
End Sub


Private Sub cmdLogin_Click()
    If txtUser.Text = "" Then
        MsgBox "no user"
        txtUser.SetFocus
        Exit Sub
    Else
        strUserName = Trim(txtUser.Text)
        strPassWord = Trim(txtPassWord.Text)
        If (strUserName = username) And (strPassWord = passWord) Then
            MsgBox "登陆成功", , "登陆提示"
            Fusername = strUserName
            frmMain.Show
            Unload Me
        Else
            MsgBox "验证失败", , "登陆提示"
            txtUser.SetFocus
        End If
    End If
End Sub


Private Sub Form_Load()
    
    strUserName = ""
    strPassWord = ""
    txtUser.Text = ""
    txtPassWord.Text = ""
    txtPassWord.PasswordChar = "*"
    txtPassWord.MaxLength = 6
    lbDate.Caption = Date

End Sub
