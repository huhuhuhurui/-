VERSION 5.00
Begin VB.Form FrmLogin 
   Caption         =   "登录"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   13545
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "登陆"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7440
      TabIndex        =   3
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10080
      TabIndex        =   2
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   8760
      PasswordChar    =   "#"
      TabIndex        =   1
      Text            =   "001"
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Text            =   "001"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   8055
      Left            =   0
      Picture         =   "FrmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "排班系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   720
      Left            =   8160
      TabIndex        =   6
      Top             =   960
      Width           =   2940
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   8040
      TabIndex        =   5
      Top             =   3600
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "账号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   8040
      TabIndex        =   4
      Top             =   2880
      Width           =   510
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim rs As New ADODB.Recordset   '声明记录集
    rs.Open "select * from 用户 where 账号='" & Trim(Text1.Text) & "' and 密码='" & Trim(Text2.Text) & "'", Cnn   '查找用户表
    If rs.EOF = True Then    '如果查不到，提示错误
        MsgBox "密码错误或者无此用户", vbCritical
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
     '找到后，记录账号密码和权限
    strXM = Trim(rs.Fields(0).Value)
    strMM = Trim(rs.Fields(1).Value)
    strQX = Trim(rs.Fields(2).Value)
    rs.Close  '关闭记录 释放资源
    Set rs = Nothing
    
    FrmMain.Show   '显示主窗体
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


