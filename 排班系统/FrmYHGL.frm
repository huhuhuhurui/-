VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmYHGL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户管理"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11565
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   9015
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmYHGL.frx":0000
         Left            =   1680
         List            =   "FrmYHGL.frx":000A
         TabIndex        =   8
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         Height          =   1575
         Left            =   5160
         TabIndex        =   1
         Top             =   240
         Width           =   3375
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "退出"
            Height          =   495
            Index           =   4
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "删除"
            Enabled         =   0   'False
            Height          =   495
            Index           =   2
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "修改"
            Enabled         =   0   'False
            Height          =   495
            Index           =   1
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "添加"
            Height          =   495
            Index           =   0
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户权限"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   480
         TabIndex        =   11
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户账号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   480
         TabIndex        =   9
         Top             =   840
         Width           =   840
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   4335
      Left            =   960
      TabIndex        =   12
      Top             =   2880
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7646
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmYHGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click(Index As Integer)
    Select Case Index
           Case 0
           ADDData
           Case 1
           UPTData
           Case 2
           DELData

           Case 4
           Unload Me
    End Select
End Sub

Sub ShowMSH()

Dim rs As New ADODB.Recordset
rs.Open "select * from 用户", Cnn
Set MSH.DataSource = rs
MSH.ColWidth(1) = 0
rs.Close


End Sub


Private Sub ADDData()
On Error GoTo ErrH
    Dim rs As New ADODB.Recordset
    rs.Open "select * from 用户", Cnn, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs.Fields(0) = Trim(Text1(0))
    rs.Fields(1) = Trim(Text1(1))
    rs.Fields(2) = Trim(Combo1)
    rs.Update
    rs.Close
    Set rs = Nothing
    ShowMSH
    MsgBox "添加成功！"
    Exit Sub
ErrH:
    MsgBox Err.Description

End Sub
Private Sub UPTData()
On Error GoTo ErrH
    Dim rs As New ADODB.Recordset
    rs.Open "select * from 用户 where 账号='" & Trim(MSH.TextMatrix(MSH.RowSel, 0)) & "'", Cnn, adOpenDynamic, adLockOptimistic
    
    rs.Fields(0) = Trim(Text1(0))
    rs.Fields(1) = Trim(Text1(1))
    rs.Fields(2) = Trim(Combo1)
    rs.Update
    rs.Close
    Set rs = Nothing
    ShowMSH
            Command1(1).Enabled = False
        Command1(2).Enabled = False
    MsgBox "修改成功！"
    Exit Sub
ErrH:
    MsgBox Err.Description

End Sub
Private Sub DELData()
On Error GoTo ErrH
    Dim rs As New ADODB.Recordset
    rs.Open "select * from 用户 where 账号='" & Trim(MSH.TextMatrix(MSH.RowSel, 0)) & "'", Cnn, adOpenDynamic, adLockOptimistic
    
    rs.Delete
    rs.Update
    rs.Close
    Set rs = Nothing
    ShowMSH
        Command1(1).Enabled = False
        Command1(2).Enabled = False
    MsgBox "删除成功！"
    Exit Sub
ErrH:
    MsgBox Err.Description

End Sub

Private Sub Form_Load()
    ShowMSH
End Sub

Private Sub MSH_Click()
    If MSH.RowSel >= 1 Then
        
        Text1(0) = MSH.TextMatrix(MSH.RowSel, 0)
        Text1(1) = MSH.TextMatrix(MSH.RowSel, 1)
        Combo1 = MSH.TextMatrix(MSH.RowSel, 2)
        Command1(1).Enabled = True
        Command1(2).Enabled = True
        
    End If
End Sub

