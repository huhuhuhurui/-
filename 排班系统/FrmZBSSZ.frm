VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmZBSSZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "值班室设置"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10230
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   7935
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
         Index           =   0
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   2655
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
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "返回"
         Height          =   495
         Index           =   4
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "删除"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "修改"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "添加"
         Height          =   495
         Index           =   0
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   6
         Left            =   480
         TabIndex        =   8
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   1440
         Width           =   480
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   4815
      Left            =   960
      TabIndex        =   9
      Top             =   3120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8493
      _Version        =   393216
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmZBSSZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ShowMSH()
    
    Dim rs As New ADODB.Recordset
    rs.Open "select * from 值班室", Cnn
    Set MSH.DataSource = rs
   
    MSH.ColWidth(0) = 2000
    MSH.ColWidth(1) = 2500



    rs.Close
    Set rs = Nothing
    
    
End Sub

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

Private Sub Form_Load()
    ShowMSH
    
End Sub

Private Sub ADDData()
    On Error GoTo ErrH
    

    Dim rs As New ADODB.Recordset
    rs.Open "select * from 值班室", Cnn, adOpenDynamic, adLockOptimistic
    rs.AddNew
    
        For i = 0 To Text1.Count - 1
        
            rs.Fields(i) = Trim(Text1(i))
        
        Next
        
    rs.Update
    rs.Close

    ShowMSH
    
    
    MsgBox "添加成功！"
  
    Exit Sub
ErrH:
    MsgBox Err.Description

End Sub


Private Sub UPTData()
    On Error GoTo ErrH
    Dim rs As New ADODB.Recordset

    rs.Open "select * from 值班室 where 编号='" & Trim(MSH.TextMatrix(MSH.RowSel, 0)) & "'", Cnn, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then

        For i = 1 To Text1.Count - 1
        
            rs.Fields(i) = Trim(Text1(i))
        
        Next

        rs.Update
        rs.Close
        

        ShowMSH
        Command1(1).Enabled = False
        Command1(2).Enabled = False
        MsgBox "修改成功！"
    Else
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
    Exit Sub
ErrH:
    MsgBox Err.Description

End Sub

Private Sub DELData()
    On Error GoTo ErrH
    Dim rs As New ADODB.Recordset
    rs.Open "select * from 值班室 where 编号='" & Trim(MSH.TextMatrix(MSH.RowSel, 0)) & "'", Cnn, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
        rs.Delete
        rs.Update
        rs.Close

        ShowMSH
        Command1(1).Enabled = False
        Command1(2).Enabled = False
        MsgBox "删除成功！"
    Else
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
    Exit Sub
ErrH:
    MsgBox Err.Description

End Sub

Private Sub MSH_Click()
    If MSH.RowSel >= 1 Then
        
        
        
        For i = 0 To Text1.Count - 1
        Text1(i) = Trim(MSH.TextMatrix(MSH.RowSel, i))
        Next
         

        Command1(1).Enabled = True
        Command1(2).Enabled = True
    End If
End Sub








