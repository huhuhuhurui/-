VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmZBYSZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "值班员设置"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   13470
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog cd 
      Left            =   7200
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "(Excel)|*.xls"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "导入"
      Height          =   495
      Left            =   10560
      TabIndex        =   17
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   11775
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
         Index           =   4
         Left            =   1560
         TabIndex        =   15
         Top             =   1920
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
         Index           =   3
         Left            =   5640
         TabIndex        =   14
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "添加"
         Height          =   495
         Index           =   0
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "修改"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "删除"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "返回"
         Height          =   495
         Index           =   4
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "FrmZBYSZ.frx":0000
         Left            =   1560
         List            =   "FrmZBYSZ.frx":000A
         TabIndex        =   9
         Text            =   "女"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   4
         Top             =   1320
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
         Left            =   5640
         TabIndex        =   3
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
         Index           =   0
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "值班顺序"
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
         Index           =   3
         Left            =   480
         TabIndex        =   16
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
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
         Index           =   2
         Left            =   4560
         TabIndex        =   8
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
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
         Index           =   1
         Left            =   480
         TabIndex        =   7
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
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
         Left            =   4560
         TabIndex        =   6
         Top             =   720
         Width           =   480
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
         TabIndex        =   5
         Top             =   720
         Width           =   480
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   4815
      Left            =   720
      TabIndex        =   0
      Top             =   3480
      Width           =   11775
      _ExtentX        =   20770
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
Attribute VB_Name = "FrmZBYSZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ShowMSH()
    
    Dim rs As New ADODB.Recordset
    rs.Open "select * from 值班员", Cnn
    Set MSH.DataSource = rs
   
    MSH.ColWidth(0) = 2000
    MSH.ColWidth(1) = 2000
    MSH.ColWidth(2) = 2000
    MSH.ColWidth(3) = 2000
    MSH.ColWidth(4) = 2000
    
    MSH.ColAlignment(4) = 1
    MSH.ColAlignmentFixed(4) = 1
    
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

Private Sub Command2_Click()
        cd.ShowOpen
        If cd.FileName = "" Then
            
            MsgBox "请选择值班员花名册！"
            Exit Sub
            
        End If
        Dim ExcelApp As Object
        Dim ExcelWorkBook As Object
        Dim ExcelWorkSheet As Object
On Error GoTo ErrH
        Set ExcelApp = CreateObject("Excel.Application")
        Set ExcelWorkBook = ExcelApp.Workbooks.Open(cd.FileName)
        Set ExcelWorkSheet = ExcelWorkBook.Worksheets(1)
        
        ExcelApp.Visible = True
        
        For i = 2 To 10000
            If ExcelWorkSheet.cells(i, 1) = "" Then Exit For
            Cnn.Execute "insert into 值班员 values('" & Trim(ExcelWorkSheet.cells(i, 1)) & "','" & Trim(ExcelWorkSheet.cells(i, 2)) & "','" & Trim(ExcelWorkSheet.cells(i, 3)) & "','" & Trim(ExcelWorkSheet.cells(i, 4)) & "'," & Trim(ExcelWorkSheet.cells(i, 5)) & ")"
        Next
        
        MsgBox "ok"
        ShowMSH
    Exit Sub
ErrH:
    
    ExcelApp.Quit
    Set ExcelApp = Nothing
    Set ExcelWorkBook = Nothing
    Set ExcelWorkSheet = Nothing
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    ShowMSH
    

    
End Sub

Private Sub ADDData()
    On Error GoTo ErrH
    
    Text1(2) = Combo2.Text
    Dim rs As New ADODB.Recordset
    rs.Open "select * from 值班员", Cnn, adOpenDynamic, adLockOptimistic
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
    Text1(2) = Combo2.Text
    rs.Open "select * from 值班员 where 编号='" & Trim(MSH.TextMatrix(MSH.RowSel, 0)) & "'", Cnn, adOpenDynamic, adLockOptimistic
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
    rs.Open "select * from 值班员 where 编号='" & Trim(MSH.TextMatrix(MSH.RowSel, 0)) & "'", Cnn, adOpenDynamic, adLockOptimistic
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
        Combo2.Text = Trim(MSH.TextMatrix(MSH.RowSel, 2))
        

        Command1(1).Enabled = True
        Command1(2).Enabled = True
    End If
End Sub







