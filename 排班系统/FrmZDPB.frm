VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmZDPB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "自动排班"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16395
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   16395
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   495
      Left            =   11640
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始排班"
      Height          =   495
      Left            =   9480
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   10354689
      CurrentDate     =   43579
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   10354689
      CurrentDate     =   43579
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "至"
      Height          =   210
      Left            =   6360
      TabIndex        =   3
      Top             =   1200
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "起止日期"
      Height          =   210
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   840
   End
End
Attribute VB_Name = "FrmZDPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If MsgBox("是否删除历史记录", vbYesNo, "提示") = vbYes Then
    Cnn.Execute "delete from 值班表"
    End If
    Dim ArrZBS() As String
    Dim i As Integer
    
    
    Dim rs As New ADODB.Recordset
    rs.Open "select * from 值班室", Cnn
    i = -1
    If rs.EOF = False Then
        rs.MoveFirst
        Do While Not rs.EOF
            i = i + 1
            ReDim Preserve ArrZBS(i)
            ArrZBS(i) = Trim(rs.Fields("编号"))
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    Dim ArrZBY() As String
    
    rs.Open "select * from 值班员 order by 值班顺序 asc", Cnn
    i = -1
    If rs.EOF = False Then
        rs.MoveFirst
        Do While Not rs.EOF
            i = i + 1
            ReDim Preserve ArrZBY(i)
            ArrZBY(i) = Trim(rs.Fields("编号"))
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    If UBound(ArrZBS()) > UBound(ArrZBY) Then
        MsgBox "值班人数少于值班室数，无法排班！"
        Exit Sub
    End If
    
    Dim ArrSJD() As String
    
    rs.Open "select * from 时间段 order by 序号 asc", Cnn
    i = -1
    If rs.EOF = False Then
        rs.MoveFirst
        Do While Not rs.EOF
            i = i + 1
            ReDim Preserve ArrSJD(i)
            ArrSJD(i) = Trim(rs.Fields("区间"))
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim n As Long
    
    '记录日期差
    n = DateDiff("d", DTPicker1.Value, DTPicker2.Value)
    n = n - 1
    '每天每个时间段对每个值班室进行排班
    
    Dim PBRYBH As String
    Dim PBRQ As Date
    
    For i = 0 To n          '每天
        PBRQ = DateAdd("d", i, DTPicker1.Value)
        For k = 0 To UBound(ArrSJD)     '每个时间段
            For m = 0 To UBound(ArrZBS)     '每个值班室
                
                Do While 1
                    DoEvents
                    PBRYBH = ArrZBY(0)
                    '查看是否请假当前时间段当前人员是否请假
                    rs.Open "select * from 请假表 where 请假日期=#" & PBRQ & "# and 时间段='" & ArrSJD(k) & "' and 人员编号='" & PBRYBH & "'", Cnn
                    If rs.EOF = True Then   '如果没有请假，安排值班
                        rs.Close
                        Cnn.Execute "insert into 值班表(日期,时间段,人员编号,值班室编号) values(#" & PBRQ & "#,'" & ArrSJD(k) & "','" & PBRYBH & "','" & ArrZBS(m) & "')"
                        QY ArrZBY
                        Exit Do
                    End If
                    rs.Close
                    QY ArrZBY
                Loop
            Next
        Next
    Next
    
    MsgBox "排班完毕！"
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Sub QY(ArrSZ() As String)

    Dim strTmp As String
    
    strTmp = ArrSZ(0)
    
    For i = 1 To UBound(ArrSZ)
        
        ArrSZ(i - 1) = ArrSZ(i)
        
    Next
    
    ArrSZ(i - 1) = strTmp
    
    
    
End Sub

Private Sub Command3_Click()

End Sub

Private Sub DTPicker1_Change()
    DTPicker2.Value = DateAdd("m", 1, DTPicker1.Value)
End Sub

Private Sub DTPicker1_Click()
    DTPicker2.Value = DateAdd("m", 1, DTPicker1.Value)
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    DTPicker2.Value = DateAdd("m", 1, Date)
End Sub
