VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmZDPB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Զ��Ű�"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16395
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   11640
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ�Ű�"
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
      Caption         =   "��"
      Height          =   210
      Left            =   6360
      TabIndex        =   3
      Top             =   1200
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��ֹ����"
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
    If MsgBox("�Ƿ�ɾ����ʷ��¼", vbYesNo, "��ʾ") = vbYes Then
    Cnn.Execute "delete from ֵ���"
    End If
    Dim ArrZBS() As String
    Dim i As Integer
    
    
    Dim rs As New ADODB.Recordset
    rs.Open "select * from ֵ����", Cnn
    i = -1
    If rs.EOF = False Then
        rs.MoveFirst
        Do While Not rs.EOF
            i = i + 1
            ReDim Preserve ArrZBS(i)
            ArrZBS(i) = Trim(rs.Fields("���"))
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    Dim ArrZBY() As String
    
    rs.Open "select * from ֵ��Ա order by ֵ��˳�� asc", Cnn
    i = -1
    If rs.EOF = False Then
        rs.MoveFirst
        Do While Not rs.EOF
            i = i + 1
            ReDim Preserve ArrZBY(i)
            ArrZBY(i) = Trim(rs.Fields("���"))
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    If UBound(ArrZBS()) > UBound(ArrZBY) Then
        MsgBox "ֵ����������ֵ���������޷��Ű࣡"
        Exit Sub
    End If
    
    Dim ArrSJD() As String
    
    rs.Open "select * from ʱ��� order by ��� asc", Cnn
    i = -1
    If rs.EOF = False Then
        rs.MoveFirst
        Do While Not rs.EOF
            i = i + 1
            ReDim Preserve ArrSJD(i)
            ArrSJD(i) = Trim(rs.Fields("����"))
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim n As Long
    
    '��¼���ڲ�
    n = DateDiff("d", DTPicker1.Value, DTPicker2.Value)
    n = n - 1
    'ÿ��ÿ��ʱ��ζ�ÿ��ֵ���ҽ����Ű�
    
    Dim PBRYBH As String
    Dim PBRQ As Date
    
    For i = 0 To n          'ÿ��
        PBRQ = DateAdd("d", i, DTPicker1.Value)
        For k = 0 To UBound(ArrSJD)     'ÿ��ʱ���
            For m = 0 To UBound(ArrZBS)     'ÿ��ֵ����
                
                Do While 1
                    DoEvents
                    PBRYBH = ArrZBY(0)
                    '�鿴�Ƿ���ٵ�ǰʱ��ε�ǰ��Ա�Ƿ����
                    rs.Open "select * from ��ٱ� where �������=#" & PBRQ & "# and ʱ���='" & ArrSJD(k) & "' and ��Ա���='" & PBRYBH & "'", Cnn
                    If rs.EOF = True Then   '���û����٣�����ֵ��
                        rs.Close
                        Cnn.Execute "insert into ֵ���(����,ʱ���,��Ա���,ֵ���ұ��) values(#" & PBRQ & "#,'" & ArrSJD(k) & "','" & PBRYBH & "','" & ArrZBS(m) & "')"
                        QY ArrZBY
                        Exit Do
                    End If
                    rs.Close
                    QY ArrZBY
                Loop
            Next
        Next
    Next
    
    MsgBox "�Ű���ϣ�"
    
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
