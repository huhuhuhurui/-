VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPBB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Ű��鿴"
   ClientHeight    =   12615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20955
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
   ScaleHeight     =   12615
   ScaleWidth      =   20955
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   615
      Left            =   12000
      TabIndex        =   9
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   615
      Left            =   14520
      TabIndex        =   7
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�鿴"
      Height          =   615
      Left            =   9480
      TabIndex        =   6
      Top             =   360
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   87359489
      CurrentDate     =   43579
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   87359489
      CurrentDate     =   43579
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   10575
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   20775
      _ExtentX        =   36645
      _ExtentY        =   18653
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ֵ����"
      Height          =   210
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��ֹ����"
      Height          =   210
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   210
      Left            =   5640
      TabIndex        =   2
      Top             =   960
      Width           =   210
   End
End
Attribute VB_Name = "FrmPBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim rs As New ADODB.Recordset
    Dim lngCount As Long
    Dim i As Long
    Dim j As Long    '������Ǳ���
    Dim m As Long    '�������ʱ�����
    Dim k As Long    '������������к�
    Dim x As Long    'ѭ��ֵ����
    Dim n As Long    '��ǿ�ʼ��
    Dim hs As Integer
    If Combo1.Text <> "" Then  '��ѯ����ֵ���ҵ�
    
        rs.Open "select a.����,a.ʱ���,a.ֵ���ұ��,c.���� as ֵ��������,b.���,b.���� from ֵ��� a,ֵ��Ա b,ֵ���� c where a.��Ա���=b.��� and a.ֵ���ұ��=c.��� and a.����>=#" & DTPicker1.Value & "# and a.����<#" & DTPicker2.Value & "# and a.ֵ���ұ��='" & Trim(Combo1.Text) & "' order by a.ֵ���ұ�� asc,a.���� asc,a.ʱ��� asc", Cnn
        
        If rs.EOF = False Then
            rs.MoveLast
            rs.MoveFirst
            lngCount = rs.RecordCount
            MSH.Cols = 1
            MSH.ColWidth(0) = 2000
            For i = 1 To rs.RecordCount
                If MSH.Cols = 1 Then
                    MSH.Cols = MSH.Cols + 1
                    MSH.ColWidth(MSH.Cols - 1) = 2000
                    MSH.Rows = 2
                    j = 0
                    m = 0
                    k = 1
                    MSH.TextMatrix(m, j) = Trim(rs.Fields("ֵ��������").Value)
                    MSH.TextMatrix(m, k) = Trim(rs.Fields("����").Value)
                    m = m + 1
                    MSH.TextMatrix(m, 0) = Trim(rs.Fields("ʱ���").Value)
                    MSH.TextMatrix(m, k) = Trim(rs.Fields("���").Value) & "-" & Trim(rs.Fields("����").Value)
                Else
                    If Trim(rs.Fields("����")) <> Trim(MSH.TextMatrix(j, MSH.Cols - 1)) Then  '������ڲ�ͬ˵����ʼһ��������
                        MSH.Cols = MSH.Cols + 1
                        MSH.ColWidth(MSH.Cols - 1) = 2000
                        k = k + 1
                        m = 0
                        MSH.TextMatrix(m, k) = Trim(rs.Fields("����").Value)
                        m = m + 1
                        MSH.TextMatrix(m, 0) = Trim(rs.Fields("ʱ���").Value)
                        MSH.TextMatrix(m, k) = Trim(rs.Fields("���").Value) & "-" & Trim(rs.Fields("����").Value)
                    Else
                        m = m + 1
                        If MSH.Rows = m Then
                            MSH.Rows = MSH.Rows + 1
                        End If
                        MSH.TextMatrix(m, 0) = Trim(rs.Fields("ʱ���").Value)
                        MSH.TextMatrix(m, k) = Trim(rs.Fields("���").Value) & "-" & Trim(rs.Fields("����").Value)
                        
                    End If
                End If
                

                rs.MoveNext
            Next
        
        End If
    
        rs.Close
    Else

        For x = 0 To Combo1.ListCount - 1
            
            If x = 0 Then
                rs.Open "select a.����,a.ʱ���,a.ֵ���ұ��,c.���� as ֵ��������,b.���,b.���� from ֵ��� a,ֵ��Ա b,ֵ���� c where a.��Ա���=b.��� and a.ֵ���ұ��=c.��� and a.����>=#" & DTPicker1.Value & "# and a.����<#" & DTPicker2.Value & "# and a.ֵ���ұ��='" & Trim(Combo1.List(x)) & "' order by a.ֵ���ұ�� asc,a.���� asc,a.ʱ��� asc", Cnn
                
                If rs.EOF = False Then
                    rs.MoveLast
                    rs.MoveFirst
                    lngCount = rs.RecordCount
                    MSH.Cols = 1
                    MSH.ColWidth(0) = 2000
                    For i = 1 To rs.RecordCount
                        If MSH.Cols = 1 Then
                            MSH.Cols = MSH.Cols + 1
                            MSH.ColWidth(MSH.Cols - 1) = 2000
                            MSH.Rows = 2
                            j = 0
                            m = 0
                            k = 1
                            MSH.TextMatrix(m, j) = Trim(rs.Fields("ֵ��������").Value)
                            MSH.TextMatrix(m, k) = Trim(rs.Fields("����").Value)
                            m = m + 1
                            MSH.TextMatrix(m, 0) = Trim(rs.Fields("ʱ���").Value)
                            MSH.TextMatrix(m, k) = Trim(rs.Fields("���").Value) & "-" & Trim(rs.Fields("����").Value)
                        Else
                            If Trim(rs.Fields("����")) <> Trim(MSH.TextMatrix(j, MSH.Cols - 1)) Then  '������ڲ�ͬ˵����ʼһ��������
                                MSH.Cols = MSH.Cols + 1
                                MSH.ColWidth(MSH.Cols - 1) = 2000
                                k = k + 1
                                m = 0
                                MSH.TextMatrix(m, k) = Trim(rs.Fields("����").Value)
                                m = m + 1
                                MSH.TextMatrix(m, 0) = Trim(rs.Fields("ʱ���").Value)
                                MSH.TextMatrix(m, k) = Trim(rs.Fields("���").Value) & "-" & Trim(rs.Fields("����").Value)
                            Else
                                m = m + 1
                                If MSH.Rows = m Then
                                    MSH.Rows = MSH.Rows + 1
                                End If
                                MSH.TextMatrix(m, 0) = Trim(rs.Fields("ʱ���").Value)
                                MSH.TextMatrix(m, k) = Trim(rs.Fields("���").Value) & "-" & Trim(rs.Fields("����").Value)
                                
                            End If
                        End If
                        
        
                        rs.MoveNext
                    Next
                
                End If
                
                rs.Close
                hs = MSH.Rows
            Else
                
                rs.Open "select a.����,a.ʱ���,a.ֵ���ұ��,c.���� as ֵ��������,b.���,b.���� from ֵ��� a,ֵ��Ա b,ֵ���� c where a.��Ա���=b.��� and a.ֵ���ұ��=c.��� and a.����>=#" & DTPicker1.Value & "# and a.����<#" & DTPicker2.Value & "# and a.ֵ���ұ��='" & Trim(Combo1.List(x)) & "' order by a.ֵ���ұ�� asc,a.���� asc,a.ʱ��� asc", Cnn
                
                If rs.EOF = False Then
                    rs.MoveLast
                    rs.MoveFirst
                    lngCount = rs.RecordCount
                    n = MSH.Rows
                    MSH.Rows = MSH.Rows + hs
                    m = n
                    j = 0
                    k = 1
                    MSH.TextMatrix(m, j) = Trim(rs.Fields("ֵ��������").Value)
                    MSH.TextMatrix(m, k) = Trim(rs.Fields("����").Value)
                    m = m + 1
'                    MSH.Rows = MSH.Rows + 1
                    MSH.TextMatrix(m, 0) = Trim(rs.Fields("ʱ���").Value)
                    MSH.TextMatrix(m, k) = Trim(rs.Fields("���").Value) & "-" & Trim(rs.Fields("����").Value)
'                    MSH.Cols = 1
'                    MSH.ColWidth(0) = 2000
                    rs.MoveNext
                    For i = 1 To rs.RecordCount
'                        If MSH.Cols = 1 Then
'                            MSH.Cols = MSH.Cols + 1
'                            MSH.ColWidth(MSH.Cols - 1) = 2000
'                            MSH.Rows = 2
'                            j = 0
'                            m = 0
'                            k = 1
'                            MSH.TextMatrix(m, j) = Trim(rs.Fields("ֵ��������").Value)
'                            MSH.TextMatrix(m, k) = Trim(rs.Fields("����").Value)
'                            m = m + 1
'                            MSH.TextMatrix(m, 0) = Trim(rs.Fields("ʱ���").Value)
'                            MSH.TextMatrix(m, k) = Trim(rs.Fields("���").Value) & "-" & Trim(rs.Fields("����").Value)
'                        Else
                            If Trim(rs.Fields("����")) <> Trim(MSH.TextMatrix(n, k)) Then '������ڲ�ͬ˵����ʼһ��������
'                                MSH.Cols = MSH.Cols + 1
'                                MSH.ColWidth(MSH.Cols - 1) = 2000
                                k = k + 1
                                m = n
                                MSH.TextMatrix(m, k) = Trim(rs.Fields("����").Value)
                                m = m + 1
                                MSH.TextMatrix(m, 0) = Trim(rs.Fields("ʱ���").Value)
                                MSH.TextMatrix(m, k) = Trim(rs.Fields("���").Value) & "-" & Trim(rs.Fields("����").Value)
                            Else
                                m = m + 1
                                If MSH.Rows = m Then
                                    MSH.Rows = MSH.Rows + 1
                                End If
                                MSH.TextMatrix(m, 0) = Trim(rs.Fields("ʱ���").Value)
                                MSH.TextMatrix(m, k) = Trim(rs.Fields("���").Value) & "-" & Trim(rs.Fields("����").Value)
                                
                            End If
'                        End If
                        
        
                        rs.MoveNext
                        If rs.EOF = True Then Exit For
                    Next
                
                End If
            
            
            
                rs.Close
            End If
        
            
        Next

    End If
 
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
        Dim ExcelApp As Object
        Dim ExcelWorkBook As Object
        Dim ExcelWorkSheet As Object
On Error GoTo ErrH
        Set ExcelApp = CreateObject("Excel.Application")
        Set ExcelWorkBook = ExcelApp.Workbooks.Add
        Set ExcelWorkSheet = ExcelWorkBook.Worksheets(1)
        
        ExcelApp.Visible = True
        
        For i = 0 To MSH.Rows - 1
        
            For j = 0 To MSH.Cols - 1
        
                ExcelWorkSheet.Cells(i + 1, j + 1) = MSH.TextMatrix(i, j)
            
            Next
        Next
        
        

    Exit Sub
ErrH:
    
    ExcelApp.Quit
    Set ExcelApp = Nothing
    Set ExcelWorkBook = Nothing
    Set ExcelWorkSheet = Nothing
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    Dim rs As New ADODB.Recordset
    rs.Open "select * from ֵ����", Cnn
    
    If rs.EOF = False Then
        rs.MoveFirst
        Do While Not rs.EOF
            Combo1.AddItem Trim(rs.Fields(0))
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    
    
    
End Sub
