VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmQJB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ٱ�"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15270
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
   ScaleHeight     =   10635
   ScaleWidth      =   15270
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   495
      Left            =   5640
      TabIndex        =   12
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ɾ��"
      Height          =   495
      Left            =   5640
      TabIndex        =   11
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
   Begin VB.ListBox List2 
      Height          =   3420
      Left            =   9480
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   960
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   120127489
      CurrentDate     =   43579
   End
   Begin VB.ListBox List1 
      Height          =   9930
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   5655
      Left            =   5160
      TabIndex        =   13
      Top             =   4800
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9975
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "��Ա����"
      Height          =   210
      Left            =   5160
      TabIndex        =   9
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��Ա���"
      Height          =   210
      Left            =   5160
      TabIndex        =   7
      Top             =   480
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "���ʱ���"
      Height          =   210
      Left            =   9480
      TabIndex        =   5
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   210
      Left            =   5160
      TabIndex        =   3
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ֵ����Ա�б�"
      Height          =   210
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1260
   End
End
Attribute VB_Name = "FrmQJB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim rs As New ADODB.Recordset
    
    For i = 0 To List2.ListCount - 1
        
        If List2.Selected(i) = True Then
            rs.Open "select * from ��ٱ� where �������=#" & DTPicker1.Value & "# and ʱ���='" & Trim(List2.List(i)) & "' and ��Ա���='" & Trim(Text1.Text) & "'", Cnn, adOpenDynamic, adLockOptimistic
            If rs.EOF = True Then
                rs.AddNew
                    rs.Fields("�������") = DTPicker1.Value
                    rs.Fields("ʱ���") = Trim(List2.List(i))
                    rs.Fields("��Ա���") = Trim(Text1.Text)
                rs.Update
            End If
            rs.Close
        End If
        
    Next
    
    Set rs = Nothing
    
    ShowMSH
    
End Sub

Private Sub Command3_Click()
    If MSH.RowSel > 0 Then
    
        Cnn.Execute "delete from ��ٱ� where ��¼��=" & MSH.TextMatrix(MSH.RowSel, 0) & ""
        ShowMSH
    
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Sub ShowMSH()

    Dim rs As New ADODB.Recordset
    rs.Open "select a.��¼��,a.�������,a.ʱ���,b.���,b.���� from ��ٱ� a,ֵ��Ա b where a.��Ա���=b.��� order by ��¼�� desc", Cnn
    
    Set MSH.DataSource = rs
    rs.Close
    Set rs = Nothing
    
    MSH.ColWidth(0) = 0
    MSH.ColWidth(1) = 2000
    MSH.ColWidth(2) = 2000
    MSH.ColWidth(3) = 2000
    MSH.ColWidth(4) = 2000
    
    
End Sub
Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    rs.Open "select * from ֵ��Ա order by ��� asc", Cnn
    If rs.EOF = False Then
    
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            List1.AddItem Trim(rs.Fields("���")) & "-" & Trim(rs.Fields("����"))
            rs.MoveNext
            
        Loop
        
    End If
    rs.Close
    
    rs.Open "select * from ʱ��� order by ��� asc", Cnn
    If rs.EOF = False Then
    
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            List2.AddItem Trim(rs.Fields("����"))
            rs.MoveNext
            
        Loop
        
    End If
    rs.Close
    
    
    Set rs = Nothing
    
    ShowMSH
    
End Sub

Private Sub List1_Click()
    
    For i = 0 To List1.ListCount - 1
        
        If List1.Selected(i) = True Then
            
            Text1.Text = Mid(List1.List(i), 1, InStr(List1.List(i), "-") - 1)
            Text2.Text = Mid(List1.List(i), InStr(List1.List(i), "-") + 1)
            
            Dim rs As New ADODB.Recordset
            rs.Open "select a.��¼��,a.�������,a.ʱ���,b.���,b.���� from ��ٱ� a,ֵ��Ա b where a.��Ա���=b.��� and ���='" & Text1.Text & "' order by ��¼�� desc", Cnn
            
            Set MSH.DataSource = rs
            rs.Close
            Set rs = Nothing
            MSH.ColWidth(0) = 0
            MSH.ColWidth(1) = 2000
            MSH.ColWidth(2) = 2000
            MSH.ColWidth(3) = 2000
            MSH.ColWidth(4) = 2000
            
            Exit For
            
        End If
    
    Next

End Sub

