VERSION 5.00
Begin VB.Form FrmRYXZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Աѡ��"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11130
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
   ScaleHeight     =   7755
   ScaleWidth      =   11130
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   3240
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   5730
      Left            =   2280
      TabIndex        =   0
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��Ա���"
      Height          =   210
      Left            =   6600
      TabIndex        =   6
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "��Ա����"
      Height          =   210
      Left            =   6600
      TabIndex        =   5
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ѡ��ֵ����Ա�б�"
      Height          =   210
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   1680
   End
End
Attribute VB_Name = "FrmRYXZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelRQ As Date
Public SelSJD As String
'Public SelZBS As String
Public SelZBY As String

Private Sub Command1_Click()
    If Text1.Text = "" Then
    
        MsgBox "��ѡ���滻��Ա��"
        Exit Sub
        
    End If
    
    Dim rs As New ADODB.Recordset
    rs.Open "select * from ֵ��� where ����=#" & SelRQ & "# and ʱ���='" & SelSJD & "' and ��Ա���='" & Trim(Text1.Text) & "'", Cnn, adOpenDynamic, adLockOptimistic
    
    If rs.EOF = True Then
        rs.Close
        Cnn.Execute "update ֵ��� set ��Ա���='" & Trim(Text1.Text) & "' where ����=#" & SelRQ & "# and ʱ���='" & SelSJD & "' and ��Ա���='" & SelZBY & "'"
        MsgBox "�����ɹ���"
    Else
        rs.Close
        MsgBox "��ֵ��Ա�Ѿ������ʱ�䰲��ֵ�࣡"
    End If
    Set rs = Nothing
    
End Sub

Private Sub Command2_Click()
    Unload Me
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
    Set rs = Nothing
    
End Sub

Private Sub List1_Click()
    For i = 0 To List1.ListCount - 1
        
        If List1.Selected(i) = True Then
            
            Text1.Text = Mid(List1.List(i), 1, InStr(List1.List(i), "-") - 1)
            Text2.Text = Mid(List1.List(i), InStr(List1.List(i), "-") + 1)
            
            Exit For
            
        End If
    
    Next
End Sub
