VERSION 5.00
Begin VB.Form FrmLogin 
   Caption         =   "��¼"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   13545
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��½"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "�Ű�ϵͳ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�˺�"
      BeginProperty Font 
         Name            =   "����"
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
    Dim rs As New ADODB.Recordset   '������¼��
    rs.Open "select * from �û� where �˺�='" & Trim(Text1.Text) & "' and ����='" & Trim(Text2.Text) & "'", Cnn   '�����û���
    If rs.EOF = True Then    '����鲻������ʾ����
        MsgBox "�����������޴��û�", vbCritical
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
     '�ҵ��󣬼�¼�˺������Ȩ��
    strXM = Trim(rs.Fields(0).Value)
    strMM = Trim(rs.Fields(1).Value)
    strQX = Trim(rs.Fields(2).Value)
    rs.Close  '�رռ�¼ �ͷ���Դ
    Set rs = Nothing
    
    FrmMain.Show   '��ʾ������
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


