VERSION 5.00
Begin VB.Form FrmMMXG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����޸�"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10935
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   5775
      Begin VB.TextBox txtnewpwd2 
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
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtnewpwd1 
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
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1500
         Width           =   2175
      End
      Begin VB.TextBox txtoldpwd 
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
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ظ�������"
         BeginProperty Font 
            Name            =   "����"
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
         Left            =   720
         TabIndex        =   6
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
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
         Left            =   1080
         TabIndex        =   5
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ����"
         BeginProperty Font 
            Name            =   "����"
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
         Left            =   1080
         TabIndex        =   4
         Top             =   600
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmMMXG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Trim(txtoldpwd.Text) = strMM And txtnewpwd1.Text = txtnewpwd2.Text Then
        Cnn.Execute "update �û� set ����='" & Trim(txtnewpwd1.Text) & "' where �˺� ='" & strXM & "'"
        strMM = Trim(txtnewpwd1.Text)
        MsgBox "�޸ĳɹ�"
    ElseIf Trim(txtoldpwd.Text) = strMM And txtnewpwd1.Text <> txtnewpwd2.Text Then
        MsgBox "��������������벻һ����"
    Else
        MsgBox "�������������"
    End If
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


