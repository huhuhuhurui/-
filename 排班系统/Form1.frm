VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "�Ű�ϵͳ"
   ClientHeight    =   11745
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18405
   LinkTopic       =   "Form1"
   ScaleHeight     =   11745
   ScaleWidth      =   18405
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.Menu Mnu_XTGL 
      Caption         =   "ϵͳ����"
      Begin VB.Menu Mnu_YHGL 
         Caption         =   "�û�����"
      End
      Begin VB.Menu Mnu_MMXG 
         Caption         =   "�����޸�"
      End
   End
   Begin VB.Menu Mnu_SJSZ 
      Caption         =   "��������"
      Begin VB.Menu Mnu_BZYSZ 
         Caption         =   "ֵ��Ա����"
      End
      Begin VB.Menu Mnu_ZBSSZ 
         Caption         =   "ֵ��������"
      End
      Begin VB.Menu Mnu_SJDSZ 
         Caption         =   "ʱ�������"
      End
      Begin VB.Menu Mnu_QJBGL 
         Caption         =   "��ٱ����"
      End
   End
   Begin VB.Menu Mnu_PBXX 
      Caption         =   "�Ű���Ϣ"
      Begin VB.Menu Mnu_ZDPB 
         Caption         =   "�Զ��Ű�"
      End
      Begin VB.Menu Mnu_SDTB 
         Caption         =   "�ֶ�����"
      End
      Begin VB.Menu Mnu_PBBCK 
         Caption         =   "�Ű��鿴"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If strQX <> "����Ա" Then
        
        Mnu_YHGL.Visible = False
        Mnu_SJSZ.Visible = False
        Mnu_ZDPB.Visible = False
        Mnu_SDTB.Visible = False
    End If
End Sub

Private Sub Mnu_BZYSZ_Click()
FrmZBYSZ.Show
End Sub

Private Sub Mnu_MMXG_Click()
FrmMMXG.Show
End Sub

Private Sub Mnu_PBBCK_Click()
FrmPBB.Show
End Sub

Private Sub Mnu_QJBGL_Click()
FrmQJB.Show
End Sub

Private Sub Mnu_SDTB_Click()
FrmSDTZ.Show
End Sub

Private Sub Mnu_SJDSZ_Click()
FrmSJDSZ.Show
End Sub

Private Sub Mnu_YHGL_Click()
FrmYHGL.Show
End Sub

Private Sub Mnu_ZBSSZ_Click()
FrmZBSSZ.Show
End Sub

Private Sub Mnu_ZDPB_Click()
FrmZDPB.Show
End Sub
