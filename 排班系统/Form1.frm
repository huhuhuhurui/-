VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "排班系统"
   ClientHeight    =   11745
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18405
   LinkTopic       =   "Form1"
   ScaleHeight     =   11745
   ScaleWidth      =   18405
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Menu Mnu_XTGL 
      Caption         =   "系统管理"
      Begin VB.Menu Mnu_YHGL 
         Caption         =   "用户管理"
      End
      Begin VB.Menu Mnu_MMXG 
         Caption         =   "密码修改"
      End
   End
   Begin VB.Menu Mnu_SJSZ 
      Caption         =   "数据设置"
      Begin VB.Menu Mnu_BZYSZ 
         Caption         =   "值班员设置"
      End
      Begin VB.Menu Mnu_ZBSSZ 
         Caption         =   "值班室设置"
      End
      Begin VB.Menu Mnu_SJDSZ 
         Caption         =   "时间段设置"
      End
      Begin VB.Menu Mnu_QJBGL 
         Caption         =   "请假表管理"
      End
   End
   Begin VB.Menu Mnu_PBXX 
      Caption         =   "排班信息"
      Begin VB.Menu Mnu_ZDPB 
         Caption         =   "自动排班"
      End
      Begin VB.Menu Mnu_SDTB 
         Caption         =   "手动调班"
      End
      Begin VB.Menu Mnu_PBBCK 
         Caption         =   "排班表查看"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If strQX <> "管理员" Then
        
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
