Attribute VB_Name = "Module1"
Public Cnn As New ADODB.Connection   '�������ݿ����Ӷ���

Public strXM As String
Public strMM As String
Public strQX As String



Public Sub Main()
    On Error GoTo ErrH
        
        Cnn.CursorLocation = adUseClient

        Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\mydb.mdb"   '�������ݿ�
        

    FrmLogin.Show     '��ʾ��½����
    Exit Sub
ErrH:
    MsgBox Err.Description
End Sub






