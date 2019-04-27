Attribute VB_Name = "Module1"
Public Cnn As New ADODB.Connection   '声明数据库连接对象

Public strXM As String
Public strMM As String
Public strQX As String



Public Sub Main()
    On Error GoTo ErrH
        
        Cnn.CursorLocation = adUseClient

        Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\mydb.mdb"   '连接数据库
        

    FrmLogin.Show     '显示登陆窗体
    Exit Sub
ErrH:
    MsgBox Err.Description
End Sub






