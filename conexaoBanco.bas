Attribute VB_Name = "conexaoBanco"
Public cn As ADODB.Connection

Public Sub realizarConexao()
    Set cn = New ADODB.Connection
    cn.Open "Provider=SQLOLEDB; Initial Catalog=testeCrud; Data Source=127.0.0.1; Integrated Security=SSPI; Persist Security Info=True"
End Sub

