'Essa macro usa a biblioteca ADODB (ActiveX Data Objects DataBase) para se conectar com o banco de dados SQL e selecionar dados de uma tabela específica. 
'A string de conexão é definida usando as informações do seu banco de dados, como o nome do servidor, do banco de dados, do usuário e da senha. Depois, 
'é criada uma string SQL para selecionar os dados da tabela. Em seguida, o objeto de registro é preenchido com os dados selecionados. Por fim, os dados 
'são carregados nos campos do formulário e a conexão e o objeto de registro são fechados.

'Tenha em mente que essa é uma implementação básica, é importante tratar os erros e adicionar validação dos dados, além disso, é recomendado usar 
'parametrização de query para evitar possíveis vulnerabilidades de SQL injection.

'by Daniel Almeida - dcalmeida@ibm.com

Sub CarregarFormSQL()
    Dim Cn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim strConexao As String, strSQL As String

    'Definir a string de conexão com o banco de dados
    strConexao = "Provider=SQLOLEDB;Data Source=NOME_SERVIDOR;Initial Catalog=NOME_BANCO;User ID=USUARIO;Password=SENHA;"

    'Criar um objeto de conexão
    Set Cn = New ADODB.Connection
    Cn.Open strConexao

    'Definir a string SQL para selecionar dados
    strSQL = "SELECT COLUNA1, COLUNA2, COLUNA3 FROM NOME_TABELA"

    'Criar um objeto de registro
    Set Rs = New ADODB.Recordset
    Rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly

    'Carregar os dados no formulário
    Me.TextBox1.Value = Rs.Fields("COLUNA1").Value
    Me.TextBox2.Value = Rs.Fields("COLUNA2").Value
    Me.TextBox3.Value = Rs.Fields("COLUNA3").Value

    'Fechar o objeto de registro e conexão
    Rs.Close
    Cn.Close
    Set Rs = Nothing
    Set Cn = Nothing
End Sub


'EXEMPLO 2 DE COMO CARREGAR 
Sub CarregarListBoxSQL()
    Dim Cn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim strConexao As String, strSQL As String
    Dim i As Integer

    'Definir a string de conexão com o banco de dados
    strConexao = "Provider=SQLOLEDB;Data Source=NOME_SERVIDOR;Initial Catalog=NOME_BANCO;User ID=USUARIO;Password=SENHA;"

    'Criar um objeto de conexão
    Set Cn = New ADODB.Connection
    Cn.Open strConexao

    'Definir a string SQL para selecionar dados
    strSQL = "SELECT COLUNA1, COLUNA2 FROM NOME_TABELA"

    'Criar um objeto de registro
    Set Rs = New ADODB.Recordset
    Rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly

    'Limpar o ListBox
    Me.ListBox1.Clear

    'Adicionar o nome das colunas no ListBox
    For i = 0 To Rs.Fields.Count - 1
        Me.ListBox1.AddItem Rs.Fields(i).Name
    Next i

    'Adicionar os dados do Recordset no ListBox
    Rs.MoveFirst
    Do While Not Rs.EOF
        Me.ListBox1.AddItem Rs.Fields("COLUNA1").Value & " - " & Rs.Fields("COLUNA2").Value
        Rs.MoveNext
    Loop

    'Fechar o objeto de registro e conexão
    Rs.Close
    Cn.Close
    Set Rs = Nothing
    Set Cn = Nothing
End Sub
