'#############################
'essa macro verifica se o registro é duplicado antes de inserir no banco de dados Access.
'#############################
'Neste exemplo, o código usa um objeto Recordset para verificar se o registro já existe na tabela,
'comparando os valores dos campos Campo1 e Campo2. Se o registro já existir, uma mensagem é exibida
'e o registro não é inserido. Caso contrário, o registro é inserido usando o comando SQL INSERT.
'É importante observar que essa é uma forma insegura de se conectar ao banco de dados, pois não protege
'contra ataques de injeção SQL, é recomendado o uso de Prepared Statement.

Sub InserirRegistro()
    Dim Cn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim strConexao As String, strSQL As String
    Dim novoRegistro As Boolean

    'Definir a string de conexão com o banco de dados
    strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BancoDeDados.accdb;"

    'Criar um objeto de conexão
    Set Cn = New ADODB.Connection
    Cn.Open strConexao

    'Criar um objeto de registro para verificar se o registro já existe
    Set Rs = New ADODB.Recordset
    Rs.Open "SELECT * FROM Tabela WHERE Campo1 = '" & Me.TextBox1.Value & "' AND Campo2 = '" & Me.TextBox2.Value & "'", Cn, adOpenStatic, adLockReadOnly

    'Verificar se o registro já existe
    If Rs.EOF Then
        novoRegistro = True
    Else
        novoRegistro = False
    End If

    'Fechar o objeto de registro
    Rs.Close
    Set Rs = Nothing

    'Inserir o registro no banco de dados, se ele ainda não existir
    If novoRegistro = True Then
        strSQL = "INSERT INTO Tabela (Campo1, Campo2, Campo3) VALUES ('" & Me.TextBox1.Value & "', '" & Me.TextBox2.Value & "', '" & Me.TextBox3.Value & "')"
        Cn.Execute strSQL
        MsgBox "Registro inserido com sucesso!", vbInformation, "Inserir Registro"
    Else
        MsgBox "Registro já existe!", vbExclamation, "Inserir Registro"
    End If

    'Fechar o objeto de conexão
    Cn.Close
    Set Cn = Nothing
End Sub
