Attribute VB_Name = "Conexao"
Option Explicit

'propriedades publicas do modulo
Global DBTeste As New ADODB.Connection
Global DB As New ADODB.Connection
Global DB2 As New ADODB.Connection

Public Function OpenDB() As Boolean

    'tratamento de erro
    On Error Resume Next
    
    'pessimista
    OpenDB = False

    'se nao esta aberto entao abre
    If DB.State <> adStateOpen Then
        DB.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;LANGUAGE=us_english;"
        DB.ConnectionString = DB.ConnectionString & "User ID=" & Trim(Usuario) & ";"
        DB.ConnectionString = DB.ConnectionString & "Password=" & Trim(Senha) & ";"
        DB.ConnectionString = DB.ConnectionString & "Initial Catalog=" & Trim(Database) & ";"
        DB.ConnectionString = DB.ConnectionString & "Data Source=" & Trim(Servidor) & "," & Trim(PortaIP) & ";"
        DB.ConnectionString = DB.ConnectionString & "Connect Timeout=2"
        DB.Open
        If DB.State <> adStateOpen Then
            MsgBox "Não foi possível estabelecer conexão DataBase!", vbCritical, "MENSAGEM"
            Exit Function
        End If
    
    End If
    
    'chegou até aqui - OK
    OpenDB = True

End Function

Public Function CloseDB() As Boolean

    'tratamento de erro
    On Error Resume Next
    
    'pessimista
    CloseDB = False

    'se nao esta fechado entao fecha
    If DB.State = adStateOpen Then
        DB.Close
    End If
    
    'chegou até aqui - OK
    CloseDB = True
    
End Function

