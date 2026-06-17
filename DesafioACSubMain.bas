Attribute VB_Name = "SubMain"
Sub Main()
    
    'pega a estação corrente para evitar travar em copia de bases entre maquinas
    Dim tam As Long
    tam = 255
    WorkStation2 = String(255, 0)
    GetComputerName WorkStation2, tam
    WorkStation2 = Left(WorkStation2, tam)
    
    'Senha Admin
    SenhaAdmin = "admin@"
    
    'Ler do Arquivo myGasIni os parametros para Inicialização do banco de dados
    IniFile = App.Path & "\DesafioAC.ini"
    If Len(Dir(IniFile)) = 0 Then
        MsgBox "Arquivo de Inicialização 'DesafioAC.INI' não encontrado." & vbCrLf & vbCrLf & _
        "Favor informar as configurações Inicias para prosseguir.", vbCritical, "Inicialização."
        
        'força entrar para editar o DesafioAC.INI
        frm90.Show 1

        MsgBox "Favor inicialize novamente o sistema.", vbCritical, "Inicialização."
        
        Exit Sub
        
    End If

    'recupera variaveis de ambiente
        'pegar o nome do servidor
        Servidor = FormatCripto(Trim(ReadIniFile(IniFile, "IDENTIFICACAO", "ServerSQL")), False)
            
        'pegar a porta
        PortaIP = FormatCripto(Trim(ReadIniFile(IniFile, "IDENTIFICACAO", "PortaIP")), False)
        
        'pegar o nome da base
        Database = FormatCripto(Trim(ReadIniFile(IniFile, "IDENTIFICACAO", "Database")), False)
       
        'pegar o nome do usuario
        Usuario = FormatCripto(Trim(ReadIniFile(IniFile, "IDENTIFICACAO", "Usuario")), False)
       
        'pegar a senha
        Senha = FormatCripto(Trim(ReadIniFile(IniFile, "IDENTIFICACAO", "Senha")), False)
        
        
    '*** ESTA DEFINIDO ABRE O BANCO DE DADOS abre conexão com o banco de dados
    If Not Conexao.OpenDB() Then
        GoTo Sair
    End If
    
    'modal Menu
    frmLogin.Show
    
    Exit Sub
    
Sair:
    End

End Sub

