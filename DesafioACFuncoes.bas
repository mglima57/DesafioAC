Attribute VB_Name = "DesafioACFuncoes"
Option Explicit

'CONSTANTES DAS APIs
Public Const SW_SHOW = 5
Const gstrNULL = ""
Const gintMAX_SIZE = 255

'DECLARACOES DAS APIs
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal LpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Function ReadIniFile(ByVal StrIniFile As String, ByVal strSection As String, ByVal strKey As String) As String
    Dim strBuffer As String
    Dim intPos As Integer
        
    strBuffer = Space$(gintMAX_SIZE)
    
    If GetPrivateProfileString(strSection, strKey, gstrNULL, strBuffer, gintMAX_SIZE, StrIniFile) > 0 Then
        ReadIniFile = RTrim(StripTerminator(strBuffer))
    Else
        ReadIniFile = gstrNULL
    End If
End Function

Sub WriteIniFile(ByVal StrIniFile As String, ByVal strSection As String, ByVal strKey As String, ByVal strConteudo As String)
    Dim strBuffer As String
    Dim intPos As Integer
    strBuffer = WritePrivateProfileString(strSection, strKey, strConteudo, StrIniFile)
End Sub

Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Function FormatCripto(ByVal StrCripto, ByVal BolAcao) As String
    'BolAcao = true = criptografa / false = descriptografa
    Dim X, i, N, TamChave, valor  'Declaração das variáveis.
    Dim aChaves(9) 'Declaração do Array que contem as chaves de criptografia
    aChaves(0) = 77 'Atribuição dos valores dos caracteres que serão utilizados
    aChaves(1) = 84 'como chave de criptografia
    aChaves(2) = 79
    aChaves(3) = 65
    aChaves(4) = 73
    aChaves(5) = 78
    aChaves(6) = 67
    aChaves(7) = 70
    aChaves(8) = 82
    X = "" 'Atribui o valor vazio para a variável X
    For i = 1 To Len(StrCripto) 'Inicia um Loop For que deve ser repetido a quantidade de vezes de acordo com o tamanho da String que deve ser criptografada.
        N = Asc(Mid(StrCripto, i, 1))
        If N > 31 Then 'mantém controles intactos
            N = N - 32 'desconsidera controles (Caracteres de 0 a 31)
            If BolAcao Then
                valor = 1
            Else
                valor = -1
            End If
            N = N + (aChaves((i - 1) Mod 9)) * valor
            N = N Mod 224 ' Caracteres 256 - 32 desconsiderados
            If N < 0 Then
                N = 224 + N
            End If
            N = N + 32 ' Reajusta para caracteres normais
        End If
        X = X & Chr(N) 'Armazena na variável X os caracteres que foram modificados
    Next
    FormatCripto = X 'Retorna o valor da variável X para a função
End Function

Function aspas(Texto As Variant) As String
    If Not IsNull(Texto) Then
        Texto = Replace(Texto, "'", "")
        aspas = Chr(39) & Trim(Texto) & Chr(39)
    Else
        aspas = "NULL"
    End If
End Function

Public Function TiraAcento(Text As Variant, Optional AdditionalChars As String) As String
    Dim NewText As String, letra   As String
    Dim l As Long, N As Long, i As Long
    If Text = "" Then
        Exit Function
    End If
    NewText = UCase(Text)
    l = Len(NewText)
    For N = 1 To l
        letra = Mid(NewText, N, 1)
        i = InStr("ÁÀÃÂÄÊËÉÍÏÓÔÕÖÚÜÛÇ'ºª" & AdditionalChars, letra)
        'i = InStr("'" & AdditionalChars, Letra)
        If i <> 0 Then
            TiraAcento = TiraAcento + Mid("AAAAAEEEIIOOOOUUUC ", i, 1)
            'TiraAcento = TiraAcento & " "
        Else
            TiraAcento = TiraAcento & letra
        End If
    Next
End Function

Function SQLDecimal(varString As Variant) As Variant
    varString = IIf(IsNull(varString), "", varString)
    SQLDecimal = Replace(MyVal(CStr(varString)), ",", ".")
End Function

Public Function ValidaCNPJ_CPF(ByVal pstrCNPJ_CPF As String) As Boolean
    Dim Soma As Long
    Dim ind As Integer
    Dim pos As Integer
    Dim resto As Long
    Dim Digito As Integer
    Dim DV As String
    Dim X As Integer

    If IsNull(pstrCNPJ_CPF) Then
        ValidaCNPJ_CPF = False
        Exit Function
    End If

    pstrCNPJ_CPF = RetiraChar(pstrCNPJ_CPF, ".")
    pstrCNPJ_CPF = RetiraChar(pstrCNPJ_CPF, "/")
    pstrCNPJ_CPF = RetiraChar(pstrCNPJ_CPF, "-")

    If Not IsNumeric(pstrCNPJ_CPF) Then
        ValidaCNPJ_CPF = False
        Exit Function
    End If

    'se for diferente de 11 ou 14 digitos esta errado
    If Len(Trim(pstrCNPJ_CPF)) <> 11 And Len(Trim(pstrCNPJ_CPF)) <> 14 Then
        ValidaCNPJ_CPF = False
        Exit Function
    End If
    
    Digito = Right(pstrCNPJ_CPF, 2)
    DV = ""
    pstrCNPJ_CPF = Left(pstrCNPJ_CPF, Len(pstrCNPJ_CPF) - 2)

    For X = 1 To 2
        Soma = 0
        ind = 2
        For pos = Len(pstrCNPJ_CPF) To 1 Step -1
            Soma = Soma + (CLng(Mid(pstrCNPJ_CPF, pos, 1)) * ind)
            ind = ind + 1

            If Len(pstrCNPJ_CPF) > 11 And ind > 9 Then
                ind = 2
            End If
        Next

        resto = Soma - ((Soma \ 11) * 11)
 
        If resto < 2 Then
            pstrCNPJ_CPF = pstrCNPJ_CPF & "0"
            DV = DV & "0"
        Else
            pstrCNPJ_CPF = pstrCNPJ_CPF & CStr(11 - resto)
            DV = DV & CStr(11 - resto)
        End If
    Next
    If DV = Digito Then
        ValidaCNPJ_CPF = True
    Else
        ValidaCNPJ_CPF = False
    End If
    
End Function

Public Function RetiraChar(ByVal pstrTexto As String, ByVal pstrChar As String) As String
    Do While InStr(pstrTexto, pstrChar) > 0
        pstrTexto = Left(pstrTexto, InStr(pstrTexto, pstrChar) - 1) & Mid(pstrTexto, InStr(pstrTexto, pstrChar) + 1, Len(pstrTexto))
    Loop
    RetiraChar = pstrTexto
End Function

Public Function MyVal(ByVal valor As Variant) As Double
    valor = IIf(IsNull(valor), "", valor)
    MyVal = Val(Replace(CStr(valor), ",", "."))
End Function

Public Function GravaLogOperacao(Tela As String, Usuario As String, operacao As String, ByRef DB As ADODB.Connection)
    
    On Error Resume Next
    
    Dim SQL As String
    
    'coloca aspas duplas por causa do SQL
    operacao = Replace(operacao, "'", "''")
    
    'grava a operacao no log
    SQL = "insert LogOperacao ("
    SQL = SQL & "DataOperacao,"
    SQL = SQL & "Tela,"
    SQL = SQL & "Usuario,"
    SQL = SQL & "Operacao"
    SQL = SQL & ") values ("
    SQL = SQL & "getdate(),"
    SQL = SQL & aspas(Tela) & ","
    SQL = SQL & aspas(Usuario) & ","
    SQL = SQL & "'" & UCase(operacao) & "'" & ")"
    
    'grava no banco
    DB.Execute SQL, , adExecuteNoRecords

End Function

