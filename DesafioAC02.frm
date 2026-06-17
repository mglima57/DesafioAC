VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.ocx"
Begin VB.Form frm02 
   Caption         =   "02 - Usuarios/Permissões"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame frmDados 
      Height          =   1125
      Left            =   0
      TabIndex        =   0
      Top             =   570
      Width           =   10425
      _Version        =   65536
      _ExtentX        =   18389
      _ExtentY        =   1984
      _StockProps     =   14
      Caption         =   "Dados do Usuário"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtSenhaUsuario 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   7620
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox txtNomeUsuario 
         Height          =   285
         Left            =   2940
         MaxLength       =   30
         TabIndex        =   3
         Top             =   360
         Width           =   3165
      End
      Begin VB.TextBox txtCodUsuario 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1410
         TabIndex        =   2
         Top             =   390
         Width           =   825
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Resetar Senha"
         Height          =   315
         Left            =   7560
         TabIndex        =   1
         Top             =   750
         Width           =   1215
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   225
         Left            =   240
         TabIndex        =   5
         Top             =   1830
         Width           =   30
         _Version        =   65536
         _ExtentX        =   -53
         _ExtentY        =   397
         _StockProps     =   14
         Caption         =   "SSFrame2"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck chkAtivoUsuario 
         Height          =   375
         Left            =   9000
         TabIndex        =   6
         Top             =   360
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Conta Ativa"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.74
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Conta do Usuário"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   390
         Width           =   495
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Senha do Usuário"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6300
         TabIndex        =   7
         Top             =   420
         Width           =   1335
      End
   End
   Begin Threed.SSPanel frmToolBar 
      Height          =   555
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10425
      _Version        =   65536
      _ExtentX        =   18389
      _ExtentY        =   979
      _StockProps     =   15
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   2
      BevelOuter      =   0
      BevelInner      =   2
      Begin Threed.SSCommand cmdSair 
         Height          =   375
         Left            =   9660
         TabIndex        =   11
         Top             =   90
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         AutoSize        =   1
         Picture         =   "DesafioAC02.frx":0000
      End
      Begin Threed.SSCommand cmdPrimeiro 
         Height          =   375
         Left            =   5310
         TabIndex        =   12
         Top             =   90
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "DesafioAC02.frx":0A42
      End
      Begin Threed.SSCommand cmdAnterior 
         Height          =   375
         Left            =   5700
         TabIndex        =   13
         Top             =   90
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "DesafioAC02.frx":0F30
      End
      Begin Threed.SSCommand cmdProximo 
         Height          =   375
         Left            =   6090
         TabIndex        =   14
         Top             =   90
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "DesafioAC02.frx":141E
      End
      Begin Threed.SSCommand cmdUltimo 
         Height          =   375
         Left            =   6480
         TabIndex        =   15
         Top             =   90
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "DesafioAC02.frx":190C
      End
      Begin Threed.SSCommand cmdSalvar 
         Height          =   375
         Left            =   810
         TabIndex        =   16
         Top             =   90
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         AutoSize        =   1
         Picture         =   "DesafioAC02.frx":1DFA
      End
      Begin Threed.SSCommand cmdNovo 
         Height          =   375
         Left            =   210
         TabIndex        =   17
         Top             =   90
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         AutoSize        =   1
         Picture         =   "DesafioAC02.frx":1F0C
      End
      Begin Threed.SSCommand cmdExcluir 
         Height          =   375
         Left            =   1440
         TabIndex        =   18
         Top             =   90
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         AutoSize        =   1
         Picture         =   "DesafioAC02.frx":201E
      End
      Begin Threed.SSCommand cmdLocalizar 
         Height          =   375
         Left            =   3300
         TabIndex        =   19
         Top             =   90
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "DesafioAC02.frx":2550
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   20
      Top             =   1800
      Width           =   10455
      _Version        =   65536
      _ExtentX        =   18441
      _ExtentY        =   7011
      _StockProps     =   14
      Caption         =   "Permissões"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.ListBox LstTelas 
         Height          =   3375
         Left            =   150
         MultiSelect     =   2  'Extended
         TabIndex        =   22
         Top             =   510
         Width           =   4455
      End
      Begin VB.ListBox LstPermitidas 
         Height          =   3375
         Left            =   5850
         MultiSelect     =   2  'Extended
         TabIndex        =   21
         Top             =   480
         Width           =   4455
      End
      Begin Threed.SSCommand cmdAdicionar 
         Height          =   405
         Left            =   4620
         TabIndex        =   23
         Top             =   1110
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "&Adicionar ->"
         ForeColor       =   -2147483640
      End
      Begin Threed.SSCommand cmdRemover 
         Height          =   405
         Left            =   4620
         TabIndex        =   24
         Top             =   2610
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "<-&Remover"
         ForeColor       =   -2147483640
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telas com Permissão"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5940
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telas não Permitidas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Acao As String
    
'definir o recordset
Dim RS As New ADODB.Recordset
Dim SQL As String

Private Sub cmdReset_Click()

    If MsgBox("Deseja Ressetar a Senha do Usuário?", vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    'limpa preventivamente a senha
    txtSenhaUsuario.Text = ""
    
    'acao de RESET
    Acao = "R"
    
    'aqui inicia a gravacao no SQL
    grava_sql
    
    'limpa tela
    cmdLocalizar_Click
    
End Sub

Private Sub Form_Activate()
    
    'novo registro é o default da carga
    cmdNovo_Click

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   
   ' Troca a tecla "ENTER" por "TAB"
   MudaEnterToTab Me, KeyAscii

End Sub

Private Sub cmdNovo_Click()

    'definir a acao como inclusao
    Acao = "I"
    
    'esconder os controles desnecessarios
    cmdPrimeiro.Visible = False
    cmdAnterior.Visible = False
    cmdProximo.Visible = False
    cmdUltimo.Visible = False
    
    'limpar campos da tela para inclusão
    carrega_tela "L" 'limpar

End Sub

Sub carrega_tela(oque As String)
    
    'tratamento de erro
    On Error Resume Next
    
     'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If
    
    If oque = "C" Then
        
        'definir a acao como Alteracao
        Acao = "U"  'nao mexer
        
        SQL = "select * from usuario"
        SQL = SQL & " where CodUsuario = " & SQLDecimal(txtCodUsuario.Text)
        RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
        DB.Execute SQL, , adExecuteNoRecords
        If DB.Errors.Count > 0 Then
            RS.Close
            MsgBox "Problema na Leitura da Tabela [Usuario] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
            Exit Sub
        End If
                
        'para carregar todos os dados
        txtCodUsuario = RS!CodUsuario
        txtNomeUsuario.Text = Trim(RS!NomeUsuario)
        txtSenhaUsuario.Text = FormatCripto(Trim(RS!SenhaUsuario), False)
        chkAtivoUsuario.Value = IIf(RS!AtivoUsuario = "S", True, False)

        RS.Close
        
        carrega_listas
    
    Else
        'definir a acao como Insercao
        Acao = "I"  'nao mexer
        
        'limpar todos os dados da tela se chave for identity
        'se nao for limpar todos os dados que nao sao chave primaria
        txtCodUsuario.Text = ""
        txtNomeUsuario.Text = ""
        txtSenhaUsuario.Text = ""
        chkAtivoUsuario.Value = False
        LstTelas.Clear
        LstPermitidas.Clear
        txtNomeUsuario.SetFocus
    
    End If

End Sub

Private Sub cmdSalvar_Click()
    
    'validar a digitacao dos campos obrigatorios
    If Len(Trim(txtNomeUsuario.Text)) = 0 Then
        MsgBox "Nome do Usuário Não Informado!"
        txtNomeUsuario.SetFocus
        Exit Sub
    End If
    
    If Len(txtSenhaUsuario.Text) <> 6 And Len(txtSenhaUsuario.Text) <> 0 Then
        MsgBox "Senha deve ter 6 caracteres!"
        txtSenhaUsuario.SetFocus
        Exit Sub
    End If
    
    'aqui inicia a gravacao no SQL
    grava_sql
    
    'default
    If Acao = "I" Then
        cmdNovo_Click
    Else
        cmdLocalizar_Click
    End If

End Sub

Private Sub cmdExcluir_Click()

    'verificar se os campos chave primaria estao preenchidos
    If Len(txtCodUsuario.Text) = 0 Then
        MsgBox "Não Há Codigo Informado para ser Excluido!"
        Exit Sub
    End If
   
    'nao mexer daqui para baixo
    Acao = "D"
    If confirma_exclusao() = True Then
        'aqui inicia a gravacao no SQL
        grava_sql
        
        cmdLocalizar_Click
    End If
End Sub

Public Function confirma_exclusao()
    
    If MsgBox("Tem Certeza que Deseja Excluir o Registro?", vbYesNo + vbDefaultButton2, "Confirmação") = vbYes Then
        confirma_exclusao = True
    Else
        confirma_exclusao = False
    End If

End Function

Sub grava_sql()
    
    'tratamento de erro
    On Error Resume Next
    
     'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If
    
    'ínicia a transacao
    DB.BeginTrans
    
    'monta a string SQL
    If Acao = "I" Then
        SQL = "insert into Usuario (SenhaUsuario, NomeUsuario, AtivoUsuario) values ("
        
        If Len(Trim(txtSenhaUsuario.Text)) = 0 Then
            SQL = SQL & aspas("*RESETPWD*") & ","
        Else
            SQL = SQL & aspas(FormatCripto(UCase$(Trim(txtSenhaUsuario.Text)), True)) & ","
        End If
        
        SQL = SQL & aspas(Trim(txtNomeUsuario.Text)) & ","
        SQL = SQL & aspas(IIf(chkAtivoUsuario.Value, "S", "N")) & ")"
        
    ElseIf Acao = "U" Or Acao = "R" Then

        'atualizacao
        SQL = "update Usuario set "
        If Len(Trim(txtSenhaUsuario.Text)) = 0 Then
            SQL = SQL & "SenhaUsuario = " & aspas("*RESETPWD*") & ","
        Else
            SQL = SQL & "SenhaUsuario = " & aspas(FormatCripto(UCase$(Trim(txtSenhaUsuario.Text)), True)) & ","
        End If
        
        SQL = SQL & "NomeUsuario = " & aspas(Trim(txtNomeUsuario.Text)) & ","
        SQL = SQL & "AtivoUsuario = " & aspas(IIf(chkAtivoUsuario.Value, "S", "N"))
        SQL = SQL & " where CodUsuario = " & SQLDecimal(txtCodUsuario.Text)
        
    ElseIf Acao = "D" Then
        
        'excluir
        SQL = "delete Usuario "
        SQL = SQL & " where CodUsuario = " & SQLDecimal(txtCodUsuario.Text)
        
    End If
    
    'atualizar database
    DB.Execute SQL, , adExecuteNoRecords
    If DB.Errors.Count > 0 Then

        'desmancha transacao
        DB.RollbackTrans
        
        MsgBox "Problema na Inclusão da Tabela [Usuario] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
        Exit Sub
    
    End If
    
    'fechar a transacao
    DB.CommitTrans
    
    If Acao = "I" Then
       'so para insersão com campos identity
        MsgBox "Registro Incluido com Sucesso !"
    
    ElseIf Acao = "U" Then
        MsgBox "Registro Alterado com Sucesso!"
    
    ElseIf Acao = "D" Then
        MsgBox "Registro Excluido com Sucesso!"
    
    ElseIf Acao = "R" Then
        MsgBox "Senha Ressetada com Sucesso!"
    
    End If

End Sub

Sub carrega_listas()
    
    'tratamento de erro
    On Error Resume Next
    
    Me.MousePointer = 11

    'limpa preventivamente as listas
    LstPermitidas.Clear
    LstTelas.Clear
    
    'buscar todos as telas ja existentes no usuario
    SQL = "select p.CodTela, t.DescricaoTela from Permissao p inner join Telas t"
    SQL = SQL & " on p.CodTela = t.CodTela"
    SQL = SQL & " where p.CodUsuario = " & SQLDecimal(txtCodUsuario.Text)
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    If DB.Errors.Count > 0 Then
        RS.Close
        Me.MousePointer = 0
        MsgBox "Problema na Leitura da Tabela [Usuario] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
        Exit Sub
    End If
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    Do While Not RS.EOF
        LstPermitidas.AddItem RS!CodTela & " - " & RS!DescricaoTela
        RS.MoveNext
    Loop
    Me.MousePointer = 0
    RS.Close
    
    Me.MousePointer = 11

    'buscar todas as telas disponíveis do usuario
    SQL = "select * from Telas"
    SQL = SQL & " where not exists ("
    SQL = SQL & " select CodTela from Permissao "
    SQL = SQL & " where CodUsuario = " & SQLDecimal(txtCodUsuario.Text)
    SQL = SQL & " and Telas.CodTela = Permissao.CodTela )"
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    If DB.Errors.Count > 0 Then
        RS.Close
        Me.MousePointer = 0
        MsgBox "Problema na Leitura da Tabela [Telas] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
        Exit Sub
    End If
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    Do While Not RS.EOF
        LstTelas.AddItem RS!CodTela & " - " & RS!DescricaoTela
        RS.MoveNext
    Loop
    Me.MousePointer = 0
    RS.Close

    cmdAdicionar.Enabled = False
    cmdRemover.Enabled = False

End Sub


Private Sub cmdAdicionar_Click()
    
    'tratamento de erro
    On Error Resume Next
    
     'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If
    
    Dim b As Integer
    
    For b = 0 To LstTelas.ListCount - 1
        If LstTelas.Selected(b) = True Then
        
            Me.MousePointer = 11
            
            'ínicia a transacao
            DB.BeginTrans
            
            SQL = "insert Permissao (CodUsuario, CodTela) values ("
            SQL = SQL & SQLDecimal(txtCodUsuario.Text) & ","
            SQL = SQL & aspas(Mid$(LstTelas.List(b), 1, 4)) & ")"
            'atualizar database
            DB.Execute SQL, , adExecuteNoRecords
            If DB.Errors.Count > 0 Then
        
                'desmancha transacao
                DB.RollbackTrans
                
                MsgBox "Problema na Inclusão da Tabela [Permissao] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
                Exit Sub
            
            End If

            Me.MousePointer = 0
         
         End If
    Next b
    
    'fechar a transacao
    DB.CommitTrans
    
    carrega_listas
    
    cmdAdicionar.Enabled = False
    cmdRemover.Enabled = False

End Sub

Private Sub cmdAnterior_Click()
    
    'tratamento de erro
    On Error Resume Next
    
     'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If
    
    'nao mexer daqui para baixo
    Me.MousePointer = 11
    
    Dim Registro As Integer
    
    'verfica se o registro não é o primeiro
    SQL = "select count(*) as registro from Usuario"
    SQL = SQL & " where CodUsuario = "
    SQL = SQL & " (select max(CodUsuario) from Usuario "
    SQL = SQL & " where CodUsuario < " & SQLDecimal(txtCodUsuario.Text) & ")"
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    If DB.Errors.Count > 0 Then
        RS.Close
        Me.MousePointer = 0
        MsgBox "Problema na Leitura da Tabela [Usuario] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
        Exit Sub
    End If
    
    Me.MousePointer = 0
    If RS.EOF() = True Then
        RS.Close
        cmdNovo_Click
    Else
        'registro
        Registro = RS!Registro
        RS.Close
        
        'se registo = 0 recupera o primeiro
        If Registro = 0 Then
            Call cmdPrimeiro_Click
        Else
        
            'recupera o anterior
            SQL = "select * from Usuario"
            SQL = SQL & " where CodUsuario = "
            SQL = SQL & " (select max(CodUsuario) from Usuario"
            SQL = SQL & " where CodUsuario < " & SQLDecimal(txtCodUsuario.Text) & ")"
            RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If DB.Errors.Count > 0 Then
                RS.Close
                Me.MousePointer = 0
                MsgBox "Problema na Leitura da Tabela [Usuario] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
                Exit Sub
            End If
            
            Me.MousePointer = 0
            If RS.EOF() = True Then
                RS.Close
                cmdNovo_Click
            Else
                'codigo
                txtCodUsuario.Text = RS!CodUsuario
                RS.Close
            End If
            
        End If
        
        'carregar demais campos da tela
        carrega_tela "C"  'carregar tela
        
    End If
    
End Sub


Private Sub cmdLocalizar_Click()
    
    'mostrar os controles necessarios
    cmdPrimeiro.Visible = True
    cmdAnterior.Visible = True
    cmdProximo.Visible = True
    cmdUltimo.Visible = True
    
    'limpar a tela
    carrega_tela "L"   'limpar a tela
    
    'colocar o foco no 1o. campo da chave primaria
    txtNomeUsuario.SetFocus

End Sub

Private Sub cmdPrimeiro_Click()

    'tratamento de erro
    On Error Resume Next
    
     'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If
    
    'nao mexer daqui para baixo
    Me.MousePointer = 11
    
    SQL = "declare @limite char(30)" & vbCrLf
    SQL = SQL & "select @limite = min( CodUsuario ) From Usuario" & vbCrLf
    SQL = SQL & "select * From Usuario where CodUsuario = @limite"
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    If DB.Errors.Count > 0 Then
        RS.Close
        Me.MousePointer = 0
        MsgBox "Problema na Leitura da Tabela [Usuario] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
        Exit Sub
    End If
    
    Me.MousePointer = 0
    If RS.EOF() = True Then
        RS.Close
        cmdNovo_Click
    Else
        'codigo
        txtCodUsuario.Text = RS!CodUsuario
        RS.Close
        
        'carregar demais campos da tela
        carrega_tela "C"  'carregar tela
        
    End If
    
End Sub

Private Sub cmdProximo_Click()
    
    'tratamento de erro
    On Error Resume Next
    
     'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If
    
    'nao mexer daqui para baixo
    Me.MousePointer = 11
    
    Dim Registro As Integer
    
    'verfica se o registro não é o primeiro
    SQL = "select count(*) as registro from Usuario"
    SQL = SQL & " where CodUsuario = "
    SQL = SQL & " (select min(CodUsuario) from Usuario "
    SQL = SQL & " where CodUsuario > " & SQLDecimal(txtCodUsuario.Text) & ")"
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    If DB.Errors.Count > 0 Then
        RS.Close
        Me.MousePointer = 0
        MsgBox "Problema na Leitura da Tabela [Usuario] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
        Exit Sub
    End If
    
    Me.MousePointer = 0
    If RS.EOF() = True Then
        RS.Close
        cmdNovo_Click
    Else
        'registro
        Registro = RS!Registro
        RS.Close
        
        'se registo = 0 recupera o primeiro
        If Registro = 0 Then
            Call cmdultimo_Click
        Else
        
            'recupera o anterior
            SQL = "select * from Usuario"
            SQL = SQL & " where CodUsuario = "
            SQL = SQL & " (select min(CodUsuario) from Usuario"
            SQL = SQL & " where CodUsuario > " & SQLDecimal(txtCodUsuario.Text) & ")"
            RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If DB.Errors.Count > 0 Then
                RS.Close
                Me.MousePointer = 0
                MsgBox "Problema na Leitura da Tabela [Usuario] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
                Exit Sub
            End If
            
            Me.MousePointer = 0
            If RS.EOF() = True Then
                RS.Close
                cmdNovo_Click
            Else
                'codigo
                txtCodUsuario.Text = RS!CodUsuario
                RS.Close
            End If
            
        End If
        
        'carregar demais campos da tela
        carrega_tela "C"  'carregar tela
        
    End If
    
End Sub

Private Sub cmdRemover_Click()
    
    'tratamento de erro
    On Error Resume Next
    
     'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If
    
    Dim b As Integer
    
    For b = 0 To LstPermitidas.ListCount - 1
        If LstPermitidas.Selected(b) = True Then
        
            Me.MousePointer = 11
            
            'ínicia a transacao
            DB.BeginTrans
            
            SQL = "delete Permissao"
            SQL = SQL & " where CodUsuario = " & SQLDecimal(txtCodUsuario.Text)
            SQL = SQL & " and CodTela = " & aspas(Mid$(LstPermitidas.List(b), 1, 4))
            'atualizar database
            DB.Execute SQL, , adExecuteNoRecords
            If DB.Errors.Count > 0 Then
        
                'desmancha transacao
                DB.RollbackTrans
                
                MsgBox "Problema na Exclusão da Tabela [Permissao] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
                Exit Sub
            
            End If

            Me.MousePointer = 0
         
         End If
    Next b
    
    'fechar a transacao
    DB.CommitTrans
    
    carrega_listas
    
    cmdAdicionar.Enabled = False
    cmdRemover.Enabled = False

End Sub

Private Sub cmdsair_click()
    Unload Me
End Sub


Private Sub cmdultimo_Click()
    
    'tratamento de erro
    On Error Resume Next
    
     'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If
    
    'nao mexer daqui para baixo
    Me.MousePointer = 11
    
    SQL = "declare @limite char(30)" & vbCrLf
    SQL = SQL & "select @limite = max( CodUsuario ) From Usuario" & vbCrLf
    SQL = SQL & "select * From Usuario where CodUsuario = @limite"
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    If DB.Errors.Count > 0 Then
        RS.Close
        MsgBox "Problema na Leitura da Tabela [Usuario] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
        Exit Sub
    End If
    
    Me.MousePointer = 0
    If RS.EOF() = True Then
        RS.Close
        cmdNovo_Click
    Else
        'codigo
        txtCodUsuario.Text = RS!CodUsuario
        RS.Close
        
        'carregar demais campos da tela
        carrega_tela "C"  'carregar tela
        
    End If
    
End Sub


Private Sub lstPermitidas_Click()
    cmdAdicionar.Enabled = False
    cmdRemover.Enabled = True
End Sub

Private Sub lstTelas_Click()
    cmdAdicionar.Enabled = True
    cmdRemover.Enabled = False
End Sub

Sub MudaEnterToTab(frm As Form, ByRef KeyAscii As Integer)
    If frm.ActiveControl.Tag = "LIB" Then
        Exit Sub
    End If
    Select Case KeyAscii
        Case vbKeyBack, vbKeySpace, vbKeyCancel, vbKeyTab, 22, 24
            Exit Sub
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{TAB}"
        Case Asc("a") To Asc("z")
            KeyAscii = KeyAscii - 32
        Case 34, Asc("^"), Asc("`"), Asc("´"), Asc("~")
            KeyAscii = 0
        Case Asc("ª"), Asc("º")
            KeyAscii = Asc("R")
        Case 192 To 197, 224 To 229 '(a acentuado)
            KeyAscii = Asc("A")
        Case 200 To 203, 232 To 235 '(e acentuado)
            KeyAscii = Asc("E")
        Case 204 To 207, 236 To 239 '(i acentuado)
            KeyAscii = Asc("I")
        Case 210 To 214, 242 To 246 '(o acentuado)
            KeyAscii = Asc("O")
        Case 217 To 220, 249 To 252 '(u acentuado)
            KeyAscii = Asc("U")
        Case 199, 231
            KeyAscii = Asc("C")
        Case Asc("'")
            KeyAscii = Asc("`")
    End Select
    If frm.ActiveControl.Tag = "N" Then
      If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then '0 a 9
          KeyAscii = 0
          Beep
       End If
    ElseIf frm.ActiveControl.Tag = "R" Then
       If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") Then '0 a 9
          If Chr(KeyAscii) = "." Then
             If InStr(frm.ActiveControl, ".") > 0 Then
                KeyAscii = 0
                Beep
             End If
           Else
              KeyAscii = 0
              Beep
           End If
       End If
    End If
End Sub


