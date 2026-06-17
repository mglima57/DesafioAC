VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Begin VB.Form frm01 
   Caption         =   "01 - Transações"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10620
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Transações"
      TabPicture(0)   =   "DesafioAC01.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmToolBar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "optCancelada"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "optPendente"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "optAprovada"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtDescricao"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDtTransacao"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtVlrTransacao"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtIdTransacao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtNumeroCartao"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Lista dos Registros"
      TabPicture(1)   =   "DesafioAC01.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(4)=   "cmdNovoPesq"
      Tab(1).Control(5)=   "cmdsairPesq"
      Tab(1).Control(6)=   "cmdLocalizarPesq"
      Tab(1).Control(7)=   "ListView1"
      Tab(1).Control(8)=   "txtNumeroCartaoPesq"
      Tab(1).Control(9)=   "txtDtTransacaoPesq"
      Tab(1).Control(10)=   "txtVlrTransacaoPesq"
      Tab(1).Control(11)=   "chkAprovada"
      Tab(1).Control(12)=   "chkPendente"
      Tab(1).Control(13)=   "chkCancelada"
      Tab(1).ControlCount=   14
      Begin VB.CheckBox chkCancelada 
         Caption         =   "Cancelada"
         Height          =   255
         Left            =   -68280
         TabIndex        =   44
         Top             =   870
         Width           =   1215
      End
      Begin VB.CheckBox chkPendente 
         Caption         =   "Pendente"
         Height          =   255
         Left            =   -69480
         TabIndex        =   43
         Top             =   870
         Width           =   1215
      End
      Begin VB.CheckBox chkAprovada 
         Caption         =   "Aprovada"
         Height          =   255
         Left            =   -70680
         TabIndex        =   42
         Top             =   870
         Width           =   1215
      End
      Begin VB.TextBox txtVlrTransacaoPesq 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   -73590
         TabIndex        =   39
         Top             =   840
         Width           =   1260
      End
      Begin VB.TextBox txtDtTransacaoPesq 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -68880
         MaxLength       =   8
         TabIndex        =   19
         Top             =   480
         Width           =   1125
      End
      Begin VB.TextBox txtNumeroCartaoPesq 
         Height          =   285
         Left            =   -73590
         MaxLength       =   16
         TabIndex        =   18
         Top             =   480
         Width           =   3165
      End
      Begin VB.TextBox txtNumeroCartao 
         Height          =   285
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   11
         Top             =   1650
         Width           =   3165
      End
      Begin VB.TextBox txtIdTransacao 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1200
         Width           =   1065
      End
      Begin VB.TextBox txtVlrTransacao 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1440
         TabIndex        =   13
         Top             =   4770
         Width           =   1260
      End
      Begin VB.TextBox txtDtTransacao 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         MaxLength       =   19
         TabIndex        =   14
         Text            =   "99/99/9999 99:99:99"
         Top             =   5250
         Width           =   1725
      End
      Begin VB.TextBox txtDescricao 
         Height          =   2595
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   12
         Top             =   2160
         Width           =   8715
      End
      Begin VB.OptionButton optAprovada 
         Caption         =   "Aprovada"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   5730
         Width           =   1575
      End
      Begin VB.OptionButton optPendente 
         Caption         =   "Pendente"
         Height          =   345
         Left            =   1440
         TabIndex        =   16
         Top             =   5970
         Width           =   1575
      End
      Begin VB.OptionButton optCancelada 
         Caption         =   "Cancelada"
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Frame frmCodigo 
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -74400
         TabIndex        =   24
         Top             =   1620
         Width           =   1635
         Begin VB.TextBox txtCodigo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   180
            MaxLength       =   3
            TabIndex        =   25
            Top             =   300
            Width           =   1275
         End
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   675
         Left            =   -71250
         TabIndex        =   26
         Top             =   570
         Visible         =   0   'False
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1191
         _StockProps     =   78
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "DesafioAC01.frx":0038
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   675
         Left            =   -70320
         TabIndex        =   27
         Top             =   570
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1191
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "DesafioAC01.frx":01D2
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   675
         Left            =   -69495
         TabIndex        =   28
         Top             =   570
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1191
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "DesafioAC01.frx":02E4
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   675
         Left            =   -68580
         TabIndex        =   29
         ToolTipText     =   "Sair da Tela"
         Top             =   570
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1191
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "DesafioAC01.frx":03F6
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5355
         Left            =   -74880
         TabIndex        =   23
         Top             =   1200
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   9446
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nome"
            Object.Width           =   2540
         EndProperty
      End
      Begin Threed.SSPanel frmToolBar 
         Height          =   555
         Left            =   120
         TabIndex        =   36
         Top             =   480
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
            TabIndex        =   9
            Top             =   90
            Width           =   405
            _Version        =   65536
            _ExtentX        =   714
            _ExtentY        =   661
            _StockProps     =   78
            ForeColor       =   -2147483640
            BevelWidth      =   1
            AutoSize        =   1
            Picture         =   "DesafioAC01.frx":0748
         End
         Begin Threed.SSCommand cmdPrimeiro 
            Height          =   375
            Left            =   5310
            TabIndex        =   5
            Top             =   90
            Width           =   405
            _Version        =   65536
            _ExtentX        =   714
            _ExtentY        =   661
            _StockProps     =   78
            ForeColor       =   -2147483640
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Picture         =   "DesafioAC01.frx":118A
         End
         Begin Threed.SSCommand cmdAnterior 
            Height          =   375
            Left            =   5700
            TabIndex        =   6
            Top             =   90
            Width           =   405
            _Version        =   65536
            _ExtentX        =   714
            _ExtentY        =   661
            _StockProps     =   78
            ForeColor       =   -2147483640
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Picture         =   "DesafioAC01.frx":1678
         End
         Begin Threed.SSCommand cmdProximo 
            Height          =   375
            Left            =   6120
            TabIndex        =   7
            Top             =   90
            Width           =   405
            _Version        =   65536
            _ExtentX        =   714
            _ExtentY        =   661
            _StockProps     =   78
            ForeColor       =   -2147483640
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Picture         =   "DesafioAC01.frx":1B66
         End
         Begin Threed.SSCommand cmdUltimo 
            Height          =   375
            Left            =   6480
            TabIndex        =   8
            Top             =   90
            Width           =   405
            _Version        =   65536
            _ExtentX        =   714
            _ExtentY        =   661
            _StockProps     =   78
            ForeColor       =   -2147483640
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Picture         =   "DesafioAC01.frx":2054
         End
         Begin Threed.SSCommand cmdSalvar 
            Height          =   375
            Left            =   810
            TabIndex        =   2
            Top             =   90
            Width           =   405
            _Version        =   65536
            _ExtentX        =   714
            _ExtentY        =   661
            _StockProps     =   78
            ForeColor       =   -2147483640
            BevelWidth      =   1
            AutoSize        =   1
            Picture         =   "DesafioAC01.frx":2542
         End
         Begin Threed.SSCommand cmdNovo 
            Height          =   375
            Left            =   210
            TabIndex        =   1
            Top             =   90
            Width           =   405
            _Version        =   65536
            _ExtentX        =   714
            _ExtentY        =   661
            _StockProps     =   78
            ForeColor       =   -2147483640
            BevelWidth      =   1
            AutoSize        =   1
            Picture         =   "DesafioAC01.frx":2654
         End
         Begin Threed.SSCommand cmdExcluir 
            Height          =   375
            Left            =   1440
            TabIndex        =   3
            Top             =   90
            Width           =   405
            _Version        =   65536
            _ExtentX        =   714
            _ExtentY        =   661
            _StockProps     =   78
            ForeColor       =   -2147483640
            BevelWidth      =   1
            AutoSize        =   1
            Picture         =   "DesafioAC01.frx":2766
         End
         Begin Threed.SSCommand cmdLocalizar 
            Height          =   375
            Left            =   3300
            TabIndex        =   4
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
            Picture         =   "DesafioAC01.frx":2C98
         End
      End
      Begin Threed.SSCommand cmdLocalizarPesq 
         Height          =   375
         Left            =   -65520
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   480
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "DesafioAC01.frx":2DAA
      End
      Begin Threed.SSCommand cmdsairPesq 
         Height          =   375
         Left            =   -65040
         TabIndex        =   22
         Top             =   480
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         AutoSize        =   1
         Picture         =   "DesafioAC01.frx":2EBC
      End
      Begin Threed.SSCommand cmdNovoPesq 
         Height          =   375
         Left            =   -66000
         TabIndex        =   20
         Top             =   480
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         AutoSize        =   1
         Picture         =   "DesafioAC01.frx":38FE
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72120
         TabIndex        =   41
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ValorTransação"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   40
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data Transação"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -70200
         TabIndex        =   38
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero Cartao "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Id_Transacao"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   35
         Top             =   1230
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero Cartao "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   34
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ValorTransação"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data Transação"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descrição"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   5760
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Acao As String
    
'definir o recordset
Dim RS As New ADODB.Recordset
Dim SQL As String

'PrimeiroAcesso
Dim PrimeiroAcesso As Boolean

Private Sub cmdNovoPesq_Click()
    txtNumeroCartaoPesq.Text = ""
    txtDtTransacaoPesq.Text = ""
    txtVlrTransacaoPesq.Text = ""
    chkAprovada.Value = 0
    chkPendente.Value = 0
    chkCancelada.Value = 0
    
End Sub

Private Sub cmdsairPesq_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    'primeiro acesso
    PrimeiroAcesso = True
    
    'novo registro é o default da carga
    cmdNovo_Click

    'primeiro acesso
    PrimeiroAcesso = False

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
        
        SQL = "select * from Transacoes"
        SQL = SQL & " where Id_Transacao = " & txtIdTransacao.Text
        RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
        DB.Execute SQL, , adExecuteNoRecords
        If DB.Errors.Count > 0 Then
            RS.Close
            MsgBox "Problema na Leitura da Tabela [Transacoes] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
            
            'grava LogOperacao
            GravaLogOperacao "01", usuario_logado, SQL, DB
            
            Exit Sub
        End If
                
        'para carregar todos os dados
        txtIdTransacao.Text = RS!Id_Transacao
        txtNumeroCartao.Text = RS!Numero_Cartao
        txtDescricao.Text = Trim(RS!Descricao)
        txtVlrTransacao.Text = Format(RS!Valor_Transacao, "##,###,##0.00")
        txtDtTransacao.Text = RS!Data_Transacao
        optAprovada = IIf(RS!Status_Transacao = "Aprovada", True, False)
        optPendente = IIf(RS!Status_Transacao = "Pendente", True, False)
        optCancelada = IIf(RS!Status_Transacao = "Cancelada", True, False)
        
        RS.Close
        
    Else
        'definir a acao como Insercao
        Acao = "I"  'nao mexer
        
        'limpar todos os dados da tela se chave for identity
        'se nao for limpar todos os dados que nao sao chave primaria
        txtIdTransacao.Text = ""
        txtNumeroCartao.Text = ""
        txtDescricao.Text = ""
        txtVlrTransacao.Text = ""
        txtDtTransacao.Text = Format(Now(), "dd/mm/yyyy hh:mm:ss")
        optAprovada = False
        optPendente = False
        optCancelada = False
        
        txtNumeroCartao.SetFocus
    
    End If

    'libera/bloqueia
    Call liberaBloqueia

End Sub

Private Sub cmdSalvar_Click()
    
    'validar a digitacao dos campos obrigatorios
    If Len(Trim(txtNumeroCartao.Text)) = 0 Then
        MsgBox "Numero do Cartão Não Informado!"
        txtNumeroCartao.SetFocus
        Exit Sub
    End If
    
    If Len(txtDescricao.Text) = 0 Then
        MsgBox "Descrição Não Informada!"
        txtDescricao.SetFocus
        Exit Sub
    End If
    
    If MyVal(txtVlrTransacao.Text) = 0 Then
        MsgBox "Valor da Transação Não Informado!"
        txtVlrTransacao.SetFocus
        Exit Sub
    End If
    
    If Len(txtDtTransacao.Text) = 0 Then
        MsgBox "Data/Hora da Transação Não Informada!"
        txtDtTransacao.SetFocus
        Exit Sub
    End If
    
    If optAprovada = False And optPendente = False And optCancelada = False Then
        MsgBox "Satus Não Informado!"
        Exit Sub
    End If
    
    'Permitir edição de qualquer campo exceto transações com status 'Aprovada'.
     If Acao = "U" And optAprovada.Value = True Then  'nao mexer
        MsgBox "Não é permitido efetuar alterações em Transações com status 'APROVADO'"
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
    If Len(txtNumeroCartao.Text) = 0 Then
        MsgBox "Não Há Numero de Cartão Informado para ser Excluido!"
        Exit Sub
    End If
    
    If MyVal(txtVlrTransacao.Text) = 0 Then
        MsgBox "Não Há Valor da Transação Informado para ser Excluido!"
        Exit Sub
    End If
    
    If Len(txtDtTransacao.Text) = 0 Then
        MsgBox "Não Há Data/Hora da Transação Informada para ser Excluido!"
        MsgBox "Data/Hora da Transação Não Informada!"
        txtDtTransacao.SetFocus
        Exit Sub
    End If
   
    'nao mexer daqui para baixo
    Acao = "D"
    If confirma_exclusao() = True Then
        'aqui inicia a gravacao no SQL
        grava_sql
        
    End If

    Call cmdLocalizar_Click
End Sub

Public Function confirma_exclusao()
    
    If MsgBox("Tem Certeza que Deseja Excluir o Registro?", vbYesNo + vbDefaultButton2, "Confirmação") = vbYes Then
        confirma_exclusao = True
    Else
        confirma_exclusao = False
    End If

End Function

Sub grava_sql()
    
    Dim DescricaoAcao As String
    
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
        SQL = "insert into Transacoes (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao) values ("
        
        SQL = SQL & aspas(Trim(txtNumeroCartao.Text)) & ","
        SQL = SQL & SQLDecimal(Trim(txtVlrTransacao.Text)) & ","
        SQL = SQL & aspas(Format(txtDtTransacao.Text, "yyyy/mm/dd hh:mm:ss")) & ","
        SQL = SQL & aspas(Trim(txtDescricao.Text)) & ","
        SQL = SQL & aspas(IIf(optAprovada.Value = True, aspas("Aprovada"), _
                          IIf(optPendente.Value = True, aspas("Pendente"), _
                          IIf(optCancelada.Value = True, aspas("Cancelada"), "")))) & ")"
        
    ElseIf Acao = "U" Then

        'atualizacao
        SQL = "update Transacoes set "
        
        SQL = SQL & "Numero_Cartao = " & aspas(Trim(txtNumeroCartao.Text)) & ","
        SQL = SQL & "Valor_Transacao = " & SQLDecimal(Trim(txtVlrTransacao.Text)) & ","
        SQL = SQL & "Data_Transacao = " & aspas(Format(txtDtTransacao.Text, "yyyy/mm/dd hh:mm:ss")) & ","
        SQL = SQL & "Descricao = " & aspas(Trim(txtDescricao.Text)) & ","
        SQL = SQL & "Status_Transacao = " & aspas(IIf(optAprovada.Value = True, aspas("Aprovada"), _
                                                  IIf(optPendente.Value = True, aspas("Pendente"), _
                                                  IIf(optCancelada.Value = True, aspas("Cancelada"), ""))))
        SQL = SQL & " where Id_Transacao   = " & txtIdTransacao.Text
        
    ElseIf Acao = "D" Then
        
        'excluir
        SQL = "delete Transacoes "
        SQL = SQL & " where Id_Transacao   = " & txtIdTransacao.Text
        
    End If
    
    'atualizar database
    DB.Execute SQL, , adExecuteNoRecords
    If DB.Errors.Count > 0 Then

        'desmancha transacao
        DB.RollbackTrans
        
        If Acao = "I" Then DescricaoAcao = "Inclusão"
        If Acao = "U" Then DescricaoAcao = "Atualização"
        If Acao = "D" Then DescricaoAcao = "Exclusão"
        
        MsgBox "Problema na " & DescricaoAcao & " da Tabela [Transacoes] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
        
        GravaLogOperacao "01", usuario_logado, SQL, DB
        
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
       
    End If

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
    SQL = "select count(*) as registro from Transacoes"
    SQL = SQL & " where Id_Transacao = "
    SQL = SQL & " (select max(Id_Transacao) from Transacoes "
    SQL = SQL & " where Id_Transacao < " & SQLDecimal(txtIdTransacao.Text) & ")"
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    If DB.Errors.Count > 0 Then
        RS.Close
        Me.MousePointer = 0
        MsgBox "Problema na Leitura da Tabela [Transacoes] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
        
        'grava LogOperacao
        GravaLogOperacao "01", usuario_logado, SQL, DB
        
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
            SQL = "select * from Transacoes"
            SQL = SQL & " where Id_Transacao = "
            SQL = SQL & " (select max(Id_Transacao) from Transacoes"
            SQL = SQL & " where Id_Transacao < " & SQLDecimal(txtIdTransacao.Text) & ")"
            RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If DB.Errors.Count > 0 Then
                RS.Close
                Me.MousePointer = 0
                MsgBox "Problema na Leitura da Tabela [Transacoes] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
                
                'grava LogOperacao
                GravaLogOperacao "01", usuario_logado, SQL, DB
                
                Exit Sub
            End If
            
            Me.MousePointer = 0
            If RS.EOF() = True Then
                RS.Close
                cmdNovo_Click
            Else
                'codigo
                txtIdTransacao.Text = RS!Id_Transacao
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
    txtNumeroCartao.SetFocus

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
    SQL = SQL & "select @limite = min( Id_Transacao ) From Transacoes" & vbCrLf
    SQL = SQL & "select * From Transacoes where Id_Transacao = @limite"
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    If DB.Errors.Count > 0 Then
        RS.Close
        Me.MousePointer = 0
        MsgBox "Problema na Leitura da Tabela [Transacoes] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
        
        'grava LogOperacao
        GravaLogOperacao "01", usuario_logado, SQL, DB
    
        Exit Sub
    End If
    
    Me.MousePointer = 0
    If RS.EOF() = True Then
        RS.Close
        cmdNovo_Click
    Else
        'codigo
        txtIdTransacao.Text = RS!Id_Transacao
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
    SQL = "select count(*) as registro from Transacoes"
    SQL = SQL & " where Id_Transacao = "
    SQL = SQL & " (select min(Id_Transacao) from Transacoes "
    SQL = SQL & " where Id_Transacao > " & SQLDecimal(txtIdTransacao.Text) & ")"
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    If DB.Errors.Count > 0 Then
        RS.Close
        Me.MousePointer = 0
        MsgBox "Problema na Leitura da Tabela [Transacoes] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
        
        'grava LogOperacao
        GravaLogOperacao "01", usuario_logado, SQL, DB
    
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
            SQL = "select * from Transacoes"
            SQL = SQL & " where Id_Transacao = "
            SQL = SQL & " (select min(Id_Transacao) from Transacoes"
            SQL = SQL & " where Id_Transacao > " & SQLDecimal(txtIdTransacao.Text) & ")"
            RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If DB.Errors.Count > 0 Then
                RS.Close
                Me.MousePointer = 0
                MsgBox "Problema na Leitura da Tabela [Transacoes] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
                
                'grava LogOperacao
                GravaLogOperacao "01", usuario_logado, SQL, DB
                
                Exit Sub
            End If
            
            Me.MousePointer = 0
            If RS.EOF() = True Then
                RS.Close
                cmdNovo_Click
            Else
                'codigo
                txtIdTransacao.Text = RS!Id_Transacao
                RS.Close
            End If
            
        End If
        
        'carregar demais campos da tela
        carrega_tela "C"  'carregar tela
        
    End If
    
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
    SQL = SQL & "select @limite = max( Id_Transacao ) From Transacoes" & vbCrLf
    SQL = SQL & "select * From Transacoes where Id_Transacao = @limite"
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    If DB.Errors.Count > 0 Then
        RS.Close
        MsgBox "Problema na Leitura da Tabela [Transacoes] " & DB.Errors.Item(0).Description & " //Programa: " & UCase$(Trim(App.EXEName)) & " // Conexão: " & Trim(DB.ConnectionString), vbCritical, "T I R E   U M   P R I N T - S C R E E N   D E S T A   M E N S A G E M"
        
        'grava LogOperacao
        GravaLogOperacao "01", usuario_logado, SQL, DB
        
        Exit Sub
    End If
    
    Me.MousePointer = 0
    If RS.EOF() = True Then
        RS.Close
        cmdNovo_Click
    Else
        'codigo
        txtIdTransacao.Text = RS!Id_Transacao
        RS.Close
        
        'carregar demais campos da tela
        carrega_tela "C"  'carregar tela
        
    End If
    
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

Private Sub liberaBloqueia()

    'Tratamento de erro
    On Error Resume Next

    'libera otimista
    txtNumeroCartao.Enabled = True
    txtDescricao.Enabled = True
    txtVlrTransacao.Enabled = True
    txtDtTransacao.Enabled = True
    optAprovada.Enabled = True
    optPendente.Enabled = True
    optCancelada.Enabled = True
    cmdNovo.Visible = True
    cmdSalvar.Visible = True
    cmdExcluir.Visible = True

    'Permitir edição de qualquer campo exceto transações com status 'Aprovada'.
    If optAprovada.Value = True Then  'nao mexer
        txtNumeroCartao.Enabled = False
        txtDescricao.Enabled = False
        txtVlrTransacao.Enabled = False
        txtDtTransacao.Enabled = False
        optAprovada.Enabled = False
        optPendente.Enabled = False
        optCancelada.Enabled = False
        cmdSalvar.Visible = False
        cmdExcluir.Visible = False
    End If

End Sub

Private Sub txtDtTransacaoPesq_LostFocus()

    'Tratamento de erro
    On Error Resume Next

    'formata data
    txtDtTransacaoPesq.Text = Replace(Replace(Replace(Trim(txtDtTransacaoPesq.Text), ".", ""), "-", ""), "/", "")
    If Len(txtDtTransacaoPesq.Text) = 8 Then
        txtDtTransacaoPesq.Text = Format(Mid(txtDtTransacaoPesq.Text, 1, 2) & "/" & Mid(txtDtTransacaoPesq.Text, 3, 2) & "/" & Mid(txtDtTransacaoPesq.Text, 5, 4), "dd/mm/YYYY")
    Else
        MsgBox "Data Invalida", vbInformation, "MENSAGEM"
        txtDtTransacaoPesq.Text = ""
    End If
    
End Sub

Private Sub txtNumeroCartao_KeyPress(KeyAscii As Integer)
    
    'Tratamento de erro
    On Error Resume Next
    
    If KeyAscii <> 8 And KeyAscii <> 44 Then
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

End Sub

Private Sub txtVlrTransacao_KeyPress(KeyAscii As Integer)
    
    'Tratamento de erro
    On Error Resume Next
    
    If KeyAscii <> 8 And KeyAscii <> 44 Then
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

End Sub

Private Sub txtVlrTransacao_GotFocus()

    'Tratamento de erro
    On Error Resume Next
    
    'verifica se pegou ou total ou seja ja tem virgula
    Dim virgula() As String
    Dim valor As String
    valor = txtVlrTransacao.Text
    
    virgula = Split(valor, ",")
    If UBound(virgula) = 0 Then
        'formata o valor sem decimal
        valor = MyVal(valor) / 100
    End If

    'formata o valor
    If Len(valor) > 0 Then
        txtVlrTransacao.Text = Format(MyVal(valor), "#######0.00")
    End If

End Sub

Private Sub txtVlrTransacao_LostFocus()

    'Tratamento de erro
    On Error Resume Next
    
    'verifica se pegou ou total ou seja ja tem virgula
    Dim virgula() As String
    virgula = Split(txtVlrTransacao.Text, ",")
    If UBound(virgula) = 0 Then
        'formata o valor
        txtVlrTransacao.Text = Format(MyVal(txtVlrTransacao.Text) / 100, "##,###,##0.00")
    Else
        'formata o valor
        txtVlrTransacao.Text = Format(MyVal(txtVlrTransacao.Text), "##,###,##0.00")
    End If
   
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ListView1.SortKey <> ColumnHeader.Index - 1 Then
        ListView1.SortKey = ColumnHeader.Index - 1
        ListView1.SortOrder = lvwAscending
    Else
        If ListView1.SortOrder = lvwAscending Then
            ListView1.SortOrder = lvwDescending
        Else
            ListView1.SortOrder = lvwAscending
        End If
    End If
    ListView1.Sorted = -1

End Sub

Private Sub ListView1_DblClick()
    SSTab1.Tab = 0
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

    'Tratamento de erro
    On Error Resume Next
    
    'dimensionamento/criaçao do recordset
    Dim RS As ADODB.Recordset, SQL As String
    
    'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
       Exit Sub
    End If
    
    Set RS = New ADODB.Recordset
    
    'abertura do recordset
    SQL = "select * from Transacoes where Id_Transacao = " & aspas(ListView1.ListItems(ListView1.SelectedItem.Index).Text)
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    txtIdTransacao = RS!Id_Transacao
    RS.Close
    Set RS = Nothing
    
    'carregar demais campos da tela
    carrega_tela "C"  'carregar tela

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    'Tratamento de erro
    On Error Resume Next
    
    If SSTab1.Tab = 1 Then    'lista
        
        'limpa preventivamente a pesquisa
        cmdNovoPesq_Click
        
        'parametrizar a ListView
        ListView1.View = lvwReport
        ListView1.FullRowSelect = True
        ListView1.LabelEdit = lvwManual
        Dim objItensLista As MSComctlLib.ListItem
        
        'limpar a listview
        ListView1.ListItems.Clear
    
        'colocar as colunas
        ListView1.ColumnHeaders.Clear
        ListView1.GridLines = True
        ListView1.ColumnHeaders.Add , "Id_Transacao", "Id Transacao"
        ListView1.ColumnHeaders.Add , "Numero_Cartao", "Numero Cartao"
        ListView1.ColumnHeaders.Add , "Valor_Transacao", "Valor Transacao"
        ListView1.ColumnHeaders.Add , "Data_Transacao", "Data Transacao"
        ListView1.ColumnHeaders.Add , "Status_Transacao", "Status"
        ListView1.ColumnHeaders.Add , "Descricao", "Descricao"
    
        'formata as colunas
        ListView1.ColumnHeaders.Item(1).Width = 1500
        ListView1.ColumnHeaders.Item(2).Width = 2000
        ListView1.ColumnHeaders.Item(3).Width = 2000
        ListView1.ColumnHeaders.Item(4).Width = 2000
        ListView1.ColumnHeaders.Item(5).Width = 2000
        ListView1.ColumnHeaders.Item(6).Width = 5000
        
        'abertura do recordset
        SQL = "select * from Transacoes"
        SQL = SQL & " order by Numero_Cartao, Data_Transacao, Valor_Transacao"
        RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
        
        
        Do While Not RS.EOF
            Set objItensLista = ListView1.ListItems.Add(, , RS!Id_Transacao)
            objItensLista.SubItems(1) = RS!Numero_Cartao
            objItensLista.SubItems(2) = Format(MyVal(RS!Valor_Transacao), "##,###,##0.00")
            objItensLista.SubItems(3) = Format(RS!Data_Transacao, "dd/MM/yyyy hh:mm:ss")
            objItensLista.SubItems(4) = RS!Status_Transacao
            objItensLista.SubItems(5) = RS!Descricao
            RS.MoveNext
        Loop
        
        'fechamento do recordset
        RS.Close
        Set RS = Nothing
    End If

End Sub

Private Sub cmdLocalizarPesq_Click()

    'Tratamento de erro
    On Error Resume Next
    
    Dim parametrosPesquisa As String
    Dim temStatus As Boolean
    parametrosPesquisa = ""
    temStatus = False
    
    'parametrizar a ListView
    ListView1.View = lvwReport
    ListView1.FullRowSelect = True
    ListView1.LabelEdit = lvwManual
    Dim objItensLista As MSComctlLib.ListItem
    
    'limpar a listview
    ListView1.ListItems.Clear

    'colocar as colunas
    ListView1.ColumnHeaders.Clear
    ListView1.GridLines = True
    ListView1.ColumnHeaders.Add , "Id_Transacao", "Id Transacao"
    ListView1.ColumnHeaders.Add , "Numero_Cartao", "Numero Cartao"
    ListView1.ColumnHeaders.Add , "Valor_Transacao", "Valor Transacao"
    ListView1.ColumnHeaders.Add , "Data_Transacao", "Data Transacao"
    ListView1.ColumnHeaders.Add , "Status_Transacao", "Status"
    ListView1.ColumnHeaders.Add , "Descricao", "Descricao"

    'formata as colunas
    ListView1.ColumnHeaders.Item(1).Width = 1500
    ListView1.ColumnHeaders.Item(2).Width = 2000
    ListView1.ColumnHeaders.Item(3).Width = 2000
    ListView1.ColumnHeaders.Item(4).Width = 2000
    ListView1.ColumnHeaders.Item(5).Width = 2000
    ListView1.ColumnHeaders.Item(6).Width = 5000
    
    'abertura do recordset
    SQL = "select * from Transacoes"
    
    'se foi selecionado algum parametro
    If Len(txtNumeroCartaoPesq.Text) > 0 Or _
       Len(txtDtTransacaoPesq.Text) > 0 Or _
       Len(txtVlrTransacaoPesq.Text) > 0 Or _
       chkAprovada.Value = 1 Or _
       chkPendente.Value = 1 Or _
       chkCancelada.Value = 1 Then
    
        'parametros selecionados
        If Len(txtNumeroCartaoPesq.Text) > 0 Then
            parametrosPesquisa = parametrosPesquisa & " Numero_Cartao like '%" & Trim(txtNumeroCartaoPesq.Text) & "%'"
        End If
        
        If Len(txtDtTransacaoPesq.Text) > 0 Then
            If Len(Trim(parametrosPesquisa)) > 0 Then
                parametrosPesquisa = parametrosPesquisa & " and "
            End If
            parametrosPesquisa = parametrosPesquisa & "Data_Transacao >= " & aspas(Format(txtDtTransacaoPesq.Text & " 00:00:00", "yyyymmdd hh:mm:ss"))
            parametrosPesquisa = parametrosPesquisa & "Data_Transacao <= " & aspas(Format(txtDtTransacaoPesq.Text & " 23:59:59", "yyyymmdd hh:mm:ss"))
        End If
        
        If Len(txtVlrTransacaoPesq.Text) > 0 Then
            If Len(Trim(parametrosPesquisa)) > 0 Then
                parametrosPesquisa = parametrosPesquisa & " and "
            End If
            parametrosPesquisa = parametrosPesquisa & "Valor_Transacao = " & SQLDecimal(txtVlrTransacaoPesq)
        End If
        
        If chkAprovada.Value = 1 Or chkPendente.Value = 1 Or chkCancelada.Value = 1 Then
            If Len(Trim(parametrosPesquisa)) > 0 Then
                parametrosPesquisa = parametrosPesquisa & " and "
            End If
            
            parametrosPesquisa = parametrosPesquisa & " ( "
            
            If chkAprovada.Value = 1 Then
                parametrosPesquisa = parametrosPesquisa & " Status_Transacao = 'Aprovada' "
                temStatus = True
            End If
            
            If chkPendente.Value = 1 Then
                parametrosPesquisa = parametrosPesquisa & " Status_Transacao = 'Pendente' "
                If temStatus = True Then
                    parametrosPesquisa = parametrosPesquisa & " or "
                End If
                
                temStatus = True
            End If
            
            If chkCancelada.Value = 1 Then
                parametrosPesquisa = parametrosPesquisa & " Status_Transacao = 'Cancelada' "
                If temStatus = True Then
                    parametrosPesquisa = parametrosPesquisa & " or "
                End If
                
                temStatus = True
            End If
            
            parametrosPesquisa = parametrosPesquisa & " ) "
        
        End If
        
        'montagem condicional com os parametros
        SQL = SQL & " where " & parametrosPesquisa
    
    End If
    
    SQL = SQL & " order by Numero_Cartao, Data_Transacao, Valor_Transacao"
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    
    Do While Not RS.EOF
        Set objItensLista = ListView1.ListItems.Add(, , RS!Id_Transacao)
        objItensLista.SubItems(1) = RS!Numero_Cartao
        objItensLista.SubItems(2) = Format(MyVal(RS!Valor_Transacao), "##,###,##0.00")
        objItensLista.SubItems(3) = Format(RS!Data_Transacao, "dd/MM/yyyy hh:mm:ss")
        objItensLista.SubItems(4) = RS!Status_Transacao
        objItensLista.SubItems(5) = RS!Descricao
        RS.MoveNext
    Loop
    
    'fechamento do recordset
    RS.Close
    Set RS = Nothing

    If ListView1.ListItems.Count = 0 Then
        MsgBox "Nenhum registro localizado com a Combinação informada", vbInformation, "MENSAGEM"
    End If

End Sub

Private Sub txtVlrTransacaoPesq_KeyPress(KeyAscii As Integer)
    
    'Tratamento de erro
    On Error Resume Next
    
    If KeyAscii <> 8 And KeyAscii <> 44 Then
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

End Sub

Private Sub txtVlrTransacaoPesq_GotFocus()

    'Tratamento de erro
    On Error Resume Next
    
    'verifica se pegou ou total ou seja ja tem virgula
    Dim virgula() As String
    Dim valor As String
    valor = txtVlrTransacaoPesq.Text
    
    virgula = Split(valor, ",")
    If UBound(virgula) = 0 Then
        'formata o valor sem decimal
        valor = MyVal(valor) / 100
    End If

    'formata o valor
    If Len(valor) > 0 Then
        txtVlrTransacaoPesq.Text = Format(MyVal(valor), "#######0.00")
    End If

End Sub

Private Sub txtVlrTransacaoPesq_LostFocus()

    'Tratamento de erro
    On Error Resume Next
    
    'verifica se pegou ou total ou seja ja tem virgula
    Dim virgula() As String
    virgula = Split(txtVlrTransacaoPesq.Text, ",")
    If UBound(virgula) = 0 Then
        'formata o valor
        txtVlrTransacaoPesq.Text = Format(MyVal(txtVlrTransacaoPesq.Text) / 100, "##,###,##0.00")
    Else
        'formata o valor
        txtVlrTransacaoPesq.Text = Format(MyVal(txtVlrTransacaoPesq.Text), "##,###,##0.00")
    End If
   
End Sub

