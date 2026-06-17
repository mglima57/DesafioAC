VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.ocx"
Begin VB.Form frm91 
   Caption         =   "Back-UP"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame cmdFechar 
      Height          =   4695
      Left            =   90
      TabIndex        =   0
      Top             =   600
      Width           =   5895
      Begin VB.Frame Frame1 
         Caption         =   "Nome do Arquivo que será Gerado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   13
         Top             =   3210
         Width           =   5655
         Begin VB.TextBox txtArquivo 
            BackColor       =   &H8000000F&
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
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   14
            Top             =   240
            Width           =   5355
         End
      End
      Begin VB.Frame frame2 
         Caption         =   "Progresso"
         Height          =   585
         Left            =   120
         TabIndex        =   11
         Top             =   3930
         Width           =   5655
         Begin VB.Label lblMensagem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Executando a Re-criação de Índices ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   5235
         End
      End
      Begin VB.Frame frmDescricao 
         Caption         =   "Caminho Onde Será Gerado o Arquivo Backup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   8
         Top             =   2460
         Width           =   5655
         Begin VB.TextBox txtDisk 
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
            MaxLength       =   50
            TabIndex        =   10
            Top             =   240
            Width           =   4935
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   375
            Left            =   5160
            TabIndex        =   9
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame frmOpcoes 
         Caption         =   "Opções Adicionais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   90
         TabIndex        =   3
         Top             =   900
         Width           =   5655
         Begin Threed.SSCheck chkDB 
            Height          =   195
            Left            =   800
            TabIndex        =   4
            Top             =   600
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "CHECKDB"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkIndices 
            Height          =   195
            Left            =   800
            TabIndex        =   5
            Top             =   900
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Re-criação de índices"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkUpdate 
            Height          =   195
            Left            =   800
            TabIndex        =   6
            Top             =   1200
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Atualização Estatísticas"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkDefrag 
            Height          =   195
            Left            =   795
            TabIndex        =   7
            Top             =   300
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Desfragmentar Banco de Dados"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Banco de Dados para Back-UP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   90
         TabIndex        =   1
         Top             =   150
         Width           =   5655
         Begin VB.TextBox txtDatabase 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   5415
         End
      End
   End
   Begin Threed.SSPanel frmToolBar 
      Height          =   555
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6105
      _Version        =   65536
      _ExtentX        =   10769
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
         Left            =   5460
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
         Picture         =   "DesafioAC91.frx":0000
      End
      Begin Threed.SSCommand cmdProcessar 
         Height          =   375
         Left            =   90
         TabIndex        =   17
         Top             =   90
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "DesafioAC91.frx":0A42
      End
   End
End
Attribute VB_Name = "frm91"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private WithEvents oBackup As SQLDMO.Backup

Private Sub cmdBrowse_Click()
    Dim Form As New frmPath
    Form.PathDefault = txtDisk.Text
    Form.Show 1
    txtDisk.Text = Form.PathSelecionado
End Sub

Private Sub cmdProcessar_Click()

    'tratamento de erro
    On Error Resume Next
    
    If Not chkIndices Then
        Dim msg As String
        msg = "Não Foi Selecionado a Opção para Recriar os Índices para a manutenção dos níveis de" & vbCrLf
        msg = msg & "Performance e Eficiência. Esta opção deve ser executada pelo menos" & vbCrLf
        msg = msg & "duas vezes na semana! Continuar a Execução ?"
        If MsgBox(msg, vbInformation + vbYesNo, "ATENÇÃO!") = vbNo Then
            Exit Sub
        End If
    End If
        
    Me.MousePointer = 11
    Me.Refresh
    DoEvents
    
    'fecha preventivamente
    If DB.State = adStateOpen Then
        DB.Close
    End If
    
    DB2.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;LANGUAGE=us_english;"
    DB2.ConnectionString = DB2.ConnectionString & "User ID=" & Trim(Usuario) & ";"
    DB2.ConnectionString = DB2.ConnectionString & "Password=" & Trim(Senha) & ";"
    DB2.ConnectionString = DB2.ConnectionString & "Initial Catalog=" & Trim(Database) & ";"
    DB2.ConnectionString = DB2.ConnectionString & "Data Source=" & Trim(Servidor) & "," & Trim(PortaIP) & ";"
    DB2.ConnectionString = DB2.ConnectionString & "Connect Timeout=2"
    DB2.Open
    
    Dim SQL As String
    lblMensagem.Visible = True
    lblMensagem.Caption = "Executando Backup FULL ..."
    DoEvents
        
    'não faz compressão
    SQL = "BACKUP DATABASE [" & Database + "] TO disk = '" & txtDisk.Text & "\" & txtArquivo.Text & "' with FORMAT, INIT, NAME = '" & Database & "', DESCRIPTION = 'Executado pela Tela 99'"

    'executar
    DB2.CommandTimeout = 1800     '30 minutos = 1800 segundos
    DB2.Execute SQL, , adExecuteNoRecords

    If chkIndices Then
        lblMensagem.Visible = True
        lblMensagem.Caption = "Executando a Re-criação de Índices ..."
        DoEvents
        
        SQL = "exec PrMyGasCheckDB @funcao = 'I'"
        DB2.CommandTimeout = 900     '15 minutos = 900 segundos
        DB2.Execute SQL, , adExecuteNoRecords
        
        Me.MousePointer = 0
        lblMensagem.Caption = ""
        DoEvents
        
    End If
    If chkUpdate Then
        lblMensagem.Visible = True
        lblMensagem.Caption = "Atualizando as Estatísticas dos Índices ..."
        DoEvents
        
        SQL = "exec PrMyGasCheckDB @funcao = 'U'"
        DB2.CommandTimeout = 900     '15 minutos = 900 segundos
        DB2.Execute SQL, , adExecuteNoRecords
        
        Me.MousePointer = 0
        lblMensagem.Caption = ""
        DoEvents
        
    End If
    If chkDB Then
        lblMensagem.Visible = True
        lblMensagem.Caption = "Executando DBCC Check-DB ..."
        DoEvents
        
        SQL = "exec PrMyGasCheckDB @funcao = 'C'"
        DB2.CommandTimeout = 900     '15 minutos = 900 segundos
        DB2.Execute SQL, , adExecuteNoRecords
        
        Me.MousePointer = 0
        lblMensagem.Caption = ""
        DoEvents
        
    End If
    
    If chkDefrag Then
        lblMensagem.Visible = True
        lblMensagem.Caption = "Executando Desfragmentação ..."
        DoEvents
        
        SQL = "exec PrDBADesfragmentar @funcao = 'A', @minfrag = 0, @maxfrag = 100"
        DB2.CommandTimeout = 1800     '15 minutos = 900 segundos
        DB2.Execute SQL, , adExecuteNoRecords
        
        Me.MousePointer = 0
        lblMensagem.Caption = ""
        DoEvents
        
    End If
    
    'fechar a conexao secundaria da base
    If DB2.State = adStateOpen Then
        DB2.Close
    End If
    
    lblMensagem.Visible = False
    DoEvents
    
    MsgBox "Operação Concluída!"
        
    'sai do form
    Unload Me
End Sub

Friend Sub Form_Load()

    'nome da base
    txtDatabase.Text = Database

    lblMensagem.Visible = False
    lblMensagem.Left = 180
    lblMensagem.Top = 300
    
    'trava para nao editar
    txtDisk.Locked = True
    txtArquivo.Locked = True
    
    'Path de Backup
    txtDisk.Text = App.Path
    
    'nome do Arquivo DEFAULT
    txtArquivo.Text = "DesafioAC" & Format(Now(), "yyMMddHHmm") & ".BAK"
        
    'ajusta as opções
    chkDB.Enabled = True
    chkDefrag.Value = True
    chkDB.Value = True
    chkIndices.Value = True
    chkUpdate.Value = True

End Sub

Private Sub cmdsair_click()

    'sair da tela
    Unload Me
    Exit Sub
End Sub

