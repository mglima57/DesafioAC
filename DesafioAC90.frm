VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.ocx"
Begin VB.Form frm90 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DesafioAC90 - Identificação do Sistema"
   ClientHeight    =   2835
   ClientLeft      =   3285
   ClientTop       =   3630
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2835
   ScaleMode       =   0  'User
   ScaleWidth      =   4391.483
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "SQL Central "
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5895
      Begin VB.TextBox txtSenha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1740
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "viceri@"
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtUsuario 
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "desafioAC"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtPortaIP 
         Height          =   315
         Left            =   5040
         MaxLength       =   5
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtDatabase 
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "DesafioAC"
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtServidor 
         Height          =   315
         Left            =   1740
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuário SQL:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1380
         Width           =   1515
      End
      Begin VB.Label lblSenha 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Senha:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porta IP:"
         Height          =   195
         Left            =   4320
         TabIndex        =   7
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nome Database:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   900
         Width           =   1515
      End
      Begin VB.Label lblServidor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Servidor:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   1515
      End
   End
   Begin Threed.SSPanel frmToolBar 
      Height          =   555
      Left            =   0
      TabIndex        =   11
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
         TabIndex        =   12
         Top             =   90
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         AutoSize        =   1
         Picture         =   "DesafioAC90.frx":0000
      End
      Begin Threed.SSCommand cmdSalvar 
         Height          =   375
         Left            =   90
         TabIndex        =   13
         Top             =   90
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483640
         BevelWidth      =   1
         AutoSize        =   1
         Picture         =   "DesafioAC90.frx":0A42
      End
   End
End
Attribute VB_Name = "frm90"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
     On Error Resume Next
    
    txtServidor.SetFocus
    
End Sub

Sub Form_Load()

    'tratamento de erro
    On Error Resume Next

    'Inicia a Tela
    Call IniciaTela
    
End Sub
    
Private Sub IniciaTela()

    'Ler do Arquivo DesafioAC os parametros para Inicialização do banco de dados
    IniFile = App.Path & "\DesafioAC.ini"
    If Len(Dir(IniFile)) <> 0 Then
        
        'pegar o nome do servidor
        txtServidor.Text = FormatCripto(Trim(ReadIniFile(IniFile, "IDENTIFICACAO", "ServerSQL")), False)
            
        'pegar a porta
        txtPortaIP.Text = FormatCripto(Trim(ReadIniFile(IniFile, "IDENTIFICACAO", "PortaIP")), False)
        
        'pegar o nome da base
        txtDatabase.Text = FormatCripto(Trim(ReadIniFile(IniFile, "IDENTIFICACAO", "Database")), False)
       
        'pegar o nome do usuario
        txtUsuario.Text = FormatCripto(Trim(ReadIniFile(IniFile, "IDENTIFICACAO", "Usuario")), False)
       
        'pegar a senha
        txtSenha.Text = FormatCripto(Trim(ReadIniFile(IniFile, "IDENTIFICACAO", "Senha")), False)
    
    Else
    
        txtServidor.Text = ""
        txtPortaIP.Text = ""
        
    End If
    
End Sub

Private Sub cmdsair_click()

    'sair da tela
    Unload Me
    Exit Sub

End Sub

Private Sub cmdSalvar_Click()

    'tratamento de erro
    On Error Resume Next
    
    'verifica a consistencia
    If Not consistente() Then
        Exit Sub
    End If
    
    'testa a concexão do so Servidor e o DataBase
    DBTeste.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;LANGUAGE=us_english;"
    DBTeste.ConnectionString = DBTeste.ConnectionString & "User ID=" & Trim(txtUsuario.Text) & ";"
    DBTeste.ConnectionString = DBTeste.ConnectionString & "Password=" & Trim(txtSenha.Text) & ";"
    DBTeste.ConnectionString = DBTeste.ConnectionString & "Initial Catalog=" & Trim(txtDatabase.Text) & ";"
    DBTeste.ConnectionString = DBTeste.ConnectionString & "Data Source=" & Trim(txtServidor.Text) & "," & Trim(txtPortaIP.Text) & ";"
    DBTeste.ConnectionString = DBTeste.ConnectionString & "Connect Timeout=2"
    DBTeste.Open
    If DBTeste.State <> adStateOpen Then
        MsgBox "Não foi possível estabelecer conexão com:" & vbCrLf & vbCrLf & vbCrLf & _
        "Servidor  : " & Trim(txtServidor.Text) & vbCrLf & vbCrLf & _
        "DataBase  : " & Trim(txtDatabase.Text) & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
        "Favor verificar as informações digitadas.", vbCritical, "Erro Validação"
        
        Exit Sub
    End If
    
    'fecha conexão
    If DBTeste.State = adStateOpen Then
        DBTeste.Close
    End If
        
    'grava ini
    IniFile = App.Path & "\DesafioAC.ini"
    
    '*** Database
    WriteIniFile IniFile, "IDENTIFICACAO", "ServerSQL", FormatCripto(txtServidor.Text, True)
    WriteIniFile IniFile, "IDENTIFICACAO", "PortaIP", FormatCripto(txtPortaIP.Text, True)
    WriteIniFile IniFile, "IDENTIFICACAO", "Database", FormatCripto(txtDatabase.Text, True)
    WriteIniFile IniFile, "IDENTIFICACAO", "Usuario", FormatCripto(txtUsuario.Text, True)
    WriteIniFile IniFile, "IDENTIFICACAO", "Senha", FormatCripto(txtSenha.Text, True)
 
    MsgBox "Validação executada com Sucesso!", vbInformation, "Validação"
    
    'sair da tela
    Unload Me
    Exit Sub
    
End Sub

Private Function consistente() As Boolean

    'otimista
    consistente = True

    'verificações preliminar
    If Len(Trim(txtServidor.Text)) = 0 Then
        MsgBox "Servidor não Informado", vbCritical, "Alerta"
        consistente = False
        
        'retorno o foco
        txtServidor.SetFocus
            
        Exit Function
        
    End If
    
    If Len(Trim(txtPortaIP.Text)) = 0 Then
        MsgBox "Porta IP não Informada", vbCritical, "Alerta"
        consistente = False
        
        'retorno o foco
        txtPortaIP.SetFocus
            
        Exit Function
        
    End If
    
    If Len(Trim(txtDatabase.Text)) = 0 Then
        MsgBox "Nome da Database não Informado", vbCritical, "Alerta"
        consistente = False
        
        'retorno o foco
        txtDatabase.SetFocus
            
        Exit Function
        
    End If
    
    If Len(Trim(txtUsuario.Text)) = 0 Then
        MsgBox "Usuario da Database não Informado", vbCritical, "Alerta"
        consistente = False
        
        'retorno o foco
        txtUsuario.SetFocus
            
        Exit Function
        
    End If
    
    If Len(Trim(txtSenha.Text)) = 0 Then
        MsgBox "Senda do Usuario da Database não Informado", vbCritical, "Alerta"
        consistente = False
        
        'retorno o foco
        txtSenha.SetFocus
            
        Exit Function
        
    End If
    
End Function
