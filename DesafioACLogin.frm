VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Logon no Sistema"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5370
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   1530
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtSenha 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1860
      Width           =   2535
   End
   Begin VB.TextBox txtUsuario 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   1500
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   180
      Picture         =   "DesafioACLogin.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   1860
      Width           =   480
   End
   Begin VB.PictureBox pctImage 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      Picture         =   "DesafioACLogin.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   5325
      TabIndex        =   2
      Top             =   0
      Width           =   5325
   End
   Begin VB.TextBox txtServidor 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   780
      Width           =   2535
   End
   Begin VB.TextBox txtDatabase 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1140
      Width           =   2535
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Usuário:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblSenha 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Senha:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   1920
      Width           =   1515
   End
   Begin VB.Label lblServidor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Servidor:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   840
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Nome Database:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   1200
      Width           =   1515
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Erros As String
Dim tam As Long
Dim RS As Recordset
Dim SenhaDigitada As String
Dim Servidores() As String
Dim Databases() As String
Dim Fantasias() As String
Dim ComandoExterno As String
Dim Comandos() As String

Private Sub Form_Activate()
     On Error Resume Next
    
    txtUsuario.SetFocus
End Sub

Sub Form_Load()
    
    'tratamento de erro
    On Error Resume Next
    
    Dim offset As Long
    If DB.State = adStateOpen Then
        DB.Close
    End If
        
    'Carregar campos
    txtServidor.Text = Servidor
    txtDatabase.Text = Database
    txtUsuario.Text = ""
    txtSenha.Text = ""

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
    
    If KeyAscii = 27 Then
        Call cmdCancelar_Click
    End If

End Sub

Private Sub cmdCancelar_Click()

    'sair sistema
    End

End Sub

Private Sub cmdOK_Click()

    'tratamento de erro
    On Error Resume Next
    
    'definir o recordset
    Dim RS As New ADODB.Recordset, SQL As String
    Dim msg As String
    
     'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If
    
    'verifica se é um OK de confirmacao de Senha
    If lblSenha.Caption = "Conf.Senha:" Then
    
        If Trim(txtSenha.Text) <> Trim(SenhaDigitada) Then
            MsgBox "Senha Não Confere!"
            txtUsuario.Text = ""
            txtUsuario.SetFocus
            Exit Sub
        End If
        If Len(Trim(txtSenha.Text)) > 10 Or Len(Trim(SenhaDigitada)) > 10 Then
            MsgBox "Tamanho da Seha não pode ser MAIOR que 10 Caracteres (Letras/Numeros)"
            txtUsuario.Text = ""
            txtUsuario.SetFocus
            Exit Sub
        End If
        
        'grava a senha
        SQL = "update usuario set "
        SQL = SQL & "SenhaUsuario = " & aspas(FormatCripto(UCase(Trim(txtSenha.Text)), True))
        SQL = SQL & " where NomeUsuario = " & aspas(txtUsuario.Text)
        DB.Execute SQL, , adExecuteNoRecords
        
        'sai do sistema
        MsgBox "Senha Alterada! Entre Novamente na Aplicação!"
        DoEvents
        
        'sai do programa
        End
    
    End If
    
    Dim strcurrency As String, MYCURRENCY As String
    MYCURRENCY = FormatCurrency(1000)
    strcurrency = Format(MYCURRENCY, "0,000.00")
    
    If Len(Trim(txtUsuario)) = 0 Then
        Beep
        MsgBox "Usuário não Informado", vbCritical, "Usuário"
        txtUsuario.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtSenha)) = 0 Then
        Beep
        MsgBox "Senha não Informada", vbCritical, "Password"
        txtSenha.SetFocus
        Exit Sub
    End If
        
    'VERIFICAR O REGIONAL SETTINGS
    'Dim length As Long, Formato As String
    'Dim buf As String * 1024
    'lLocal = GetSystemDefaultLCID()
    'length = GetLocaleInfo(lLocal, LOCALE_SSHORTDATE, buf, Len(buf))
    'Formato = Left$(buf, length - 1)
    'If UCase(Formato) <> "DD/MM/YYYY" And UCase(Formato) <> "DD/MM/AAAA" Then
    '    msg = "As Opções de Data no Windows no 'Painel de Controle / Regional Settings' estão Inválidas."
    '    msg = msg & vbCrLf & "O Formato Atual Configurado é: " & Formato
    '    msg = msg & vbCrLf & "O Formato deve ser: dd/MM/aaaa (ou dd/MM/yyyy se seu Windows é Ingles)"
    '    MsgBox msg, vbCritical, "MENSAGEM"
        'MsgBox "O Programa Será Encerrado!", vbCritical, "MENSAGEM"
        'End
    'End If
      
    '*** INICIAR NORMALMENTE
    Me.MousePointer = 11
    
    'pegar o nome da workstation e usuario
    usuario_logado = txtUsuario.Text
    
    'verifica se o usuario e senha OK
    If UCase(txtUsuario.Text) <> "ADMIN" Then
        SQL = "select * from usuario where upper(NomeUsuario) = " & aspas(UCase(txtUsuario.Text))
        RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
        Me.MousePointer = 0
        If RS.EOF Then
            MsgBox "Usuario Nao existe"
            RS.Close
            Set RS = Nothing
            DB.Close
            Set DB = Nothing
            txtUsuario.SetFocus
            Exit Sub
        End If
        If IIf(IsNull(RS!SenhaUsuario), "", RS!SenhaUsuario) = "*RESETPWD*" Then
            'confirma a senha
            lblSenha.Caption = "Conf.Senha:"
            If Len(Trim(SenhaDigitada)) = 0 Then
                SenhaDigitada = Trim(txtSenha.Text)
            End If
            txtSenha.Text = ""
            txtSenha.SetFocus
            Exit Sub
        End If
        If UCase(Trim(RS!SenhaUsuario)) <> UCase(FormatCripto(UCase(Trim(txtSenha.Text)), True)) Then
            MsgBox "Senha Invalida"
            RS.Close
            Set RS = Nothing
            DB.Close
            Set DB = Nothing
            Exit Sub
        End If
        If IIf(IsNull(RS!AtivoUsuario), "N", RS!AtivoUsuario) = "N" Then
            MsgBox "Conta do Usuário Não Esta Ativa!"
            RS.Close
            Set RS = Nothing
            DB.Close
            Set DB = Nothing
            Exit Sub
        End If
        
        'fecha o recordset
        RS.Close
        Set RS = Nothing
        
        'oculta o logon
        Me.Visible = False
        
        'abre o menu
        frm00.Show
        Exit Sub
    
    Else
        Me.MousePointer = 0
        If UCase(txtUsuario.Text) = "ADMIN" And UCase(txtSenha.Text) <> UCase(SenhaAdmin) Then
            MsgBox "Senha Administrador Invalida"
            DB.Close
            Set DB = Nothing
        Else
            'oculta o logon
            Me.Visible = False
            
            'abre o menu
            frm00.Show
        
        End If
    
    End If

End Sub

Private Sub txtSenha_GotFocus()
    txtSenha.SelStart = 0
    txtSenha.SelLength = Len(txtSenha.Text)
End Sub

Private Sub txtUsuario_GotFocus()
    txtUsuario.SelStart = 0
    txtUsuario.SelLength = Len(txtUsuario.Text)
    txtSenha.Text = ""
    SenhaDigitada = ""
    lblSenha.Caption = "Senha:"
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtUsuario_LostFocus()
    txtUsuario.Text = UCase(txtUsuario.Text)
End Sub


