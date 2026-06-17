VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.ocx"
Begin VB.MDIForm frm00 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Grupo Adriano Cabuccio"
   ClientHeight    =   7845
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12330
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7515
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   9
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4128
            MinWidth        =   2646
            Text            =   "IP Local: 999.999.999.999"
            TextSave        =   "IP Local: 999.999.999.999"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4128
            MinWidth        =   2646
            Text            =   "ESTAÇAO:"
            TextSave        =   "ESTAÇAO:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   1659
            MinWidth        =   177
            Text            =   "USUARIO:"
            TextSave        =   "USUARIO:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2295
            MinWidth        =   2295
            Text            =   "Versão: 9.9.99"
            TextSave        =   "Versão: 9.9.99"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Ultimo Setup: "
            TextSave        =   "Ultimo Setup: "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Alignment       =   1
            Object.Width           =   741
            MinWidth        =   741
            TextSave        =   "INS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17/06/2026"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "10:59"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel ImagemLogo 
      Align           =   1  'Align Top
      Height          =   7530
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12330
      _Version        =   65536
      _ExtentX        =   21749
      _ExtentY        =   13282
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.Image Imagem 
         Height          =   2985
         Left            =   4560
         Picture         =   "DesafioAC00.frx":0000
         Top             =   1800
         Width           =   3015
      End
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "Cadastro"
      Begin VB.Menu DesafioAC01 
         Caption         =   "01 - Transações"
      End
      Begin VB.Menu DesafioAC02 
         Caption         =   "02 - Usuarios/Permissões"
      End
      Begin VB.Menu DesafioACSair 
         Caption         =   "00 - Sair"
      End
   End
   Begin VB.Menu mnuExportacao 
      Caption         =   "Exportação"
      Begin VB.Menu DesafioAC20 
         Caption         =   "20 - Exportação Excel"
      End
   End
   Begin VB.Menu mnuUtilitarios 
      Caption         =   "Utilitários"
      Begin VB.Menu DesafioAC99 
         Caption         =   "99 - Backup"
      End
   End
End
Attribute VB_Name = "frm00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DesafioAC01_Click()
    frm01.Show 1
End Sub

Private Sub DesafioAC02_Click()
    frm02.Show 1
End Sub

Private Sub DesafioAC20_Click()
    frm20.Show 1
End Sub

Private Sub DesafioAC99_Click()
    frm91.Show 1
End Sub

Private Sub DesafioACSair_Click()
    Unload Me
End Sub

Private Sub MDIForm_Load()

    Dim RS As New ADODB.Recordset, SQL As String
    Dim i As Integer, aux As String
    
     'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If

    'default desabilitado
    If UCase$(usuario_logado) <> "ADMIN" Then
    
        'pessimista
        DesafioAC01.Enabled = False
        DesafioAC02.Enabled = False
        DesafioAC20.Enabled = False
        DesafioAC99.Enabled = False
        
        'habilitar as permissoes
        SQL = "select p.CodTela from Permissao p inner join Usuario u"
        SQL = SQL & " on p.CodUsuario = u.CodUsuario"
        SQL = SQL & " where u.NomeUsuario = " & aspas(usuario_logado)
        
        RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
        Do While Not RS.EOF
        
            'Tela 01
            If Trim(RS!CodTela) = "01" Then
                DesafioAC01.Enabled = True
            End If
            
            'Tela 02
            If Trim(RS!CodTela) = "02" Then
                DesafioAC02.Enabled = True
            End If
            
            'Tela 20
            If Trim(RS!CodTela) = "20" Then
                DesafioAC20.Enabled = True
            End If
            
            'Tela 99
            If Trim(RS!CodTela) = "99" Then
                DesafioAC99.Enabled = True
            End If
            
            RS.MoveNext
        Loop
        RS.Close
        Set RS = Nothing
        
    End If

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'fechamento de database e matar a instancia
    If DB.State = adStateOpen Then
        DB.Close
    End If
    Set DB = Nothing
    
    'encerra o programa
    End

End Sub

Private Sub MDIForm_Resize()
    ImagemLogo.Visible = False
    
    'deixa o panel no tamanho do form
    If frm00.Height - 555 - 555 - 555 > 0 Then
        ImagemLogo.Height = frm00.Height - 555 - 555 - 555  'desconta a barra de ferramentas e status bar
        'centraliza a imagem dentro do panel
        Imagem.Left = (ImagemLogo.Width - Imagem.Width) / 2
        Imagem.Top = ((ImagemLogo.Height - Imagem.Height) / 2) - 555        'desconta a barra de ferramentas
        ImagemLogo.Visible = True
    End If
    
End Sub

