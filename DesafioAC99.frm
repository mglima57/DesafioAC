VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm99 
   Caption         =   "99 - Pesquisa"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   7011
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
   Begin Threed.SSCommand cmdSair 
      Height          =   495
      Left            =   9960
      TabIndex        =   1
      ToolTipText     =   "Sair Tela"
      Top             =   0
      Width           =   495
      _Version        =   65536
      _ExtentX        =   882
      _ExtentY        =   882
      _StockProps     =   78
      ForeColor       =   -2147483640
      BevelWidth      =   1
      AutoSize        =   1
      Picture         =   "DesafioAC99.frx":0000
   End
   Begin Threed.SSCommand cmdRefresh 
      Height          =   495
      Left            =   9960
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Refresh Tela"
      Top             =   600
      Width           =   495
      _Version        =   65536
      _ExtentX        =   882
      _ExtentY        =   882
      _StockProps     =   78
      ForeColor       =   -2147483640
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "DesafioAC99.frx":0A42
   End
   Begin Threed.SSCommand cmdConfirma 
      Height          =   495
      Left            =   9960
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Aplicar Seleção"
      Top             =   1200
      Width           =   495
      _Version        =   65536
      _ExtentX        =   882
      _ExtentY        =   882
      _StockProps     =   78
      ForeColor       =   -2147483640
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "DesafioAC99.frx":131C
   End
End
Attribute VB_Name = "frm99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
'definir o recordset
Dim RS As New ADODB.Recordset
Dim SQL As String
Dim i As Integer

'chaves primarias
Dim CodCorretor As Integer
Dim CodCliente As Integer
Dim CodUF As Integer
Dim CodCidade As Integer

Public TituloJanela As String

Private Sub cmdsair_click()
    Unload Me
End Sub

Private Sub cmdConfirma_Click()

    'tratamento de erro
    On Error Resume Next
    
    'dimensionamento/criaçao do recordset
    Dim RS As New ADODB.Recordset, SQL As String
    
    'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If
    
    'consistencias
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Não Existe Registros Na Lista"
        Exit Sub
    End If
    If Len(Trim((ListView1.ListItems(ListView1.SelectedItem.Index).Text))) = 0 Then
        MsgBox "Registro Nao Selecionado"
        Exit Sub
    End If

    If QualTela = "CR" Then
        frmQualTela.txtCodCorretor.Text = Trim(ListView1.ListItems(ListView1.SelectedItem.Index).Text)
    
    ElseIf QualTela = "CC" Then
        frmQualTela.txtCodCliente.Text = Trim((ListView1.ListItems(ListView1.SelectedItem.Index).Text))
    
    End If
    
    'sai do form
    Unload Me

End Sub

Private Sub cmdRefresh_Click()
    If QualTela <> "" Then
        Call Form_Activate
        
        'colocar o foco
        ListView1.SetFocus
    
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmQualTela = Nothing
    QualTela = ""
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

Private Sub Form_Activate()

    'se a vida em caso de erro
    On Error Resume Next
    
    'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If
    
    'titulo da janela
    Me.Caption = "DesafioAC99 - Pesquisa"
    If Len(Trim(TituloJanela)) <> 0 Then
        Me.Caption = Me.Caption & "  (" & TituloJanela & ")"
    End If

    'parametrizar a ListView
    ListView1.HideSelection = False
    ListView1.View = lvwReport
    ListView1.FullRowSelect = True
    ListView1.LabelEdit = lvwManual
    Dim objItensLista As MSComctlLib.ListItem
    
    'limpar a listview
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.GridLines = True
   
    If QualTela = "CR" Then
        
        'Corretores
        ListView1.ColumnHeaders.Add , "CodCorretor", "CodCorretor"
        ListView1.ColumnHeaders.Add , "NomeCorretor", "NomeCorretor"
        ListView1.ColumnHeaders.Add , "CPFCorretor", "CPFCorretor"
        
        'colocar o tamanho das colunas compativel
        ListView1.ColumnHeaders(1).Width = 1500
        ListView1.ColumnHeaders(2).Width = 4000
        ListView1.ColumnHeaders(3).Width = 2000
        
        'alinhamento
        ListView1.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
        
        'abertura do recordset
        SQL = "select * from Corretor"
        RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
        Do While Not RS.EOF
            Set objItensLista = ListView1.ListItems.Add(, , RS!CodCorretor)
            objItensLista.SubItems(1) = Trim(RS!NomeCorretor)
            objItensLista.SubItems(2) = Trim(RS!CPFCorretor)
            RS.MoveNext
        Loop

    ElseIf QualTela = "CC" Then
        
        'Clientes
        ListView1.ColumnHeaders.Add , "CodCliente", "CodCliente"
        ListView1.ColumnHeaders.Add , "NomeCliente", "NomeCliente"
        ListView1.ColumnHeaders.Add , "CPFCliente", "CPFCliente"
        ListView1.ColumnHeaders.Add , "EnderecoCliente", "EnderecoCliente"
        ListView1.ColumnHeaders.Add , "AtivoCliente", "AtivoCliente"
        
        'colocar o tamanho das colunas compativel
        ListView1.ColumnHeaders(1).Width = 1500
        ListView1.ColumnHeaders(2).Width = 4000
        ListView1.ColumnHeaders(3).Width = 2000
        ListView1.ColumnHeaders(4).Width = 4000
        ListView1.ColumnHeaders(5).Width = 1500
        
        'alinhamento
        ListView1.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
        
        'abertura do recordset
        SQL = "select * from Cliente"
        RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
        Do While Not RS.EOF
            Set objItensLista = ListView1.ListItems.Add(, , RS!CodCliente)
            objItensLista.SubItems(1) = RS!NomeCliente
            objItensLista.SubItems(2) = RS!CPFCliente
            objItensLista.SubItems(3) = RS!EnderecoCliente
            objItensLista.SubItems(4) = RS!AtivoCliente
            RS.MoveNext
        Loop
        
    End If

    'fechamento do recordset
    If RS.State = adStateOpen Then
        RS.Close
    End If
    Set RS = Nothing

End Sub

