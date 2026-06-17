VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frm20 
   Caption         =   "20 - Exportação Excel"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand cmdSair 
      Height          =   555
      Left            =   3240
      TabIndex        =   7
      Top             =   840
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   979
      _StockProps     =   78
      ForeColor       =   -2147483640
      BevelWidth      =   1
      AutoSize        =   1
      Picture         =   "DesafioAC20.frx":0000
   End
   Begin Threed.SSCommand cmdExportacaoExcel 
      Height          =   555
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Exportar"
      ForeColor       =   -2147483640
      BevelWidth      =   1
      Picture         =   "DesafioAC20.frx":0A42
   End
   Begin VB.Frame relatorio1 
      Caption         =   "Data ...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4275
      Begin MSComCtl2.DTPicker txtDe2 
         Height          =   375
         Left            =   810
         TabIndex        =   1
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   99221507
         CurrentDate     =   42650
      End
      Begin MSMask.MaskEdBox txtDe 
         Height          =   375
         Left            =   810
         TabIndex        =   2
         Top             =   330
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker txtAte2 
         Height          =   375
         Left            =   780
         TabIndex        =   3
         Top             =   930
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   99221507
         CurrentDate     =   42650
      End
      Begin MSMask.MaskEdBox txtAte 
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Inicial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Final:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   5
         Top             =   990
         Width           =   555
      End
   End
   Begin FPSpreadADO.fpSpread SprdMultipla 
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   240
      Width           =   1575
      _Version        =   524288
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "DesafioAC20.frx":0C0F
   End
End
Attribute VB_Name = "frm20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdsair_click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   
   ' Troca a tecla "ENTER" por "TAB"
   MudaEnterToTab Me, KeyAscii

End Sub

Private Sub Form_Load()
    
    'defaults
    txtDe2.CustomFormat = "dd/MM/yyyy"
    txtAte2.CustomFormat = "dd/MM/yyyy"
    txtDe2.Value = DateSerial(Year(Now()), Month(Now()), 1)
    txtAte2.Value = DateSerial(Year(Now()), Month(Now()) + 1, 0)

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

Private Sub cmdExportacaoExcel_Click()
    
    'tratamento de erro
    'On Error GoTo CheckExcel
    
    'definir o recordset
    Dim RS As New ADODB.Recordset
    Dim SQL As String
    
    Dim k As Integer, i As Integer, j As Integer
    
    If MsgBox("Deseja Exportar para uma Planilha de Trabalho do EXCEL? Um volume de dados muito grande pode resultar em lentidão.", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
     'abre conexão se estiver fechado em algum momento
    If Not OpenDB() Then
        Exit Sub
    End If
    
    'primeira linha já está definido os cabeçalhos
    SprdMultipla.MaxRows = 1
    
    ''***para testes, colocar top 10 para processar rapidinho****
    SQL = "SELECT "      ''top 10 "
    SQL = SQL & " Id_Transacao,"
    SQL = SQL & " Numero_Cartao,"
    SQL = SQL & " Valor_Transacao,"
    SQL = SQL & " Data_Transacao,"
    SQL = SQL & " Descricao,"
    SQL = SQL & " Status_Transacao,"
    SQL = SQL & " dbo.funcCategoria(Valor_Transacao) as Categoria"
    SQL = SQL & " FROM      Transacoes"
    SQL = SQL & " WHERE     Data_Transacao >= " & aspas(Format(txtDe2.Value & " 00:00:00", "yyyymmdd hh:mm:ss"))
    SQL = SQL & "       AND Data_Transacao <= " & aspas(Format(txtAte2.Value & " 23:59:59", "yyyymmdd hh:mm:ss"))
    
    SQL = SQL & " order by Numero_Cartao, Data_Transacao, Valor_Transacao"
    RS.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
    
    Do While Not RS.EOF
    
        SprdMultipla.MaxRows = SprdMultipla.MaxRows + 1
        SprdMultipla.Row = SprdMultipla.MaxRows
        
        SprdMultipla.Col = 1
        SprdMultipla.Text = RS!Id_Transacao
        
        SprdMultipla.Col = 2
        SprdMultipla.Text = RS!Numero_Cartao
        
        SprdMultipla.Col = 3
        SprdMultipla.Text = Format(MyVal(RS!Valor_Transacao), "##,###,##0.00")
        
        SprdMultipla.Col = 4
        SprdMultipla.Text = Format(RS!Data_Transacao, "dd/mm/YYYY hh:mm:ss")
        
        SprdMultipla.Col = 5
        SprdMultipla.Text = IIf(IsNull(RS!Descricao), "", RS!Descricao)
        
        SprdMultipla.Col = 6
        SprdMultipla.Text = RS!Status_Transacao
        
        SprdMultipla.Col = 7
        SprdMultipla.Text = IIf(IsNull(RS!Categoria), "", RS!Categoria)
        
        RS.MoveNext
    Loop
    
    'fechamento do recordset
    RS.Close
    Set RS = Nothing
    
    If SprdMultipla.MaxRows = 0 Then
        MsgBox "Nenhum registro localizado!", vbInformation, "MENSAGEM"
        Exit Sub
    End If
    
    'dimensionar os objetos
    Dim exc As New Excel.Application
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet

    'adicionar dinamicamente os objetos
    Set wb = exc.Workbooks.Add
    Set ws = wb.Worksheets.Add
    ws.Cells.Clear
    
    'cabecalho
    SprdMultipla.Row = 0
    SprdMultipla.MaxCols = 8
    For i = 1 To SprdMultipla.MaxRows + 1
        SprdMultipla.Row = i
        For j = 1 To SprdMultipla.MaxCols
            SprdMultipla.Col = j
            If i = 1 Then
                ws.Cells(i, j).Font.Bold = True
            End If
            ws.Cells(i, j) = "'" & SprdMultipla.Text
        
            'ajustar as colunas automaticamente
            ws.Columns(j).AutoFit
        Next
        SprdMultipla.Row = i
    Next
    
    'ajusta manualmente o tamanho das coluna
    ws.Columns(1).ColumnWidth = 15                  'Id Transação
    ws.Columns(2).ColumnWidth = 20                  'Numero Cartão
    ws.Columns(3).ColumnWidth = 20                  'Valor Transação
    ws.Columns(4).ColumnWidth = 20                  'Data Transação
    ws.Columns(5).ColumnWidth = 60                  'Descrição
    ws.Columns(6).ColumnWidth = 20                  'Status Transação
    ws.Columns(7).ColumnWidth = 20                  'Categoria
    
    'ajusta manualmente o alinhamento das coluna
    ws.Columns(1).HorizontalAlignment = xlCenter    'Id Transação
    ws.Columns(2).HorizontalAlignment = xlLeft      'Numero Cartão
    ws.Columns(3).HorizontalAlignment = xlRight     'Valor Transação
    ws.Columns(4).HorizontalAlignment = xlCenter    'Data Transação
    ws.Columns(5).HorizontalAlignment = xlLeft      'Descrição
    ws.Columns(6).HorizontalAlignment = xlLeft      'Status Transação
    ws.Columns(7).HorizontalAlignment = xlLeft      'Categoria
    
    exc.Visible = True
    
    Set exc = Nothing
    Set wb = Nothing
    Set ws = Nothing
    
    Exit Sub
    
CheckExcel:
    MsgBox Err.Description

    'grava LogOperacao
    GravaLogOperacao "20", usuario_logado, SQL, DB

End Sub

