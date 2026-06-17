VERSION 5.00
Begin VB.Form frmPath 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GCPath - Selecionar Pasta/Diretorio"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   2340
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmArquivos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Duplo-Click no Arquivo Para Seleção"
      Height          =   3375
      Left            =   60
      TabIndex        =   7
      Top             =   4080
      Width           =   5775
      Begin VB.ListBox lstArquivos 
         Height          =   2985
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame frmOpen 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancela"
         Height          =   495
         Left            =   3480
         TabIndex        =   4
         Top             =   420
         Width           =   1035
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   495
         Left            =   4620
         TabIndex        =   3
         Top             =   420
         Width           =   1035
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   180
         TabIndex        =   2
         Top             =   1200
         Width           =   5475
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   3135
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Pastas/Diretórios:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Drives:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PathDefault As String
Public PathSelecionado As String
Public MostrarArquivos As Boolean
Public FiltroExtensao As String
Public ArquivoSelecionado As String
Public ArquivoComPathSelecionado

Private Sub Dir1_Change()

    'tratamento de erro
    On Error Resume Next
    
    Dim i As Integer, j As Integer
    Dim Extensao() As String
    
    Dim TiposExtensao() As String
    
    'Filtro de Extensão separado por ";" para multiplos tipos de extensão
    TiposExtensao = Split(FiltroExtensao, ";")
    
    File1.Path = Dir1.Path
    File1.Refresh
    lstArquivos.Clear
    For i = 0 To File1.ListCount - 1
        'filtro
        If Len(Trim(FiltroExtensao)) <> 0 Then
            'tem filtro especificado, verifica se atende
            Extensao = Split(File1.List(i), ".")
            If UBound(Extensao) = 0 Then
                GoTo Proximo
            End If
            
'            'Processo Antigo
'            If Len(Trim(Extensao(1))) <> Len(Trim(FiltroExtensao)) Then
'                GoTo Proximo
'            End If
'            If UCase(Trim(Extensao(1))) <> UCase(Trim(FiltroExtensao)) Then
'                GoTo Proximo
'            End If
            
            'procura nos tipos de extensao
            For j = 0 To UBound(TiposExtensao)
            
                If UCase(Trim(Extensao(1))) = UCase(Trim(TiposExtensao(j))) Then
                    GoTo Achou
                End If
            Next
            
            'não achou
            GoTo Proximo
        End If
Achou:
        lstArquivos.AddItem File1.List(i)
Proximo:
    Next
    'limpa preventivamente
    ArquivoSelecionado = ""
End Sub
Private Sub Form_Load()
    'tamanho padrao
    Me.Height = 4400
    'se for mostrar arquivo ajusta
    If MostrarArquivos Then
        Me.Height = 7900
    End If
    
    On Error Resume Next
    Err.Clear
    Drive1.Drive = Left(PathDefault, 2)

    If Len(Trim(PathDefault)) <> 0 Then
        Dir1.Path = PathDefault
    End If
End Sub
Private Sub cmdCancel_Click()
    PathSelecionado = ""
    ArquivoSelecionado = ""
    'sai do form
    Unload Me
End Sub
Private Sub cmdOK_Click()
    
    'tratamento de erro
    On Error Resume Next
    
    If Len(Trim(Dir1.Path)) = 0 Then
        MsgBox "Pasta/Diretório Não Selecionada!", vbInformation, "MENSAGEM"
        Exit Sub
    End If
    PathSelecionado = UCase(Trim(Dir1.Path))
    'sai do form
    Unload Me
End Sub
Private Sub Drive1_Change()
    
    'tratamento de erro
    On Error GoTo Error
    
    Dir1.Path = Drive1.Drive
    
    Exit Sub
    
Error:
    MsgBox Err.Description
    
    Drive1.Drive = Left(App.Path, 2)

End Sub
Private Sub lstArquivos_Click()
    
    'tratamento de erro
    On Error Resume Next
    
    ArquivoComPathSelecionado = Dir1.Path & "\" & lstArquivos.Text
    ArquivoSelecionado = lstArquivos.Text
    
    'remove "\\"
    ArquivoComPathSelecionado = Replace(ArquivoComPathSelecionado, "\\", "\")
    ArquivoSelecionado = Replace(ArquivoSelecionado, "\\", "\")
    
End Sub

Private Sub lstArquivos_DblClick()
    Call cmdOK_Click
End Sub
