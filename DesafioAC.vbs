Dim PastaServidor
PastaServidor = "Z:"    'inserir o drive\path onde esta o GC.EXE no servidor

'***
'*** NAO MEXER DAQUI PARA BAIXO
'***

Dim NomeArquivo
NomeArquivo = "GC.EXE"

'dimensionar objetos do script
Dim wshell
set wshell = createobject("wscript.shell")

Dim fso
set fso = CreateObject("Scripting.FileSystemObject")

'dimensionar variaveis de trabalho
Dim copiar
Dim ArquivoServidor
Dim arqServ, arqL
dim DataHoraArquivoServidor, DataHoraArquivoLocal

'otimista - não vai precisar copiar
copiar = false

'verificar se existe o arquivo no servidor
ArquivoServidor = PastaServidor & "\" & NomeArquivo
if not fso.FileExists(ArquivoServidor) then
    msgbox "Arquivo no Servidor Não Localizado : " &   ArquivoServidor  & vbcrlf & "Contate Suporte Técnico"
    'encerra
    wscript.quit    
end if
'pega a datahora do arquivo no servidor
Set arqServ = fso.GetFile(ArquivoServidor)
DataHoraArquivoServidor = arqServ.DateLastModified

'verificar se existe o arquivo Local
if not fso.FileExists(NomeArquivo) then
    copiar = true
else
    'existe o arquivo - verificar a data hora
    Set arqL = fso.GetFile(NomeArquivo)    
    DataHoraArquivoLocal = arqL.DateLastModified
    'verifica se a data arquivo local é mais velha
    if cdate(DataHoraArquivoLocal) < DataHoraArquivoServidor then
        copiar = true
    end if
end if 

'se precisa copiar entao copia
if  copiar then
    msgbox "Atenção! A Cópia Local do Sistema Será Atualizada " & vbcrlf & vbcrlf & " -> Data/Hora do Arquivo do Servidor: " &  DataHoraArquivoServidor & vbcrlf & " -> Data/Hora do Arquivo Local: " &  DataHoraArquivoLocal
    fso.CopyFile ArquivoServidor, NomeArquivo
end if

'*** executar
set wshell = createobject("wscript.shell")
wshell.run   NomeArquivo  
wscript.quit

'*** documentacao referencia
'msgbox "Tamanho  		 = " & FormatNumber(arq.Size, 0)
'msgbox "Data de criação  = " & arq.DateCreated
'msgbox "Último acesso	 = " & arq.DateLastAccessed
'msgbox "Última alteração = " & arq.DateLastModified
'msgbox "nome abreviado	 = " & arq.ShortName

'pegar propriedades de um DIRETORIO/PASTA
'dim pasta
'Set pasta = fso.GetDrive("C:")
'If pasta.IsReady Then
'	msgbox "Letra associada ao drive : " & pasta.DriveLetter
'  	msgbox "Nome do Volume : " & pasta.VolumeName
'  	msgbox "Espaço total do drive : " & FormatNumber(pasta.TotalSize / 1024, 0)
'  	msgbox "Espaço disponível livre no drive : " & FormatNumber(pasta.FreeSpace / 1024, 0)
'   msgbox "Número de série do drive : " & pasta.SerialNumber
'   msgbox "Tipo do drive: " & pasta.DriveType
'end if