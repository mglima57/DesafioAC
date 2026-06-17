-- criação do banco de dados
CREATE DATABASE DESAFIOAC
GO
-- crição de novo usuario no SQL
CREATE LOGIN DesafioAC WITH PASSWORD='DesafioAC@', CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF;
GO
-- acessando novo banco
USE DESAFIOAC
GO
-- apaga usuario peventivamente
DROP USER DesafioAC
GO
-- criando usuario e permissaão no novo banco
CREATE USER DesafioAC FOR LOGIN DesafioAC
GO
-- criando permissao de OWNER para o usuario
EXEC sp_addrolemember N'db_owner', N'DesafioAC'
GO
-- crianda a tabela Transacoes
CREATE TABLE Transacoes(
	Id_Transacao		INT IDENTITY(1,1)	, 
	Numero_Cartao	    CHAR(16) NOT NULL	,
    Valor_Transacao     DECIMAL(10, 2) NOT NULL,
    Data_Transacao      DATETIME DEFAULT CURRENT_TIMESTAMP NOT NULL,
    Descricao           VARCHAR(255),
    Status_Transacao    VARCHAR(10) NOT NULL,
    CONSTRAINT CHK_Valor_Positivo CHECK (Valor_Transacao > 0), 										
    CONSTRAINT CHK_Status_Permitido CHECK (Status_Transacao IN ('Aprovada', 'Pendente', 'Cancelada'))
	)
GO
-- criando tabela Usuarios
CREATE TABLE Usuario(
	CodUsuario		INT IDENTITY(1,1)	, 
	SenhaUsuario	CHAR(10) NOT NULL	,
	NomeUsuario		CHAR(30) NULL		,
	AtivoUsuario	CHAR(01)
	)
GO
-- criando usuario Teste - Inativo
-- senha criptografada ('1234')
INSERT Usuario (SenhaUsuario, NomeUsuario, AtivoUsuario) VALUES ('~†‚u','teste','N')
GO
-- criando tabela Permissões Usuarios
CREATE TABLE Permissao(
	CodTela		CHAR(5) NOT NULL,
	CodUsuario	INT
	)
GO
-- criando Telas
CREATE TABLE Telas (
	CodTela		  CHAR(05)	NOT NULL ,
	DescricaoTela CHAR(50)  NULL 
) ON [PRIMARY]
GO
-- criando registro das Tela
INSERT Telas (CodTela, DescricaoTela ) VALUES ('01','Transacoes')
GO
INSERT Telas (CodTela, DescricaoTela ) VALUES ('02','Usuarios/Permissoes')
GO
INSERT Telas (CodTela, DescricaoTela ) VALUES ('20','Exportacao Excel')
GO
INSERT Telas (CodTela, DescricaoTela ) VALUES ('99','Back-Up')
GO
-- criando LogOperacao
CREATE TABLE LogOperacao(
	Data_Operacao	DATETIME NOT NULL,
	Tela			char(20) NOT NULL,
	Usuario			CHAR(30) NOT NULL,
	Operacao		VARCHAR(255)
	)
GO


