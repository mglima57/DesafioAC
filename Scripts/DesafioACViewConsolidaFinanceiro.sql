-- acessando novo banco
USE DESAFIOAC
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- apaga preventivamente
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[viewConsolidadoFinanceiro]'))
DROP VIEW [dbo].[viewConsolidadoFinanceiro]
GO
CREATE VIEW dbo.viewConsolidadoFinanceiro
AS
SELECT 
    YEAR(Data_Transacao) 				AS Ano,
    MONTH(Data_Transacao) 				AS Mes,
    dbo.funcCategoria(Valor_Transacao) 	AS Categoria,
    COUNT(ID_Transacao) 				AS Qtd_Transacoes,
    SUM(Valor_Transacao) 				AS Total_Valor,
    AVG(Valor_Transacao) 				AS Media_Valor
FROM 
    Transacoes
GROUP BY 
    YEAR(Data_Transacao),
    MONTH(Data_Transacao),
    dbo.funcCategoria(Valor_Transacao);
GO
