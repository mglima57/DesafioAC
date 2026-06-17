-- acessando novo banco
USE DESAFIOAC
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- apaga preventivamente
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[listarTransacoesPorPeriodo]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
DROP FUNCTION [dbo].[listarTransacoesPorPeriodo]
GO
-- cria a function
CREATE FUNCTION dbo.listarTransacoesPorPeriodo (
    @Data_Inicial 	DATETIME	= 	NULL,
    @Data_Final 	DATETIME	= 	NULL
)
RETURNS TABLE
AS
RETURN
(

    -- retorno
	SELECT  	ID_Transacao,
				Numero_Cartao,
				Valor_Transacao,
				Data_Transacao,
				Descricao,
				Status_Transacao,
				dbo.funcCategoria(Valor_Transacao) AS Categoria
	FROM 		Transacoes
	WHERE		Data_Transacao >= ISNULL(@Data_Inicial, CONVERT(DATETIME, '1900-01-01 00:00:00', 120))
			AND Data_Transacao <= ISNULL(@Data_Final, 	CONVERT(DATETIME, '31-12-2055 23:59:59', 103))
);	
GO
