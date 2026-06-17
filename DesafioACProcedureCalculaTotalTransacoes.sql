-- acessando novo banco
USE DESAFIOAC
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- apaga preventivamente
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[procCalculaTotalTransacoes]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[procCalculaTotalTransacoes]
GO

-- cria procedure
CREATE PROCEDURE [dbo].[procCalculaTotalTransacoes]
@Data_Inicial 		DATETIME	= 	NULL,
@Data_Final 		DATETIME	= 	NULL,
@Status_Transacao 	VARCHAR(10)	=	NULL
AS

SET NOCOUNT ON
SET ANSI_Warnings OFF

BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	SET NOCOUNT ON;

	-- tratamento dos parametros passados de datas
	IF @Data_Inicial IS NULL
	BEGIN
		SELECT	@Data_Inicial = '1900-01-01'
	END
	IF @Data_Final IS NULL
	BEGIN
		SELECT @Data_Final = '2055-31-12'
	END
	SELECT @Data_Inicial 	= @Data_Inicial + ' 00:00:00'
	SELECT @Data_Final 		= @Data_Final 	+ ' 23:59:59'

    -- retorno
	SELECT  	Numero_Cartao,
				Valor_Total 			= SUM(Valor_Transacao),
				Quantidade_Transacoes 	= COUNT(*),
				Status_Transacao
	FROM 		Transacoes
	WHERE		Data_Transacao		>= @Data_Inicial
			AND	Data_Transacao		<= @Data_Final
			AND Status_Transacao	= CASE WHEN @Status_Transacao IS NOT NULL
										   THEN @Status_Transacao		
										   ELSE Status_Transacao						
										   END
	GROUP BY 	Numero_Cartao, 
				Valor_Transacao, 
				Status_Transacao

END
GO
