-- acessando novo banco
USE DESAFIOAC
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- apaga preventivamente
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[funcCategoria]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
DROP FUNCTION [dbo].[funcCategoria]
GO
-- cria function
CREATE FUNCTION funcCategoria
(
@Valor_Transacao DECIMAL(18, 2)
)
RETURNS VARCHAR(10)
AS
BEGIN

    DECLARE @Categoria VARCHAR(10);

    SET @Categoria = CASE 
        WHEN @Valor_Transacao > 2000 	THEN 'Premium'
        WHEN @Valor_Transacao >= 1000 	AND @Valor_Transacao <= 2000 	THEN 'Alta'
        WHEN @Valor_Transacao >= 500 	AND @Valor_Transacao < 1000 	THEN 'Média'
        WHEN @Valor_Transacao < 500 	THEN 'Baixa'
        ELSE 'Indefinida'
    END;

    RETURN @Categoria;

END
GO

