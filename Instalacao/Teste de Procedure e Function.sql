-- Teste Procedure
exec [procCalculaTotalTransacoes] '2026-15-06', '2026-15-06', 'Pendente'
GO
-- Teste Function
SELECT * 
FROM listarTransacoesPorPeriodo('2026-01-01 00:00:00', '2026-31-12 23:59:59')
ORDER BY Data_Transacao;
-- Teste VIEW
SELECT * FROM dbo.vw_ConsolidadoFinanceiro
ORDER BY Ano, Mes, Total_Valor DESC;