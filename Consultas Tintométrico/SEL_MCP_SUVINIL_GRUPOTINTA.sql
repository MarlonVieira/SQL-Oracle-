SELECT 
  '5' AS IDCOLECAO,
  CONVERT(VARCHAR,COD_GRUPO) AS IDGRUPOTINTA,	
  SUBSTRING(DES_GRUPO,1,60) AS DESCRICAO,
  SITUACAO
FROM SELFCOLOR.dbo.V_TWEB_GRUPOS 
WHERE SITUACAO = 'A';