SELECT 
   CONVERT(VARCHAR,C.ID) AS IDFORMULA,
   '5' AS IDCOLECAO,
   CONVERT(VARCHAR,C.COD_BASE) AS IDBASE,
   CONVERT(VARCHAR,C.COD_PRODUTO) AS IDPRODUTO,
   COALESCE((SELECT CAST(T.COD_MENSAGEM AS VARCHAR)
   FROM SELFCOLOR.dbo.TMENSAGENS_FORMULAS T 
   JOIN SELFCOLOR.dbo.TMENSAGENS T2 ON T2.COD_MENSAGEM = T.COD_MENSAGEM
   WHERE T.COD_BASE = C.COD_BASE
   AND T.COD_COR = C.COD_COR
   AND T.COD_EMBALAGEM = C.COD_EMBALAGEM
   AND T.COD_GRUPO = C.COD_GRUPO
   AND T.COD_PRODUTO = C.COD_PRODUTO
   AND T.SITUACAO = 'A'),'') AS IDMSGTINTA,
   C.NOM_COR AS DESCRICAO,
   C.DES_COR AS CODCATALOGO,
   CONVERT(VARCHAR,C.COD_EMBALAGEM) AS IDEMBTINTA,
   CONVERT(VARCHAR,C.COD_GRUPO) AS IDGRUPOTINTA,
   'N' AS INATIVO,
   'N' AS PERSONALIZADA
 FROM SELFCOLOR.dbo.V_TWEB_CORES C
 INNER JOIN SELFCOLOR.dbo.V_TWEB_GRUPOS G ON C.COD_GRUPO = G.COD_GRUPO
 INNER JOIN SELFCOLOR.dbo.V_TWEB_PRODUTOS P ON C.COD_PRODUTO = P.COD_PRODUTO
 INNER JOIN SELFCOLOR.dbo.V_TWEB_TIPOS_BASES TB ON C.COD_BASE = TB.COD_BASE
 INNER JOIN SELFCOLOR.dbo.V_TWEB_EMBALAGENS E ON C.COD_EMBALAGEM = E.COD_EMBALAGEM
 WHERE C.SITUACAO = 'A'
 AND G.SITUACAO = 'A';