SELECT 
   CONVERT(VARCHAR,F.ID) AS IDFORMULA,
   CONVERT(VARCHAR,C.COD_COLORANTE) AS IDCORANTE,
   CONVERT(VARCHAR,ROW_NUMBER() OVER (PARTITION BY F.ID ORDER BY F.ID)) AS SEQCORANTE,
   '' AS QTDEOZ1,
   '' AS QTDEOZ2,
   REPLACE(C.QTDE * 0.154,'.',',') AS QTDEML,
   '' AS QTDEGR,
   '5' AS IDCOLECAO
 FROM SELFCOLOR.dbo.V_TWEB_CORES F
 INNER JOIN SELFCOLOR.dbo.V_TWEB_CORES_COLORANTES C
    ON F.COD_GRUPO     = C.COD_GRUPO
   AND F.COD_PRODUTO   = C.COD_PRODUTO
   AND F.COD_BASE      = C.COD_BASE
   AND F.COD_EMBALAGEM = C.COD_EMBALAGEM
   AND F.COD_COR       = C.COD_COR
 INNER JOIN SELFCOLOR.dbo.V_TWEB_GRUPOS G ON C.COD_GRUPO = G.COD_GRUPO 
 INNER JOIN SELFCOLOR.dbo.V_TWEB_COLORANTES R ON C.COD_BASE = R.COD_COLORANTE
 INNER JOIN SELFCOLOR.dbo.V_TWEB_PRODUTOS P ON C.COD_PRODUTO = P.COD_PRODUTO
 INNER JOIN SELFCOLOR.dbo.V_TWEB_TIPOS_BASES TB ON C.COD_BASE = TB.COD_BASE
 INNER JOIN SELFCOLOR.dbo.V_TWEB_EMBALAGENS E ON C.COD_EMBALAGEM = E.COD_EMBALAGEM
 WHERE C.QTDE > 0
 AND F.SITUACAO = 'A'
 AND G.SITUACAO = 'A';