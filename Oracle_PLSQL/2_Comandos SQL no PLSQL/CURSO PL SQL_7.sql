
DECLARE
   v_ID SEGMERCADO.ID%type := 3;
BEGIN
   DELETE FROM SEGMERCADO WHERE ID = v_ID;
   COMMIT;
END;

SELECT * FROM SEGMERCADO;