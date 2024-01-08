SET SERVEROUTPUT ON;

DECLARE
    V_ID NUMBER(5) := 10;
 BEGIN
    V_ID := 20;
    dbms_output.put_line(V_ID);
 END;