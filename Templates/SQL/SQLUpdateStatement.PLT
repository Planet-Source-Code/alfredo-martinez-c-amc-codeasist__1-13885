

/*
--  Nombre: Instruccion de actuallizacion de registros
--	  : de la tabla %TABLE_NAME%
--  Autor : %USER_NAME%
*/

DECLARE 
%FOR_EACH_FLD%  
   @%FLD_NAME% %FLD_TYPE% , 
%END_FOR%

%REMOVE_LAST_COMMA%
UPDATE  %TABLE_NAME% SET 

%FOR_EACH_FLD% 
   %FLD_NAME% = @%FLD_NAME% , 
%END_FOR%

%REMOVE_LAST_COMMA%
WHERE
%FOR_EACH_PK_FLD%  
   %FLD_NAME%  = @%FLD_NAME% AND 
%END_FOR%

%REMOVE_LAST_AND%

