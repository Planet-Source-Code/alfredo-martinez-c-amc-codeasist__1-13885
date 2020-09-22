

/*
--  Nombre: Declaracion de variables de la tabla %TABLE_NAME%
*/

DECLARE 
%FOR_EACH_FLD%  
   @%FLD_NAME% %FLD_TYPE% , 
%END_FOR%

%REMOVE_LAST_COMMA%