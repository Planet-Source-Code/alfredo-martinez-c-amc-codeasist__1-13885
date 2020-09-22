

/*
--  Nombre: Instruccion de inserccion de registros
--	  : de la tabla %TABLE_NAME%
--  Autor : %USER_NAME%
*/


DECLARE 
%FOR_EACH_FLD%  
   @%FLD_NAME% %FLD_TYPE% , 
%END_FOR%

%REMOVE_LAST_COMMA%
INSERT INTO %TABLE_NAME% 
(
%FOR_EACH_FLD% 
   %FLD_NAME% , 
%END_FOR%

%REMOVE_LAST_COMMA%
)

VALUES

(

%FOR_EACH_FLD%  
   @%FLD_NAME%  , 
%END_FOR%

%REMOVE_LAST_COMMA%
)
