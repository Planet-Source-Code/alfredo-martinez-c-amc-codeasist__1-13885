

If (Select Count(*) From sysobjects Where name = 'sp_Sel_%TABLE_NAME%') > 0 
   Begin
      Drop Procedure sp_Sel_%TABLE_NAME%
   End
GO


/*
--  Nombre: sp_Sel_%TABLE_NAME% basado en la tabla %TABLE_NAME%
--  Autor : %USER_NAME%
--  Fecha : %DATE%   %TIME%
--  Observaciones: Procedimiento almacenado para la seleccion
--                 de registro de la tabla %TABLE_NAME%
*/


CREATE PROCEDURE sp_Sel_%TABLE_NAME%
%FOR_EACH_PK_FLD% 
   @%FLD_NAME% %FLD_TYPE% , 
%END_FOR%

%REMOVE_LAST_COMMA%

AS
   SELECT  

%FOR_EACH_FLD% 
      %FLD_NAME% , 
%END_FOR%

%REMOVE_LAST_COMMA%

   FROM
      %TABLE_NAME%
   WHERE

%FOR_EACH_PK_FLD% 
      %FLD_NAME% = @%FLD_NAME% AND 
%END_FOR%


%REMOVE_LAST_AND%

