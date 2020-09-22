

If (Select Count(*) From sysobjects Where name = 'sp_Udp_%TABLE_NAME%') > 0 
   Begin
      Drop Procedure sp_Udp_%TABLE_NAME%
   End
GO


/*
--  Nombre: sp_Upd_%TABLE_NAME% basado en la tabla %TABLE_NAME%
--  Autor : %USER_NAME%
--  Fecha : %DATE%   %TIME%
--  Observaciones: Procedimiento almacenado para la seleccion
--                 de registro de la tabla %TABLE_NAME%
*/


CREATE PROCEDURE sp_Udp_%TABLE_NAME%
%FOR_EACH_FLD% 
   @%FLD_NAME% %FLD_TYPE% , 
%END_FOR%

%REMOVE_LAST_COMMA%
AS

Declare @upd_error int , 
	@CTransaccion varchar(30)

BEGIN TRAN @CTransaccion

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

   Set @upd_error = @@Error

   If @upd_error = 0 
      Begin 
         Print 'El registro se ha eliminado satisfactoriamente'
         Commit Transaction @CTransaccion
      End 
   Else 
      Begin 
         If @upd_error <> 0
         Print 'Ha ocurrido un error al eliminar el registro'
         Rollback Transaction 
      End

GO


