

If (Select Count(*) From sysobjects Where name = 'sp_Del_%TABLE_NAME%') > 0 
   Begin
      Drop Procedure sp_Del_%TABLE_NAME%
   End
GO


/*
---  Nombre: sp_Del_%TABLE_NAME% basado en la tabla %TABLE_NAME%
---  Autor : %USER_NAME%
---  Fecha : %DATE%   %TIME%
---  Observaciones: Procedimiento almacenado para la eliminacion
---                 de registro de la tabla %TABLE_NAME%
*/


CREATE PROCEDURE sp_Del_%TABLE_NAME%
%FOR_EACH_PK_FLD% 
   @%FLD_NAME% %FLD_TYPE% , 
%END_FOR%

%REMOVE_LAST_COMMA%

AS

Declare @del_error int , 
        @CTransaccion varchar(30)

BEGIN TRAN @CTransaccion

   DELETE FROM 
      %TABLE_NAME%
   WHERE
%FOR_EACH_PK_FLD% 
      %FLD_NAME% = @%FLD_NAME% AND
%END_FOR%

%REMOVE_LAST_AND%

   Set @del_error = @@Error

   If @del_error = 0 
      Begin 
         Print 'El registro se ha eliminado satisfactoriamente'
	 Commit Transaction @CTransaccion
      End 
   Else 
      Begin 
         If @del_error <> 0
         Print 'Ha ocurrido un error al eliminar el registro'
         Rollback Transaction 
   End

GO

