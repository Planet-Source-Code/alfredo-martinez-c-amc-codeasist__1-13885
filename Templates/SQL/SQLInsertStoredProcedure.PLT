

If (Select Count(*) From sysobjects Where name = 'sp_Ins_%TABLE_NAME%') > 0 
   Begin
      Drop Procedure sp_Ins_%TABLE_NAME%
   End
GO

/*
--  Nombre: sp_Ins_%TABLE_NAME% basado en la tabla %TABLE_NAME%
--  Autor : %USER_NAME%
--  Fecha : %DATE%   %TIME%
--  Observaciones: Procedimiento almacenado para la inserccion
                 de registro de la tabla %TABLE_NAME%
*/

CREATE PROCEDURE sp_Ins_%TABLE_NAME%
%FOR_EACH_FLD% 
   @%FLD_NAME% %FLD_TYPE% , 
%END_FOR%

%REMOVE_LAST_COMMA%

AS

Declare @ins_error int , 
	@CTransaccion varchar(30)

BEGIN TRAN @CTransaccion

   INSERT INTO 
      %TABLE_NAME%
   (

%FOR_EACH_FLD% 
      %FLD_NAME% , 
%END_FOR%

%REMOVE_LAST_COMMA%

   )
   VALUES
   (

%FOR_EACH_FLD% 
      @%FLD_NAME% , 
%END_FOR%

%REMOVE_LAST_COMMA%

)

   Set @ins_error = @@Error
   If @ins_error = 0 
      Begin 
         Print 'El registro se ha eliminado satisfactoriamente'
         Commit Transaction @CTransaccion
      End 
   Else 
      Begin 
         If @ins_error <> 0
         Print 'Ha ocurrido un error al eliminar el registro'
         Rollback Transaction 
      End
GO

