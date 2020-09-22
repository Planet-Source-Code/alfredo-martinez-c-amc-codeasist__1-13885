

'Usar con SQL
'Basado en la tabla %TABLE_NAME%
'Fecha - %DATE%
'Autor - %USER_NAME%

Type typ%TABLE_NAME%
%FOR_EACH_FLD%
   %VB_FLD_TYPE_PREFIX%_%FLD_NAME% As %VB_FLD_TYPE%
%END_FOR%

End Type

