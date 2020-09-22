

Sub Select%TABLE_NAME% (obj%TABLE_NAME% As cls%TABLE_NAME%)
Dim sSQL As String
Dim RS As ADODB.Recordset

	'Usar con SQL
	'Basado en la tabla %TABLE_NAME%
	'Los parametros son basados en la clase obj%TABLE_NAME%
	'La plantilla requiere una funcion de validaciones de error
	'La plantilla requiere una conexion ADO
	'Fecha - %DATE%
	'Autor - %USER_NAME%

%FOR_EACH_FLD%
CONST %FLD_NAME% = %FLD_POS_BASE_0%
%END_FOR%

On Error Goto SelectError%TABLE_NAME%

   Set RS = New ADODB.Recordset

WITH obj%TABLE_NAME%	
   sSQL = "SELECT"
%PK_OFF%
%FOR_EACH_FLD%
   sSQL = sSQL & " %FLD_NAME% , "
%END_FOR%

%REMOVE_LAST_COMMA% %PK_ON%
   sSQL = sSQL & " FROM %TABLE_NAME%"
   sSQL = sSQL & " WHERE "
%FOR_EACH_PK_FLD%
   sSQL = sSQL & " %FLD_NAME% = " & .%FLD_NAME% & " AND "
%END_FOR%

%REMOVE_LAST_AND%
END WITH

   Set RS.ActiveConnection = ADOConn
   RS.Open sSQL

   If RS.RecordCount Then
      WITH obj%TABLE_NAME% 
%FOR_EACH_FLD%
         .%VB_FLD_TYPE_PREFIX%_%FLD_NAME% = CheckNull(RS(%FLD_NAME%))
%END_FOR%

      END WITH
   End If

SelectExit%TABLE_NAME%:
	If Not RS Is Nothing Then
		If RS.State = adStateOpen Then
			RS.Close
		End If
		Set RS = Nothing
	End If
	Exit Sub

SelectError%TABLE_NAME%:
	GenError
	Resume SelectExit%TABLE_NAME%

End Sub

