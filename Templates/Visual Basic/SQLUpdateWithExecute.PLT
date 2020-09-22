

OPTION EXPLICIT
'Usar con SQL
'Basado en la tabla %TABLE_NAME%
'Los parametros son basados en la clase obj%TABLE_NAME%
'La plantilla requiere una funcion de validaciones de error
'La plantilla requiere una conexion ADO
'Fecha - %DATE%
'Autor - %USER_NAME%

Sub Update%TABLE_NAME% (obj%TABLE_NAME% As cls%TABLE_NAME%)
Dim sSQL As String
Dim CMD As ADODB.Command
Dim nRowsAffected As Long


On Error Goto UpdateError%TABLE_NAME%

	sSQL = "UPDATE %TABLE_NAME%"
	sSQL = sSQL & " SET"
   WITH obj%TABLE_NAME%
%PK_OFF%
%FOR_EACH_FLD%
	sSQL = sSQL & " %FLD_NAME% = " & .%VB_FLD_TYPE_PREFIX%_%FLD_NAME% & ","
%END_FOR%

%REMOVE_LAST_COMMA% %PK_ON%


	sSQL = sSQL & " WHERE"
%FOR_EACH_PK_FLD%
	sSQL = sSQL & " %FLD_NAME% = " & .%VB_FLD_TYPE_PREFIX%_%FLD_NAME% & "AND"
%END_FOR%

%REMOVE_LAST_AND%
   END WITH 

	Set CMD = New ADODB.Command

	With CMD
		.ActiveConnection = adoConn
		.CommandText = sSQL
		.CommandType = adCmdText
		.Execute nRowsAffected
	End With

UpdateExit%TABLE_NAME%:
	Exit Sub

UpdateError%TABLE_NAME%:
	GenError
	Resume UpdateExit%TABLE_NAME%

End Sub


