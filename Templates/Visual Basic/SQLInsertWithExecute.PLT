

Sub Insert%TABLE_NAME% (obj%TABLE_NAME% As cls%TABLE_NAME%)
Dim sSQL As String
Dim CMD As ADODB.Command
Dim nRowsAffected As Long

	'Usar con SQL
	'Basado en la tabla %TABLE_NAME%
	'Los parametros son basados en la clase obj%TABLE_NAME%
	'La plantilla requiere una funcion de validaciones de error
	'La plantilla requiere una conexion ADO
	'Fecha - %DATE%
	'Autor - %USER_NAME%

On Error Goto InsertError%TABLE_NAME%

WITH obj%TABLE_NAME%

	sSQL = "INSERT %TABLE_NAME% ("
%FOR_EACH_FLD%
	sSQL = sSQL & " %FLD_NAME% , " 
%END_FOR%

%REMOVE_LAST_COMMA%
	sSQL = sSQL & ") VALUES ("
%FOR_EACH_FLD%
	sSQL = sSQL & .%VB_FLD_TYPE_PREFIX%_%FLD_NAME% , 
%END_FOR%

%REMOVE_LAST_COMMA%
	sSQL = sSQL & ")"
END WITH

	Set CMD = New ADODB.Command

	With CMD
		.ActiveConnection = adoConn
		.CommandText = sSQL
		.CommandType = adCmdText
		.Execute nRowsAffected
	End With

InsertExit%TABLE_NAME%:
	Exit Sub

InsertError%TABLE_NAME%:
	GenError
	Resume InsertExit%TABLE_NAME%

End Sub

