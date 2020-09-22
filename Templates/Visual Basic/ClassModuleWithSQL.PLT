

Option Explicit

	'Usar con SQL
	'Basado en la tabla %TABLE_NAME%
	'Los parametros son basados en la clase obj%TABLE_NAME%
	'La plantilla requiere una funcion de validaciones de error
	'La plantilla requiere una conexion ADO
	'Fecha - %DATE%
	'Autor - %USER_NAME%

%FOR_EACH_FLD%
Private m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME% As %VB_FLD_TYPE%
%END_FOR%

%FOR_EACH_FLD%

Public Property Let %FLD_NAME% (ByVal vNew%FLD_NAME% As %VB_FLD_TYPE%)
	m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME% = vNew%FLD_NAME%
End Property

Public Property Get %FLD_NAME%  () As %VB_FLD_TYPE%
	%FLD_NAME% = m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME% 
End Property

%END_FOR%

'Procedimiento para eliminar un registro
Public Sub Delete()
	Dim sSQL As String
	Dim CMD As ADODB.Command
	Dim nRowsAffected As Long

	On Error Goto DeleteError%TABLE_NAME%

	sSQL = "DELETE FROM %TABLE_NAME%"
	sSQL = sSQL & " WHERE"  
%PK_OFF%
%FOR_EACH_PK_FLD%
	sSQL = sSQL & " %FLD_NAME% = " & m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME% & " AND "
%END_FOR%

%REMOVE_LAST_AND%

%PK_ON%
	Set CMD = New ADODB.Command

	With CMD
		.ActiveConnection = adoConn
		.CommandText = sSQL
		.CommandType = adCmdText
		.Execute nRowsAffected
	End With

DeleteExit%TABLE_NAME%:
	Exit Sub

DeleteError%TABLE_NAME%:
	GenError
	Resume DeleteExit%TABLE_NAME%

End Sub


'Procedimiento para insertar un registro

Public Sub Insert()
	Dim sSQL As String
	Dim CMD As ADODB.Command
	Dim nRowsAffected As Long

	On Error Goto InsertError%TABLE_NAME%

	sSQL = "INSERT %TABLE_NAME% ("
%FOR_EACH_FLD%
	sSQL = sSQL & " %FLD_NAME% , "
%END_FOR%

%REMOVE_LAST_COMMA%

	sSQL = sSQL & ") VALUES (" 
%FOR_EACH_FLD%
	sSQL = sSQL & m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME% & " , "
%END_FOR%

%REMOVE_LAST_COMMA%
	sSQL = sSQL & ")"

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


'//Procedimiento para seleccionar un registro

Public Sub GetData()
	Dim sSQL As String
	Dim RS As ADODB.Recordset

%FOR_EACH_FLD%
	CONST %FLD_NAME% = %FLD_POS_BASE_0%
%END_FOR%

	On Error Goto GetDataError%TABLE_NAME%

	Set RS = New ADODB.Recordset
	
	sSQL = "SELECT"
%FOR_EACH_FLD%
	sSQL = sSQL & " %FLD_NAME%,"
%END_FOR%

%REMOVE_LAST_COMMA%
	sSQL = sSQL & " FROM %TABLE_NAME%"
	sSQL = sSQL & " WHERE "
%FOR_EACH_PK_FLD%
	sSQL = sSQL & " %FLD_NAME% = " & m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME% & "AND"
%END_FOR%

%REMOVE_LAST_AND%
	Set RS.ActiveConnection = ADOConn
	RS.Open sSQL

	If RS.RecordCount Then

%FOR_EACH_FLD%
        If Not IsNull(RS(%FLD_NAME%)) Then m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME% = RS(%FLD_NAME%)
%END_FOR%

	End If

GetDataExit%TABLE_NAME%:
	If Not RS Is Nothing Then
		If RS.State = adStateOpen Then
			RS.Close
		End If
		Set RS = Nothing
	End If
	Exit Sub

GetDataError%TABLE_NAME%:
	GenError
	Resume GetDataExit%TABLE_NAME%

End Sub



'//Procedimiento para actualizar un registro

Public Sub Update()
	Dim sSQL As String
	Dim CMD As ADODB.Command
	Dim nRowsAffected As Long

	On Error Goto UpdateError%TABLE_NAME%

	sSQL = "UPDATE %TABLE_NAME%"
	sSQL = sSQL & " SET"
%PK_OFF%
%FOR_EACH_FLD%
	sSQL = sSQL & " %FLD_NAME% = " & m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME% & "," 
%END_FOR%

%REMOVE_LAST_COMMA%
%PK_ON%

	sSQL = sSQL & " WHERE"

%FOR_EACH_PK_FLD%
	sSQL = sSQL & "%FLD_NAME% = " & m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME% & "AND"
%END_FOR%

%REMOVE_LAST_AND%

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

