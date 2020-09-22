

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


Public Sub Delete()
	Dim CMD As ADODB.Command
	Dim PM as ADODB.Parameter

	On Error Goto DeleteError

	Set CMD = New ADODB.Command

	With CMD
		.ActiveConnection = adoConn
		.CommandText = "Delete%TABLE_NAME%"
		.CommandType = adCmdStoredProc
	End With

%FOR_EACH_PK_FLD%
	Set PM = CMD.CreateParameter("@%FLD_NAME%", %ADO_PARAM_TYPE%, adParamInput)
	CMD.Parameters.Append PM
	PM.Value = m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME%

%END_FOR%

	CMD.Execute

DeleteExit:
	Exit Sub

DeleteError:
	GenError
	Resume DeleteExit

End Sub

Public Sub Insert()
	Dim CMD As ADODB.Command
	Dim PM as ADODB.Parameter

	On Error Goto InsertError

	Set CMD = New ADODB.Command

	With CMD
		.ActiveConnection = adoConn
		.CommandText = "Insert%TABLE_NAME%"
		.CommandType = adCmdStoredProc
	End With

%FOR_EACH_FLD% 
	Set PM = CMD.CreateParameter("@%FLD_NAME%", %ADO_PARAM_TYPE%, adParamInput)
	CMD.Parameters.Append PM
	PM.Value = m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME%

%END_FOR%

	CMD.Execute

InsertExit:
	Exit Sub

InsertError:
	GenError
	Resume InsertExit

End Sub

Public Sub GetData()
	Dim CMD As ADODB.Command
	Dim PM as ADODB.Parameter
	Dim RS As ADODB.Recordset

%FOR_EACH_FLD%
	CONST %FLD_NAME%% = %FLD_POS_BASE_0%
%END_FOR%

	On Error Goto GetDataError

	Set CMD = New ADODB.Command

	With CMD
		.ActiveConnection = adoConn
		.CommandText = "Select%TABLE_NAME%"
		.CommandType = adCmdStoredProc
	End With

%FOR_EACH_PK_FLD%
	Set PM = CMD.CreateParameter("@%FLD_NAME%", %ADO_PARAM_TYPE%, adParamInput)
	CMD.Parameters.Append PM
	PM.Value = m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME%

%END_FOR%

	Set RS = CMD.Execute

	If RS.RecordCount Then

		Do Until RS.EOF

%FOR_EACH_FLD%
			m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME% = RS(%FLD_NAME%)
%END_FOR%

			RS.MoveNext

		Loop

	End If

GetDataExit:
	If Not RS Is Nothing Then
		If RS.State = adStateOpen Then
			RS.Close
		End If
		Set RS = Nothing
	End If
	Exit Sub

GetDataError:
	GenError
	Resume GetDataExit

End Sub

Public Sub Update()
	Dim CMD As ADODB.Command
	Dim PM as ADODB.Parameter

	On Error Goto UpdateError

	Set CMD = New ADODB.Command

	With CMD
		.ActiveConnection = adoConn
		.CommandText = "Update%TABLE_NAME%"
		.CommandType = adCmdStoredProc
	End With

%FOR_EACH_FLD%
	Set PM = CMD.CreateParameter("@%FLD_NAME%", %ADO_PARAM_TYPE%, adParamInput)
	CMD.Parameters.Append PM
	PM.Value = m_%VB_FLD_TYPE_PREFIX%_%FLD_NAME%

%END_FOR%

	CMD.Execute

UpdateExit:
	Exit Sub

UpdateError:
	GenError
	Resume UpdateExit

End Sub

