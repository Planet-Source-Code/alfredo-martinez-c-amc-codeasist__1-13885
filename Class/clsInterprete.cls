VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterprete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Variable publica para la transcripcion a generar
Dim strCodigoGenerado As String

Private m_strTemplate As String
Private m_strTabla As String
Private m_objCampos As New clsCampos
Private m_objVariables As New clsCampos
Private m_objWhere As New clsCampos
Private m_objOrderBy As New clsCampos

Private obj_Lineas As New clsLineas

'Variables de uso interno de la clase
Private bolConPK As Boolean

Private i As Long, I2 As Long, I3 As Long, I4 As Long


Public Property Get Template() As String
   Template = m_strTemplate
End Property

Public Property Let Template(ByVal vNewValue As String)
   m_strTemplate = vNewValue
End Property

Public Property Get Tabla() As String
   Tabla = m_strTabla
End Property

Public Property Let Tabla(ByVal vNewValue As String)
   m_strTabla = vNewValue
End Property

Public Property Get Campos() As clsCampos
   Set Campos = m_objCampos
End Property

Public Property Let Campos(ByVal vNewValue As clsCampos)
   Set m_objCampos = vNewValue
End Property

Public Property Get Variables() As clsCampos
   Set Variables = m_objVariables
End Property

Public Property Let Variables(ByVal vNewValue As clsCampos)
   Set m_objVariables = vNewValue
End Property

Public Property Get Where() As clsCampos
   Set Where = m_objWhere
End Property

Public Property Let Where(ByVal vNewValue As clsCampos)
   Set m_objWhere = vNewValue
End Property

Public Property Get OrderBy() As clsCampos
   Set OrderBy = m_objOrderBy
End Property

Public Property Let OrderBy(ByVal vNewValue As clsCampos)
   Set m_objOrderBy = vNewValue
End Property



'Funcion para abrir la plantilla de codigo
Private Function OpenTemplete() As Boolean
Dim strLinea As String
   If Trim$(m_strTemplate) = "" Then
      MsgBox "No se ha especificado el nombre de la plantilla", vbInformation, App.Title
      OpenTemplete = False: Exit Function
   End If
   With obj_Lineas
      Open m_strTemplate For Input As #1
      While Not EOF(1)
          Line Input #1, strLinea
          .Add strLinea & vbCrLf
      Wend
      Close #1
    End With
   
   OpenTemplete = True
End Function


'Funcion de entrada para la generacion de codigo fuente.
Public Function GeneraCodigo() As String
Dim lngItem As Long, lngCounter As Long
   strCodigoGenerado = ""
   If OpenTemplete = True Then
      With obj_Lineas
         lngItem = 1
         For lngCounter = 1 To .Count
            strCodigoGenerado = strCodigoGenerado & TranscribeLinea(.Item(lngCounter), lngItem)
            lngCounter = lngItem
         Next lngCounter
      End With
   End If
   GeneraCodigo = strCodigoGenerado
End Function

'Funcion para indentificar el comando en la linea de texto
Private Function TranscribeLinea(ByVal objLinea As clsLinea, ByRef lngItem As Long) As String
Dim lngPosIni As Long, lngPosFin As Long, i As Long, I2 As Long
Dim strKeyWord As String, strTextoLinea As String
   With objLinea
      For i = 1 To Len(.Texto)
         If Mid(.Texto, i, 1) = "%" Then
            lngPosIni = i
            strKeyWord = "%"
            For I2 = (lngPosIni) + 1 To Len(.Texto)
               If Mid(.Texto, I2, 1) = "%" Then
                  strKeyWord = strKeyWord & Mid(.Texto, I2, 1)
                  strKeyWord = InterpretaComando(strKeyWord, lngItem)
                  Exit For
               Else
                  strKeyWord = strKeyWord & Mid(.Texto, I2, 1)
               End If
            Next I2
            i = (I2 + 1)
         End If
         strTextoLinea = strTextoLinea & strKeyWord
         strTextoLinea = strTextoLinea & Mid(.Texto, i, 1)
         strKeyWord = ""
      Next i
   End With
   TranscribeLinea = strTextoLinea
   lngItem = lngItem + 1
End Function


'Funcion para interpretar un comando
Private Function InterpretaComando(ByVal strComando As String, ByRef lngItem As Long) As String
Dim strResultado As String
   strResultado = ""
   Select Case strComando
      Case "%DATE%"
         strResultado = Format(CDate(Now), "DD/MM/YYYY")
      Case "%TIME%"
         strResultado = Format(CDate(Now), "HH:MM:SS AMPM")
      Case "%USER_NAME%"
         strResultado = "Alfredo Martínez C."
      Case "%REMOVE_LAST_COMMA%"
         Call QuitaUltimaComa
      Case "%REMOVE_LAST_AND%"
         Call QuitaUltimoAnd
      Case "%PK_ON%"
         bolConPK = True
      Case "%PK_OFF%"
         bolConPK = False
      Case "%TABLE_NAME%"
         strResultado = m_strTabla
      Case "%FOR_EACH_FLD%"
         strResultado = InterpretaNextFld(lngItem)
      Case "%FOR_EACH_PK_FLD%"
         strResultado = InterpretaForEachPKFld(lngItem)
      
         
   End Select
   If Trim$(strComando) = "%REMOVE_LAST_COMMA%" Or Trim$(strComando) = "%REMOVE_LAST_AND%" _
      Or Trim$(strComando) = "%PK_ON%" Or Trim$(strComando) = "%PK_OFF%" Then
      InterpretaComando = " "
   Else
      If Trim$(strResultado) = "" Then
         InterpretaComando = strComando
      Else
         InterpretaComando = strResultado
      End If
   End If
End Function



'procedimiento para remover la ultima coma en un bucle de campos
Private Sub QuitaUltimaComa()
Dim lngCh As Long, strCad1 As String, strCad2 As String
   On Error Resume Next
   For lngCh = 1 To Len(strCodigoGenerado)
      If Mid$(strCodigoGenerado, Len(strCodigoGenerado) - (lngCh + 1), 1) = "," Then
         strCad1 = Mid$(strCodigoGenerado, 1, Len(strCodigoGenerado) - (lngCh + 2))
         strCad2 = Mid$(strCodigoGenerado, Len(strCodigoGenerado) - (lngCh + 2), lngCh)
         strCodigoGenerado = ""
         strCodigoGenerado = strCad1
         Exit For
      End If
   Next lngCh
End Sub


'procedimiento para remover la ultima coma en un bucle de campos
Private Sub QuitaUltimoAnd()
Dim lngCh As Long, strCad1 As String, strCad2 As String
   On Error Resume Next
   For lngCh = 1 To Len(strCodigoGenerado)
      If Mid$(strCodigoGenerado, Len(strCodigoGenerado) - (lngCh + 1), 3) = "AND" Then
         strCad1 = Mid$(strCodigoGenerado, 1, Len(strCodigoGenerado) - (lngCh + 2))
         strCad2 = Mid$(strCodigoGenerado, Len(strCodigoGenerado) - (lngCh - 2), lngCh)
         strCodigoGenerado = ""
         strCodigoGenerado = strCad1 & strCad2
         Exit For
      End If
   Next lngCh
End Sub




'Funcion para generar una lista de acuerdo al bucle
Private Function InterpretaNextFld(ByRef lngItem As Long) As String
Dim strResultado As String, lngCounter As Long, strCadLineas As String
Dim lngTCount As Long, lngFinCounter As Long, strTransResult As String
   With obj_Lineas
      strCadLineas = ""
      For lngCounter = (lngItem + 2) To .Count
         If Mid$(.Item(lngCounter).Texto, 1, 9) = "%END_FOR%" Then
            lngItem = lngCounter
            Exit For
         Else
            strCadLineas = strCadLineas & .Item(lngCounter).Texto
         End If
      Next lngCounter
   End With
   
   With m_objCampos
      For lngTCount = 1 To .Count
         If bolConPK = True Then
            strTransResult = strTransResult & TranscribeLineaFld(strCadLineas, lngTCount)
         Else
            If .Item(lngTCount).PK = False Then
               strTransResult = strTransResult & TranscribeLineaFld(strCadLineas, lngTCount)
            End If
         End If
      Next lngTCount
   End With
   
InterpretaNextFld = strTransResult

End Function





'Funcion para generar una lista de acuerdo al bucle y llaves primarias
Private Function InterpretaForEachPKFld(ByRef lngItem As Long) As String
Dim strResultado As String, lngCounter As Long, strCadLineas As String
Dim lngTCount As Long, lngFinCounter As Long, strTransResult As String
   With obj_Lineas
      strCadLineas = ""
      For lngCounter = (lngItem + 2) To .Count
         If Mid$(.Item(lngCounter).Texto, 1, 9) = "%END_FOR%" Then
            lngItem = lngCounter
            Exit For
         Else
            strCadLineas = strCadLineas & .Item(lngCounter).Texto
         End If
      Next lngCounter
   End With
   
   With m_objCampos
      For lngTCount = 1 To .Count
         If .Item(lngTCount).PK = True Then
            strTransResult = strTransResult & TranscribeLineaFld(strCadLineas, lngTCount)
         End If
      Next lngTCount
   End With
   
   InterpretaForEachPKFld = strTransResult
   
End Function





'Funcion para transcribir una linea de campo
Private Function TranscribeLineaFld(ByVal strTexto As String, ByVal lngItemCampo As Long) As String
Dim lngPosIni As Long, lngPosFin As Long, IC As Long, ICC As Long
Dim strKeyWord As String, strTextoLinea As String
      For IC = 1 To Len(strTexto)
         If Mid(strTexto, IC, 1) = "%" Then
            lngPosIni = IC
            strKeyWord = "%"
            For ICC = (lngPosIni) + 1 To Len(strTexto)
               If Mid(strTexto, ICC, 1) = "%" Then
                  strKeyWord = strKeyWord & Mid(strTexto, ICC, 1)
                  strKeyWord = InterpretaComandoFld(strKeyWord, lngItemCampo)
                  Exit For
               Else
                  strKeyWord = strKeyWord & Mid(strTexto, ICC, 1)
               End If
            Next ICC
            IC = (ICC + 1)
         End If
         strTextoLinea = strTextoLinea & strKeyWord
         strTextoLinea = strTextoLinea & Mid(strTexto, IC, 1)
         strKeyWord = ""
      Next IC
   TranscribeLineaFld = strTextoLinea
End Function


'Funcion para interpretar un comando de campo
Private Function InterpretaComandoFld(ByVal strComando As String, ByVal lngItemCampo As Long) As String
Dim strResultado As String
   With m_objCampos
      strResultado = ""
      Select Case strComando
         Case "%FLD_NAME%"
            strResultado = .Item(lngItemCampo).Nombre
         Case "%FLD_TYPE%"
            If .Item(lngItemCampo).Tipo = "datetime" Or .Item(lngItemCampo).Tipo = "smalldatetime" Or _
               .Item(lngItemCampo).Tipo = "int" Or .Item(lngItemCampo).Tipo = "float" Or _
               .Item(lngItemCampo).Tipo = "image" Or .Item(lngItemCampo).Tipo = "money" Or _
               .Item(lngItemCampo).Tipo = "nchar" Or .Item(lngItemCampo).Tipo = "smallint" Or _
               .Item(lngItemCampo).Tipo = "ntext" Or .Item(lngItemCampo).Tipo = "smallmoney" Or _
               .Item(lngItemCampo).Tipo = "timestamp" Or .Item(lngItemCampo).Tipo = "tinyint" Or _
               .Item(lngItemCampo).Tipo = "uniqueidentifier" Then
               strResultado = .Item(lngItemCampo).Tipo
            Else
               If .Item(lngItemCampo).Tipo = "decimal" Or .Item(lngItemCampo).Tipo = "numeric" Then
                  If .Item(lngItemCampo).Presicion = "0" Then
                     strResultado = .Item(lngItemCampo).Tipo & "(" & .Item(lngItemCampo).Escala & "," & .Item(lngItemCampo).Presicion & ")"
                  Else
                     strResultado = .Item(lngItemCampo).Tipo & "(" & .Item(lngItemCampo).Longitud & ")"
                  End If
               Else
                  If .Item(lngItemCampo).Presicion = "0" Then
                     strResultado = .Item(lngItemCampo).Tipo & "(" & .Item(lngItemCampo).Longitud & ")"
                  Else
                     strResultado = .Item(lngItemCampo).Tipo & "(" & .Item(lngItemCampo).Escala & "," & .Item(lngItemCampo).Presicion & ")"
                  End If
               End If
            End If
         Case "%VB_FLD_TYPE%"
            strResultado = ConvierteTipoSQLVB(.Item(lngItemCampo).Tipo)
         Case "%VB_FLD_TYPE_PREFIX%"
            strResultado = GeneraPrefijoSQLVB(.Item(lngItemCampo).Tipo)
         Case "%FLD_POS_BASE_0%"
            strResultado = lngItemCampo - 1
         Case "%ADO_PARAM_TYPE%"
            strResultado = TrascribeADO(.Item(lngItemCampo).Tipo)
      End Select
      If Trim$(strResultado) = "" Then
         InterpretaComandoFld = strComando
      Else
         InterpretaComandoFld = strResultado
      End If
   End With
End Function



'Funcion para transcribir un el tipo de SQL Server a Visual Basic (ADO)
Private Function TrascribeADO(ByVal strTipo As String) As String
   Select Case strTipo
      Case "binary", "image", "text", "ntext"
         TrascribeADO = "adBinary"
      Case "varbinary"
         TrascribeADO = "adVarBinary"
      Case "bit"
         TrascribeADO = "adBoolean"
      Case "char", "nchar"
         TrascribeADO = "adChar"
      Case "varchar"
         TrascribeADO = "adVarChar"
      Case "nvarchar"
         TrascribeADO = "adVarNumeric"
      Case "datetime", "smalldatetime"
         TrascribeADO = "adDBDate"
      Case "timestamp"
         TrascribeADO = "adDBTimeStamp"
      Case "decimal"
         TrascribeADO = "adDecimal"
      Case "smallint"
         TrascribeADO = "adSmallInt"
      Case "tinyint"
         TrascribeADO = "adTinyInt"
      Case "int"
         TrascribeADO = "adBigInt"
      Case "money", "smallmoney"
         TrascribeADO = "adCurrency"
      Case "numeric"
         TrascribeADO = "adNumeric"
      Case "float"
         TrascribeADO = "adDouble"
      Case "uniqueidentifier"
         TrascribeADO = "adGUID"
      Case "real"
         TrascribeADO = "adVarNumeric"
      
   End Select

End Function



'Funcion para generar variables de propiedad para VB
Public Function GeneraVariablePropiedad(ByVal strCampo As String, ByVal strTipo As String) As String
Dim strResult As String
   strResult = "Private m" & GeneraPrefijoSQLVB(strTipo) & strCampo & " As " & ConvierteTipoSQLVB(strTipo)
   GeneraVariablePropiedad = strResult
End Function


'Funcion para generar variables de propiedad para VB
Public Function GeneraPropiedades(ByVal strCampo As String, ByVal strTipo As String) As String
Dim strResult As String
   strResult = "'//Metodo de lectura de la propiedad " & strCampo & vbCrLf
   strResult = strResult & "Public Property GET " & strCampo & "() AS " & ConvierteTipoSQLVB(strTipo) & vbCrLf
   strResult = strResult & vbTab & strCampo & " = m" & GeneraPrefijoSQLVB(strTipo) & _
                        strCampo & vbCrLf
   strResult = strResult & "End Property" & vbCrLf
   strResult = strResult & "'//Metodo de escritura de la propiedad " & strCampo & vbCrLf
   strResult = strResult & "Public Property LET " & strCampo & "(ByVal vNewValue as " & ConvierteTipoSQLVB(strTipo) & ") " & vbCrLf
   strResult = strResult & vbTab & "m" & GeneraPrefijoSQLVB(strTipo) & strCampo & " = vNewValue" & vbCrLf
   strResult = strResult & "End Property" & vbCrLf
   GeneraPropiedades = strResult
End Function


'Funcion para conversion de tipo de SQL Server a Visual Basic
Private Function ConvierteTipoSQLVB(ByVal strTipo As String) As String
   Select Case strTipo
      Case "binary", "image", "varbinary", "text", "ntext"
         ConvierteTipoSQLVB = "Variant"
      Case "bit"
         ConvierteTipoSQLVB = "Byte"
      Case "char", "nchar", "varchar", "nvarchar"
         ConvierteTipoSQLVB = "String"
      Case "datetime", "smalldatetime", "timestamp"
         ConvierteTipoSQLVB = "Date"
      Case "float", "real", "numeric", "uniqueidentifier"
         ConvierteTipoSQLVB = "Double"
      Case "decimal"
         ConvierteTipoSQLVB = "Single"
      Case "tinyint", "smallint"
         ConvierteTipoSQLVB = "Integer"
      Case "int"
         ConvierteTipoSQLVB = "Long"
      Case "money", "smallmoney"
         ConvierteTipoSQLVB = "Currency"
   End Select
End Function


'Funcion para Generar el prefijo de tipo de SQL Server a Visual Basic
Private Function GeneraPrefijoSQLVB(ByVal strTipo As String) As String
   Select Case strTipo
      Case "binary", "image", "varbinary", "text", "ntext"
         GeneraPrefijoSQLVB = "var"
      Case "bit"
         GeneraPrefijoSQLVB = "byt"
      Case "char", "nchar", "varchar", "nvarchar"
         GeneraPrefijoSQLVB = "str"
      Case "datetime", "smalldatetime", "timestamp"
         GeneraPrefijoSQLVB = "dat"
      Case "float", "real", "numeric", "uniqueidentifier"
         GeneraPrefijoSQLVB = "dbl"
      Case "decimal"
         GeneraPrefijoSQLVB = "sgl"
      Case "tinyint", "smallint"
         GeneraPrefijoSQLVB = "int"
      Case "int"
         GeneraPrefijoSQLVB = "lng"
      Case "money", "smallmoney"
         GeneraPrefijoSQLVB = "cur"
   End Select
End Function



Private Sub Class_Initialize()
   bolConPK = True
   strCodigoGenerado = ""
End Sub
