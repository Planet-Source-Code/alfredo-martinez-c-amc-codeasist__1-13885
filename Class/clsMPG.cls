VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Nombre:       clsMPG
'Descripcion:  Libreria de metodos yProcedimientos generales
'              Para desarrollo de sistemas
'Autor:        Alfredo Martinez C.
'Fecha:        13-Febrero-2000
'La libreria de compone de varios metodos
'de uso general en desarrollo de sistemas.
'Nota: Revisar documentacion anexa.

'Variables de propiedades
Private mstrFtoNumIns As String
Private mstrFtoFloatIns As String
Private mstrFtoDateIns As String




'Variables para los metodos de validacion de datos.
Public Mensaje As String
Public TituloMensaje As String

'-------------- Metodos para efectos especiales -------------------'
'Enumeraciones para los metodos de efectos especiales
'de visualizacion.
Public Enum Colores
    Negro = vbBlack
    Rojo = vbRed
    Verde = vbGreen
    Amarillo = vbYellow
    Azul = vbBlue
    Magenta = vbMagenta
    Cyan = vbCyan
    Blanco = vbWhite
End Enum

Public Enum TipoEmbosadoTexto
   Negro3D = 0
   Gris3D = 8
End Enum

Public Enum TipoDespliegeTexto
   Embossed = 15
   NoEmbossed = -15
End Enum

'Enumeraciones para los metodos de Seleccion de datos
'de visualizacion.
Public Enum TipoControl
    MaskEditBox = 1
    TextBox = 2
End Enum



Public Property Get FtoNumIns() As String
   FtoNumIns = mstrFtoNumIns
End Property

Public Property Let FtoNumIns(ByVal strNewValue As String)
   mstrFtoNumIns = strNewValue
End Property

Public Property Get FtoFloatIns() As String
   FtoFloatIns = mstrFtoFloatIns
End Property

Public Property Let FtoFloatIns(ByVal strNewValue As String)
   mstrFtoFloatIns = strNewValue
End Property

Public Property Get FtoDateIns() As String
   FtoDateIns = mstrFtoDateIns
End Property

Public Property Let FtoDateIns(ByVal strNewValue As String)
   mstrFtoDateIns = strNewValue
End Property




'-------------- Metodos para efectos especiales -------------------'
'Metodo para simular un efecto de explosion
'al presentar las ventanas.
Public Sub Explosion(ByVal hwnd As Long, ByVal Movimientos As Integer, ByVal Color As Colores)
Dim myRect As RECT
Dim formWidth%, formHeight%, i%, X%, Y%, cx%, cy%
Dim TheScreen As Long
Dim Brush As Long
Dim hdc As Long
On Error GoTo ErrorExplosion
    hdc = GetWindowRect((hwnd), myRect)
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(Color)
    For i = 1 To Movimientos
        cx = formWidth * (i / Movimientos)
        cy = formHeight * (i / Movimientos)
        X = myRect.Left + (formWidth - cx) / 2
        Y = myRect.Top + (formHeight - cy) / 2
        Rectangle TheScreen, X, Y, X + cx, Y + cy
    Next i
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
Exit Sub
ErrorExplosion:
    MsgBox "Error: [" & Err & "] " & Error, vbCritical, App.Title
    Exit Sub
End Sub


'Metodo para simular un efecto de Implosion
'al cerrar las ventanas.
Public Sub Implosion(ByVal hwnd As Long, ByVal Movimientos As Integer, ByVal Color As Colores)
Dim myRect As RECT
Dim formWidth%, formHeight%, i%, X%, Y%, cx%, cy%
Dim TheScreen As Long
Dim Brush As Long
Dim hdc As Long
On Error GoTo ErrorImplosion
    hdc = GetWindowRect((hwnd), myRect)
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(Color)
    For i = Movimientos To 1 Step -1
        cx = formWidth * (i / Movimientos)
        cy = formHeight * (i / Movimientos)
        X = myRect.Left + (formWidth - cx) / 2
        Y = myRect.Top + (formHeight - cy) / 2
        Rectangle TheScreen, X, Y, X + cx, Y + cy
    Next i
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
Exit Sub
ErrorImplosion:
    MsgBox "Error: [" & Err & "] " & Error, vbCritical, App.Title
    Exit Sub
End Sub

'Metodo para establecer un color gradiente
'El objeto puede ser una forma o un PictureBox
'Los valores de las variable van de 0 a 255 (Valores de RGB)
'El degradalo va de en escala de 64
Public Sub Gradiente(TheObject As Object, Redval&, Greenval&, Blueval&)
Dim Step%, Reps%, FillTop%, FillLeft%, FillRight%, FillBottom%, HColor$
On Error GoTo ErrorGradiente
    Step = (TheObject.Height / 63)
    FillLeft = 0
    FillRight = TheObject.Width
    FillBottom = FillTop + Step
    For Reps = 1 To 63
        TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(Redval, Greenval, Blueval), BF
        Redval = Redval - 4
        Greenval = Greenval - 4
        Blueval = Blueval - 4
        If Redval <= 0 Then Redval = 0
        If Greenval <= 0 Then Greenval = 0
        If Blueval <= 0 Then Blueval = 0
        FillTop = FillBottom
        FillBottom = FillTop + Step
    Next
Exit Sub
ErrorGradiente:
    MsgBox "Error: [" & Err & "] " & Error, vbCritical, App.Title
    Exit Sub
End Sub

'Metodo para centrar las formas en la pantalla
Public Sub CentrarForma(ByVal oForma As Object)
    oForma.Left = (Screen.Width - oForma.Width) / 2
    oForma.Top = (Screen.Height - oForma.Height) / 2
End Sub



'Metodo para realizar un efecto de embosado.
Public Sub Embosado(ByVal oForma As Object, ByVal Texto As String, ByVal X As Long, ByVal Y As Long, ByVal TipoDeEmbosado As TipoEmbosadoTexto, ByVal TipoDespliege As TipoDespliegeTexto)
   oForma.CurrentX = X
   oForma.CurrentY = Y
   oForma.ForeColor = QBColor(15)
   oForma.Print Texto
   oForma.CurrentX = X + TipoDespliege
   oForma.CurrentY = Y + TipoDespliege
   oForma.ForeColor = QBColor(TipoDeEmbosado)
   oForma.Print Texto
End Sub

'Metodo para centrar un control en una forma u otro control
Public Sub CentrarControl(ByVal oCtlSource As Object, ByVal oCtlDestination As Object)
    oCtlSource.Left = (oCtlSource.Width - oCtlDestination.Width) / 2  ' Center form horizontally.
    'oCtlSource.Top = (oCtlSource.Height - oCtlDestination.Height) / 2  ' Center form vertically.
End Sub


'-------------------- Metodos para manejo de controles -----------------'
'      Aplica para combos, ListBox, TextBox y MaskEditBox

' The AutoSelect routine "selects" the
' control's entire contents as if it were
' doubled-clicked.
Public Sub AutoSelect(ByVal SelObject As Object, ByVal Control As TipoControl)
    SelObject.SelStart = 0
    If Control = 1 Then
        SelObject.SelLength = Len(SelObject.FormattedText)
        Exit Sub
    End If
    If Control = 2 Then
        SelObject.SelLength = Len(SelObject.Text)
        Exit Sub
    End If
End Sub

'Posiciona elementos ListBoxes y ComboBoxes a partir de la descripción del
'elemento a buscar.
Public Sub ElementoCombo(ByVal ComboBox As Object, ByVal Cadena As String)
Dim nCiclo As Long
Dim nElementos As Long
On Error GoTo ErrorElementoCombo
   Screen.MousePointer = vbHourglass
    nElementos = ComboBox.ListCount - 1
    If nElementos >= 0 Then
        For nCiclo = 0 To nElementos
            'Compara el elemento listado con el elemento
            'seleccionado.
            If ComboBox.List(nCiclo) = Cadena Then
                ComboBox.ListIndex = nCiclo
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Next nCiclo
        ComboBox.ListIndex = -1
    End If
    Screen.MousePointer = vbDefault
Exit Sub
ErrorElementoCombo:
    MsgBox "Error: [" & Err & "] " & Error, vbExclamation, App.Title
    Screen.MousePointer = vbDefault
End Sub


'--------- Metodos para manejo de cadenas de datos ----------------'

'Traduce de numeros a letras
Public Function Traduce(ByVal Monto As Currency) As String
Dim NumeroSt As String
Dim Longitud As Integer
Dim Ciclo As Integer
Dim UDoC As Integer
Dim Digito As Integer
Dim Digito2 As Integer
Dim NumeroPuente As String
Dim Mensaje As String
Dim Caracter As String
Dim HayAlgo As Integer

    NumeroPuente = Str(Fix(Monto))
    Mensaje = ""
    HayAlgo = False
    Longitud = Len(NumeroPuente)
    For Ciclo = 1 To Longitud
        Caracter = Mid(NumeroPuente, Ciclo, 1)
        If (Caracter >= "0") And (Caracter <= "9") Then
            NumeroSt = NumeroSt + Caracter
        End If
    Next Ciclo
    Longitud = Len(NumeroSt)
    If (NumeroSt = "0") Or (Longitud = 0) Then
        Mensaje = "CERO"
    Else
    For Ciclo = 1 To Longitud
        UDoC = (Longitud - Ciclo) Mod 3
        Digito = Int(Mid(NumeroSt, Ciclo, 1))
        If Digito > 0 Then
            HayAlgo = True
        End If
        Select Case UDoC
        Case 0:
            Select Case Digito
                Case 1: Mensaje = Mensaje + "UN"
                Case 2: Mensaje = Mensaje + "DOS"
                Case 3: Mensaje = Mensaje + "TRES"
                Case 4: Mensaje = Mensaje + "CUATRO"
                Case 5: Mensaje = Mensaje + "CINCO"
                Case 6: Mensaje = Mensaje + "SEIS"
                Case 7: Mensaje = Mensaje + "SIETE"
                Case 8: Mensaje = Mensaje + "OCHO"
                Case 9: Mensaje = Mensaje + "NUEVE"
            End Select
        Case 1:
            Digito2 = Int(Mid(NumeroSt, Ciclo + 1, 1))
            Select Case Digito
                Case 1:
                    Ciclo = Ciclo + 1
                    Select Case Digito2
                        Case 0: Mensaje = Mensaje + "DIEZ "
                        Case 1: Mensaje = Mensaje + "ONCE "
                        Case 2: Mensaje = Mensaje + "DOCE "
                        Case 3: Mensaje = Mensaje + "TRECE "
                        Case 4: Mensaje = Mensaje + "CATORCE "
                        Case 5: Mensaje = Mensaje + "QUINCE "
                        Case 6: Mensaje = Mensaje + "DIECISEIS "
                        Case 7: Mensaje = Mensaje + "DIECISIETE "
                        Case 8: Mensaje = Mensaje + "DIECIOCHO "
                        Case 9: Mensaje = Mensaje + "DIECINUEVE "
                    End Select
                Case 2:
                    If Digito2 = 0 Then
                        Mensaje = Mensaje + "VEINTE"
                    Else
                        Mensaje = Mensaje + "VEINTI"
                    End If
                Case 3: Mensaje = Mensaje + "TREINTA"
                Case 4: Mensaje = Mensaje + "CUARENTA"
                Case 5: Mensaje = Mensaje + "CINCUENTA"
                Case 6: Mensaje = Mensaje + "SESENTA"
                Case 7: Mensaje = Mensaje + "SETENTA"
                Case 8: Mensaje = Mensaje + "OCHENTA"
                Case 9: Mensaje = Mensaje + "NOVENTA"
            End Select
            If (Digito > 2) Then
                If (Digito2 <> 0) Then
                    Mensaje = Mensaje + " Y "
                End If
            End If
        Case 2:
            Select Case Digito
                Case 1:
                    If (Mid(NumeroSt, Ciclo + 1, 1) = "0") And (Mid(NumeroSt, Ciclo + 2, 1) = "0") Then
                         Mensaje = Mensaje + "CIEN "
                    Else
                         Mensaje = Mensaje + "CIENTO "
                    End If
                Case 2: Mensaje = Mensaje + "DOSCIENTOS "
                Case 3: Mensaje = Mensaje + "TRESCIENTOS "
                Case 4: Mensaje = Mensaje + "CUATROCIENTOS "
                Case 5: Mensaje = Mensaje + "QUINIENTOS "
                Case 6: Mensaje = Mensaje + "SEISCIENTOS "
                Case 7: Mensaje = Mensaje + "SETECIENTOS "
                Case 8: Mensaje = Mensaje + "OCHOCIENTOS "
                Case 9: Mensaje = Mensaje + "NOVECIENTOS "
            End Select
        End Select
        Select Case (Longitud - Ciclo)
            Case 12:
                If HayAlgo Then
                    If Mensaje = "UN" Then
                        Mensaje = Mensaje + " BILLON "
                    Else
                        Mensaje = Mensaje + " BILLONES "
                    End If
                    HayAlgo = False
                End If
            Case 9:
                If HayAlgo Then
                    Mensaje = Mensaje + " MIL "
                End If
            Case 6:
                If HayAlgo Then
                    If Mensaje = "UN" Then
                        Mensaje = Mensaje + " MILLON "
                    Else
                        Mensaje = Mensaje + " MILLONES "
                    End If
                    HayAlgo = False
                End If
            Case 3:
                If HayAlgo Then
                    Mensaje = Mensaje + " MIL "
                    HayAlgo = False
                End If
        End Select
    Next Ciclo
    End If
    Traduce = Mensaje
End Function

'Convierte de minuscular a Mayusculas
Public Function Mayusculas(ByRef nKeyAscii As Integer)
On Error Resume Next
    nKeyAscii = Asc(UCase(Chr(nKeyAscii)))
End Function



'-------------- Metodos para Validaciones de datos -------------------'

'Requiere informacion dentro de TextBox o ComboBox en arreglo
Public Function CamposRequeridos(ParamArray strCampos() As Variant) As Boolean
Dim Objeto As Variant
Dim nCiclo As Byte, nCampos As Integer
    On Error GoTo ErrorCamposRequeridos
    Screen.MousePointer = vbHourglass
    nCampos = UBound(strCampos) + 1 'Número de Parámetros
    'Compara con los objetos y valida
    For nCiclo = 0 To nCampos - 1
        If strCampos(nCiclo).Visible Then                    'Solo los visibles
          If Len(Trim(strCampos(nCiclo).Text)) <= 0 Then     'Valida Longitud
            MsgBox Mensaje, vbInformation, TituloMensaje
            If strCampos(nCiclo).Enabled Then
                strCampos(nCiclo).SetFocus
                CamposRequeridos = False
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            Screen.MousePointer = vbDefault
            Exit Function
          End If
        End If
    Next nCiclo
    CamposRequeridos = True
    Screen.MousePointer = vbDefault
Exit Function
ErrorCamposRequeridos:
    MsgBox "Error: [" & Err & "] " & Error, vbExclamation, App.Title
    Screen.MousePointer = vbDefault
    Exit Function
End Function

'Simulacion de una function booleana
Public Function StringToInteger(ByVal StrValor As String) As Integer

    If UCase(StrValor) = "SI" Then
        StringToInteger = 1
    Else
        StringToInteger = 0
    End If

End Function

'Simulacion de una function booleana
Public Function IntegerToString(ByVal IntValor As Integer) As String
    If IntValor = 1 Then
        IntegerToString = "Si"
    Else
        IntegerToString = "No"
    End If
End Function

'Metodo para solo capturar datos de tipo Money
'en un TextBox o MaskEditBox.
Public Function SoloDinero(ByVal Objeto As Object, ByRef nKeyAscii As Integer) As Integer
Static nPuntos As Byte
    'Limpiar el objeto
    If Objeto.SelLength > 0 Then
        Select Case nKeyAscii
            Case 0 To 7, 11, 12 To 25, 27 To 44, 47, 58 To 255
                'nada
            Case Else
                Objeto.Text = ""
        End Select
    End If

    'Verifica que ya tenga el punto decimal
    If InStr(Objeto.Text, ".") > 0 And nKeyAscii = 46 Then
        'If nKeyAscii = 46 Then
            SoloDinero = 0
            Beep
            Exit Function
        'End If
    End If
    
    'Setermina si existe el signo negativo o no
    If Left(Objeto.Text, 1) = "-" Then
        nPuntos = 9
    Else
        nPuntos = 8
    End If
        
    'Pone el signo negativo en cualquier posicion.
    If Left(Objeto.Text, 1) <> "-" And nKeyAscii = 45 And Len(Objeto.Text) <= 12 Then
        SoloDinero = nKeyAscii
        Exit Function
    End If
        
    'Pone limite al control si los numeros enteros exceden de lo permitido
    'Solo acepta "." u otro caracter especial ( Enter,Backspace,etc.)
    If InStr(Objeto.Text, ".") = 0 And Len(Objeto.Text) > nPuntos And nKeyAscii <> 46 And nKeyAscii <> 13 And nKeyAscii <> 9 And nKeyAscii <> 10 And nKeyAscii <> 8 Then
        SoloDinero = 0
        Beep
        Exit Function
    End If
        
    'Delimitar numero de Caracteres
    Select Case Len(Objeto.Text)
        Case 0 To 12
            Select Case nKeyAscii
                Case 0 To 7, 11, 12 To 25, 27 To 45, 47, 58 To 255
                        SoloDinero = 0
                        Beep
                Case Else
                        SoloDinero = nKeyAscii
            End Select
        Case Is >= 13
            Select Case nKeyAscii
                Case 0 To 7, 11, 12 To 25, 27 To 255
                    SoloDinero = 0
                    Beep
                Case Else
                        SoloDinero = nKeyAscii
            End Select
    End Select
End Function

'Metodo para solo permitir la introduccion de una fecha
'en un TextBox o MaskEditBox.
Public Sub SoloFecha(ByVal Objeto As Object, ByRef nKeyAscii As Integer)
   If Len(Objeto.Text) > 10 Then
      nKeyAscii = 0
      Exit Sub
   End If
    
   Select Case Len(Objeto.Text)
      Case 0, 1, 3, 4, 6, 7
         Select Case nKeyAscii
            Case 0 To 7, 11, 12 To 47, 58 To 255
               nKeyAscii = 0
               Beep
         End Select
      Case 2, 5
         Select Case nKeyAscii
            Case 0 To 7, 11, 12 To 46, 48 To 255
               nKeyAscii = 0
               Beep
         End Select
   End Select
End Sub

'Metodo para solo permitir la entrada de numeros
'en un TextBox o MaskEditBox.
Public Sub SoloNumeros(ByRef nKeyAscii As Integer)
Dim strCaracter
    Select Case nKeyAscii
        Case 0 To 7, 11, 12 To 47, 58 To 255
        nKeyAscii = 0
        Beep
    End Select
End Sub

'Metodo para validar la longitud de un texto
Public Sub ValidaLongitud(ByVal Objeto As Object, ByVal nLongitud As Byte, ByRef KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Len(Objeto.Text) > nLongitud Then
        MsgBox "La longitud del campo es menor", vbInformation, App.Title
        KeyAscii = 0
        Objeto.SetFocus
        Exit Sub
    End If
End Sub

'Metodo que valida que los combos contengan algun valor
Public Function CombosRequeridos(ParamArray NCombos() As Variant) As Boolean
Dim Objeto As Variant, nCampos As Long
Dim nCiclo As Byte
On Error GoTo ErrorCombosRequeridos
   Screen.MousePointer = vbHourglass
   nCampos = UBound(NCombos) + 1 'Número de Parámetros
   'Compara con los objetos y valida
   For nCiclo = 0 To nCampos - 1
     If NCombos(nCiclo).ListIndex < 0 Then     'Valida Longitud
       MsgBox Mensaje, vbInformation, TituloMensaje
       
       If NCombos(nCiclo).Enabled Then
           NCombos(nCiclo).SetFocus
           CombosRequeridos = False
           Screen.MousePointer = vbDefault
           Exit Function
       End If
       Screen.MousePointer = vbDefault
       Exit Function
    End If
   Next nCiclo

   'cambie la posicion de esta variable.
   CombosRequeridos = True
   Screen.MousePointer = vbDefault
Exit Function
ErrorCombosRequeridos:
    MsgBox "Error: [" & Err & "] " & Error, vbExclamation, App.Title
    Screen.MousePointer = vbDefault
    Exit Function
End Function

'Metodo para quitar los apostrofos
Public Function SApostrofo(ByVal nKeyAscii As Integer) As Integer
   If nKeyAscii = 39 Then
      SApostrofo = 180
   Else
      SApostrofo = nKeyAscii
   End If
End Function

'Procedimiento para el tratamiento de los apostrofos
Public Function SQLQuote(ByVal strValue As String) As String
Dim strNewString As String
   strNewString = Replace(strValue, Chr(39), "''")
   SQLQuote = "'" & strNewString & "'"
End Function


'Temporizador para dar tiempo a las conexiones.
Public Function Tiempo(ByVal Segundos As String)
Dim cCadHora As String, dHFina As Variant
On Error Resume Next
   cCadHora = "00:00:" & Segundos
   dHFina = Now + TimeValue(cCadHora)
   Do While Now < dHFina
       DoEvents
   Loop
End Function

'Metodo para in item en el registro de configuraciones.
Public Function LeerItem(ByVal cAplicacion As String, ByVal cSeccion As String, _
       ByVal cKey As String, Optional ByVal cDefault As String) As String
Dim strResult As String
   
   'Verifica si se ha enviado un valor por default y consulta el registro
   If Trim$(cDefault) = "" Then
      strResult = GetSetting(appname:=cAplicacion, section:=cSeccion, _
                          Key:=cKey)
   Else
      strResult = GetSetting(appname:=cAplicacion, section:=cSeccion, _
                          Key:=cKey, Default:=cDefault)
   End If
   
   'Verifica el resultado de la busqueda
   If Trim$(strResult) = "" Then
      If Trim$(cDefault) = "" Then
         LeerItem = ""
      Else
         LeerItem = cDefault
      End If
   Else
      LeerItem = Trim$(strResult)
   End If
   
End Function



'Metodo para Borrar un item en el registro de configuraciones.
Public Function BorrarItem(ByVal cAplicacion As String, ByVal cSeccion As String) As Boolean
On Error GoTo ErrorBorrarItem
   DeleteSetting cAplicacion, cSeccion
   BorrarItem = True
Exit Function
ErrorBorrarItem:
   BorrarItem = False
End Function


'Metodo para Salvar in item en el registro de configuraciones.
Public Function NuevoItem(ByVal cAplicacion As String, ByVal cSeccion As String, _
        ByVal cKey As String, ByVal cValor As String) As Boolean
Dim strResult As String

   SaveSetting appname:=cAplicacion, section:=cSeccion, _
               Key:=cKey, setting:=cValor
   
   strResult = GetSetting(appname:=cAplicacion, section:=cSeccion, _
                       Key:=cKey)
   If Trim$(strResult) = "" Then
      NuevoItem = False
   Else
      NuevoItem = True
   End If

End Function




'Funcion para validar los tipos de parametros de datos
'y regresar un valor ya tratado.
Public Function VerifyValue(ByVal vValue As Variant) As String
   Select Case VarType(vValue)
      Case vbEmpty, vbNull
         VerifyValue = "Null"
      Case vbInteger, vbLong
         VerifyValue = Format(vValue, mstrFtoNumIns)
      Case vbSingle, vbDouble, vbCurrency
         VerifyValue = Format(vValue, mstrFtoFloatIns)
      Case vbDate
         VerifyValue = Format(vValue, mstrFtoDateIns)
      Case vbString
         VerifyValue = SQLQuote(vValue)
   End Select
End Function


'Procedimiento para el tratamiento de nombres de campos con espacios intermedios
Public Function BracketField(ByVal strValue As String) As String
   If InStr(strValue, " ") > 0 Then
      BracketField = "[" & strValue & "]"
   Else
      BracketField = strValue
   End If
End Function



'Funcion para leer un archivo ini
Public Function ReadINI(ByVal strFile As String, strKeySection, strKeyValue As String) As String
Dim lngResult As Long
Dim strResult As String * 200
   
   lngResult = GetPrivateProfileString(strKeySection, _
   strKeyValue, strFile, strResult, Len(strResult), _
   strFile)
   If lngResult = 0 Then
      'Ha ocurrido un error
      ReadINI = ""
   Else
      ReadINI = Mid$(strResult, 1, lngResult)
   End If
End Function


'Funcion para leer un archivo ini
Public Function WriteINI(ByVal strFile As String, strKeySection, strKeyValue As String, ByVal strValue As String) As Boolean
Dim lngResult As Long
   
   lngResult = WritePrivateProfileString(strKeySection, _
   strKeyValue, strValue, strFile)
   If lngResult = 0 Then
      'Ha ocurrido un error
      WriteINI = False
   Else
      WriteINI = True
   End If
End Function

