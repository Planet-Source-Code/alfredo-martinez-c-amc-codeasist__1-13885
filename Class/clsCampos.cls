VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colCampos As Collection
Private i As Long, I2 As Long


Public Sub Add(ByVal strNombre As String, ByVal strTipo As String, ByVal intLongitud As Integer, _
                             ByVal intPresicion As Integer, ByVal intEscala As Integer, Optional ByVal boolPK As Boolean)
Dim objCampo As New clsCampo
   
   With objCampo
      .Nombre = strNombre
      .Tipo = strTipo
      .Longitud = intLongitud
      .Presicion = intPresicion
      .Escala = intEscala
      .PK = boolPK
   End With
   
   With m_colCampos
      .Add objCampo
   End With
   Set objCampo = Nothing
End Sub


Public Sub Clear()
   With m_colCampos
      For i = 1 To .Count
         m_colCampos.Remove (1)
      Next i
   End With
End Sub

Public Function Item(ByVal Index As Long) As clsCampo
   Set Item = m_colCampos(Index)
End Function


Public Function Count() As Long
   Count = m_colCampos.Count
End Function


Private Sub Class_Initialize()
   Set m_colCampos = New Collection
End Sub


Private Sub Class_Terminate()
   Set m_colCampos = Nothing
End Sub

